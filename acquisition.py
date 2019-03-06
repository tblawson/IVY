# -*- coding: utf-8 -*-
"""
Created on Wed Jun 24 09:36:42 2015

DEVELOPMENT VERSION

acquisition.py:
Thread class that executes processing.
Contains definitions for usual __init__() and run() methods
 AND an abort() method. The Run() method forms the core of the
 procedure - any changes to the way the measurements are taken
 should be made here, and within included subroutines.

@author: t.lawson
"""
import os
import wx
from threading import Thread
import datetime as dt
import time
import numpy as np
import logging
import json

import IVY_events as evts
import devices

NREADS = 20
TEST_V_OUT = [0.1, 1, 10]  # O/P test voltage selection
NODES = ['V1', 'V2']  # Input node selection
TEST_V_MASK = [0, -1, 1, 0]  # Polarity / zero selection
I_MIN = 1e-11  # 10 pA (S/N issues, noise limit)
I_MAX = 0.01  # 10 mA (Opamp overheating limit)
V1_MIN = 0.01  # 10 mV (S/N issues, noise limit)
V1_MAX = 10  # 10 V (Opamp output limit)
P_MAX = 480  # Maximum progress (20 measurement-cycles * 24 rows)

logger = logging.getLogger(__name__)


class AqnThread(Thread):
    """Acquisition Thread Class."""
    def __init__(self, parent):
        # This runs when an instance of the class is created
        Thread.__init__(self)
        self.RunPage = parent
        self.SetupPage = self.RunPage.GetParent().GetPage(0)
        self.PlotPage = self.RunPage.GetParent().GetPage(2)
        self.CalcPage = self.RunPage.GetParent().GetPage(3)
        self.TopLevel = self.RunPage.GetTopLevelParent()
        self.Comment = self.RunPage.Comment.GetValue()

        self._want_abort = 0

        self.V12Data = {'V1': [], 'V2': []}
        self.V3Data = []
        self.Times = []
        self.V12m = {'V1': 0, 'V2': 0}
        self.V12sd = {'V1': 0, 'V2': 0}

        # Dictionary for accumulating data for this run:
        self.run_dict = {'Comment': self.Comment,
                         'Rs': self.RunPage.Rs_val,
                         'DUC_G': float(self.RunPage.DUCgain.GetValue()),
                         'Instruments': {},
                         'Nreads': NREADS,
                         'Date_time': [],
                         'Node': [],
                         'Nom_Vout': [],
                         'OP_V': {'val': [], 'sd': []},
                         'OPrange': [],
                         'IP_V': {'val': [], 'sd': []},
                         'IPrange': [],
                         'T_GMH': [],
                         'Pt_DVM': [],
                         'Room_conds': {'T': [], 'P': [], 'RH': []},
                         }

        print'Role -> Instrument:'
        print'------------------------------'
        logger.info('Role -> Instrument:')
        # Print all device objects
        for r in devices.ROLES_INSTR.keys():
            if r == 'IVbox':
                d = 'IV_box'
            else:
                d = devices.ROLES_INSTR[r].Descr
            sub_dict = {r: d}
            self.run_dict['Instruments'].update(sub_dict)
            print'%s \t-> %s' % (devices.INSTR_DATA[d]['role'], d)
            logger.info('%s \t-> %s', devices.INSTR_DATA[d]['role'], d)

        self.settle_time = self.RunPage.SettleDel.GetValue()

        # Local record of GMH ports and addresses
        self.GMH1Demo_status = devices.ROLES_INSTR['GMH'].demo
        self.GMH1Port = devices.ROLES_INSTR['GMH'].addr

        self.start()  # Starts the thread running on creation

    def run(self):
        '''
        Run Worker Thread.
        This is where all the important stuff goes, in a repeated cycle.
        '''
        print'\nRUN START...\n'
        logger.info('RUN START...')

        # Set button availability
        self.RunPage.StopBtn.Enable(True)
        self.RunPage.StartBtn.Enable(False)

        # Clear plots
        clr_plot_ev = evts.ClearPlotEvent()
        wx.PostEvent(self.PlotPage, clr_plot_ev)

#        self.WriteHeadings()

        stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Waiting to settle...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        time.sleep(self.settle_time)

        # Initialise all instruments (doesn't open GMH sensors yet)
        if self._want_abort:
            self.AbortRun()
            return
        self.initialise()

        stat_ev = evts.StatusEvent(msg='',
                                   field='b')  # write to both status fields
        wx.PostEvent(self.TopLevel, stat_ev)

        if self._want_abort:
            self.AbortRun()
            return
        stat_ev = evts.StatusEvent(msg=' Post-initialise delay (3s)', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        row = 1
        pbar = 0

        '''
        The following sequence progresses through three blocks with nominal
        output voltages of 0.1, 1 or 10 V. For each output voltage block
        the input DVM switches between the two nodes V1 and V2.
        For each input node, a mask is applied to the output voltage (and
        thus input V) causing the value to be set to 0 V, each polarity,
        then 0 V again.

        Nominal input voltage and current are calculated based on nominal
        DUC gain and Rs. If the nominal input current would be outside the
        scope {1e-11 A < I < 1e-3 A}, that 8-row block is skipped. If the
        nominal calculated input voltage, required to result in the prescribed
        output, would be outside the scope {0.01 V < V < 10 V}, that 8-row
        block is skipped.
        '''
        self.Rs = self.RunPage.Rs_val
        self.DUC_G = float(self.RunPage.DUCgain.GetValue())

        for abs_V3 in TEST_V_OUT:  # Loop over desired output voltages
            print'\nV3:', abs_V3, 'V'
            logger.info('V3: %d V', abs_V3)
            self.V1_nom = self.Rs*abs_V3/self.DUC_G
            self.I_nom = self.V1_nom/self.Rs
            if (abs(self.I_nom) <= I_MIN or abs(self.I_nom) >= I_MAX):
                I_nom_str = '(%.1g A)' % self.I_nom
                warning = '\nNominal I/P test-I outside scope! '+I_nom_str
                print warning
                logger.warning(warning)
                stat_ev = evts.StatusEvent(msg=warning, field=1)
                wx.PostEvent(self.TopLevel, stat_ev)
                pbar += 160
                Update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-',
                          'row': row, 'progress': 100.0*pbar/P_MAX,
                          'end_flag': 0}

                update_ev = evts.DataEvent(ud=Update)
                wx.PostEvent(self.RunPage, update_ev)
                continue

            if abs(self.V1_nom) < V1_MIN or abs(self.V1_nom) > V1_MAX:
                V_nom_str = '(%.1g V)' % self.V1_nom
                warning = '\nNom. I/P test-V outside scope! '+V_nom_str
                print warning
                logger.warning(warning)
                stat_ev = evts.StatusEvent(msg=warning, field=1)
                wx.PostEvent(self.TopLevel, stat_ev)
                pbar += 160
                Update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-',
                          'row': row, 'progress': 100.0*pbar/P_MAX,
                          'end_flag': 0}
                update_ev = evts.DataEvent(ud=Update)
                wx.PostEvent(self.RunPage, update_ev)
                continue

            Update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-', 'row': row,
                      'progress': 100.0*pbar/P_MAX, 'end_flag': 0}
            update_ev = evts.DataEvent(ud=Update)
            wx.PostEvent(self.RunPage, update_ev)
            for node in NODES:  # Select input node (V1 then V2)
                self.SetNode(node)

                for V3_mask in TEST_V_MASK:
                    '''
                    Loop over {0,+,-,0} test voltages (assumes negative gain)
                    '''
                    self.Vout = abs_V3*V3_mask  # Nominal output
                    self.V1_set = -1.0*self.Vout*self.Rs/self.DUC_G
                    if abs(self.V1_set) == 0:
                        self.V1_set = 0.0
                    print'I/P test-V = %f \tO/P test-V = %f' % (self.V1_set,
                                                                self.Vout)
                    logger.info('I/P test-V = %f \tO/P test-V = %f',
                                self.V1_set, self.Vout)

                    stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
                    wx.PostEvent(self.TopLevel, stat_ev)
                    stat_ev = evts.StatusEvent(msg='I/P test-V = ' +
                                               str(self.V1_set) +
                                               '. O/P test-V = ' +
                                               str(self.Vout), field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    self.SetUpMeasThisRow(node)  # Clear F5520A errors & data
                    if self._want_abort:
                        self.AbortRun()
                        return
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(3)  # Wait 3s after checking F5520A error

                    '''
                    Set DVM ranges to suit voltages that they're
                    about to be exposed to:
                    '''
                    devices.ROLES_INSTR['DVM12'].SendCmd('DCV AUTO')
                    devices.ROLES_INSTR['DVM3'].SendCmd('DCV AUTO')
                    if not (devices.ROLES_INSTR['DVM12'].demo and
                            devices.ROLES_INSTR['DVM3'].demo):
                                time.sleep(0.5)  # Settle after setting range

                    print 'Aqn_thread.run(): masked V1_set =', self.V1_set
                    logger.info('masked V1_set = %f', self.V1_set)
                    self.RunPage.V1Setting.SetValue(str(self.V1_set))
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(0.5)  # Settle after setting V
                    if self.V1_set == 0:
                        devices.ROLES_INSTR['SRC'].Oper()  # Over-ride 0V STBY
                    if self._want_abort:
                        self.AbortRun()
                        return
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(30)  # wait after applying V and OPER mode.

                    # Prepare DVMs...
                    stat_ev = evts.StatusEvent(msg='Preparing DVMs...',
                                               field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    devices.ROLES_INSTR['DVM12'].SendCmd('LFREQ LINE')
                    devices.ROLES_INSTR['DVM3'].SendCmd('LFREQ LINE')
                    if self._want_abort:
                        self.AbortRun()
                        return
                    if not (devices.ROLES_INSTR['DVM12'].demo and
                            devices.ROLES_INSTR['DVM3'].demo):
                                time.sleep(3)

                    devices.ROLES_INSTR['DVM12'].SendCmd('AZERO ONCE')
                    devices.ROLES_INSTR['DVM3'].SendCmd('AZERO ON')
                    if self._want_abort:
                        self.AbortRun()
                        return
                    if not (devices.ROLES_INSTR['DVM12'].demo and
                            devices.ROLES_INSTR['DVM3'].demo):
                        time.sleep(30)  # 30

                    stat_msg = 'Making {0:d} measurements each of {1:s} and'
                    'V3 (V1_nom = {2:.2f} V)'.format(NREADS, node, self.V1_nom)
                    print stat_msg
                    logger.info(stat_msg)
                    stat_ev = evts.StatusEvent(msg=stat_msg, field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    for n in range(NREADS):  # Acquire all V and t readings
                        self.MeasureV(node)
                        self.MeasureV('V3')
                        pbar += 1
                        Update = {'node': '-', 'Vm': 0, 'Vsd': 0, 'time': '-',
                                  'row': row, 'progress': 100.0*pbar/(P_MAX),
                                  'end_flag': 0}

                        update_ev = evts.DataEvent(ud=Update)
                        wx.PostEvent(self.RunPage, update_ev)
                        if self._want_abort:
                            self.AbortRun()
                            return
                    print'\n'
                    time.sleep(1)

                    msg = 'Number of {0:s} readings != {1:d}!'.format(node,
                                                                      NREADS)
                    assert len(self.V12Data[node]) == NREADS, msg

                    msg = 'Number of timestamps != {0:d}!'.format(NREADS)
                    assert len(self.Times) == NREADS, msg

                    tm_raw = dt.datetime.fromtimestamp(np.mean(self.Times))
                    self.tm = tm_raw.strftime("%d/%m/%Y %H:%M:%S")
                    self.V12m[node] = np.mean(self.V12Data[node])
                    self.V12sd[node] = np.std(self.V12Data[node], ddof=1)

                    msg = 'V12m[{0:s}] = {1:.6f}'.format(node,
                                                         self.V12m[node])
                    print msg
                    logger.info(msg)
                    self.SetNode(node)
                    Update = {'node': node, 'Vm': self.V12m[node],
                              'Vsd': self.V12sd[node], 'time': self.tm,
                              'row': row, 'progress': 100.0*pbar/(P_MAX),
                              'end_flag': 0}

                    update_ev = evts.DataEvent(ud=Update)
                    wx.PostEvent(self.RunPage, update_ev)
                    IPrange = devices.ROLES_INSTR['DVM12'].SendCmd('RANGE?')
                    if not isinstance(IPrange, float):
                        IPrange = 0.0
                    self.IPrange = IPrange

                    time.sleep(2)  # Give user time to read vals before update

                    msg = 'Number of V3 readings != {0:d}!'.format(NREADS)
                    assert len(self.V3Data) == NREADS, msg
                    self.V3m = np.mean(self.V3Data)
                    self.V3sd = np.std(self.V3Data, ddof=1)
                    self.T = devices.ROLES_INSTR['GMH'].Measure('T')
                    OPrange = devices.ROLES_INSTR['DVM3'].SendCmd('RANGE?')
                    if not isinstance(OPrange, float):
                        OPrange = 0.0
                    self.OPrange = OPrange

                    self.SetNode('V3')
                    Update = {'node': 'V3', 'Vm': self.V3m, 'Vsd': self.V3sd,
                              'time': self.tm, 'row': row, 'end_flag': 0}
                    update_ev = evts.DataEvent(ud=Update)
                    wx.PostEvent(self.RunPage, update_ev)

                    if self._want_abort:
                        self.AbortRun()
                        return
                    stat_ev = evts.StatusEvent(msg="Post-acqisn. delay (5s)",
                                               field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    # Record room conditions
                    self.Troom = devices.ROLES_INSTR['GMHroom'].Measure('T')
                    self.Proom = devices.ROLES_INSTR['GMHroom'].Measure('P')
                    self.RHroom = devices.ROLES_INSTR['GMHroom'].Measure('RH')

                    self.WriteDataThisRow(row, node)
                    self.PlotThisRow(row, node)
                    time.sleep(0.1)
                    row += 1

                    # Reset start row for next measurement
                # (end of V3_mask loop)
            # (end of node loop)
        # (end of abs_V3 loop)

        self.FinishRun()
        return

    def SetNode(self, node):
        '''
        Update Node ComboBox and Change I/P node relays in IV-box
        '''
        print'AqnThread.SetNode(): ', node
        logger.info('Node = %s', node)
        self.RunPage.Node.SetValue(node)  # Update widget value
        s = node[1]
        if s in ('1', '2'):
            msg = 'Sending IVbox "{:s}"'.format(s)
            print'AqnThread.SetNode():', msg
            logger.info(msg)
            devices.ROLES_INSTR['IVbox'].SendCmd(s)
        else:  # '3'
            msg = 'IGNORING IVbox cmd "{:s}"'.format(s)
            print'AqnThread.SetNode():', msg
            logger.info(msg)
        if not devices.ROLES_INSTR['IVbox'].demo:
            time.sleep(1)

    def PlotThisRow(self, row, node):
        # Plot data
        Dates = []
        for d in self.Times:
            Dates.append(dt.datetime.fromtimestamp(d))
        clear_plot = 0

        if row == 1:
            clear_plot = 1  # start each run with a clear plot

        plot_ev = evts.PlotEvent(t=Dates, V12=self.V12Data[node],
                                 V3=self.V3Data, clear=clear_plot, node=node)
        wx.PostEvent(self.PlotPage, plot_ev)

    def initialise(self):
        stat_ev = evts.StatusEvent(msg='Initialising instruments...', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

        for r in devices.ROLES_INSTR.keys():
            if r == 'IVbox':
                d = 'IV_box'
            else:
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
            # Open non-GMH devices:
            if 'GMH' not in devices.ROLES_INSTR[r].Descr:
                msg = 'Opening {:s}'.format(d)
                print'AqnThread.initialise():', msg
                logger.info(msg)
                devices.ROLES_INSTR[r].Open()
            else:
                msg = '{:s} already open'.format(d)
                print'AqnThread.initialise():', msg
                logger.info(msg)

            stat_ev = evts.StatusEvent(msg=d, field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR[r].Init()
            if not devices.ROLES_INSTR[r].demo:
                time.sleep(1)
        stat_ev = evts.StatusEvent(msg='Done', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

    def SetUpMeasThisRow(self, node):
        d = devices.ROLES_INSTR['SRC'].Descr
        if 'F5520A' in d:
            err = devices.ROLES_INSTR['SRC'].CheckErr()  # 'ERR?','*CLS'
            msg = 'Cleared F5520A error: "{}"'.format(err)
            print msg
            logger.info(msg)

        del self.V12Data[node][:]
        del self.V3Data[:]
        del self.Times[:]

    def MeasureV(self, node):
        assert node in ('V1', 'V2', 'V3'), 'Unknown argument to MeasureV().'

        if node == 'V1':
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvmOP = np.random.normal(self.V1_set,
                                         1.0e-5*abs(self.V1_set)+1e-6)
                self.V12Data['V1'].append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVM12'].Read()
                V = float(filter(self.filt, dvmOP))
                self.V12Data['V1'].append(V)

        elif node == 'V2':
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvmOP = np.random.normal(0.0,
                                         1.0e-5*abs(self.V1_set)+1e-6)
                self.V12Data['V2'].append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVM12'].Read()
                V = float(filter(self.filt, dvmOP))
                self.V12Data['V2'].append(V)

        elif node == 'V3':
            '''
            Just want one set of 20 timestamps -
            could have been either V1 or V2 instead.
            '''
            self.Times.append(time.time())
            if devices.ROLES_INSTR['DVM3'].demo is True:
                dvmOP = np.random.normal(self.Vout,
                                         1.0e-5*abs(self.Vout)+1e-6)
                self.V3Data.append(dvmOP)
            else:
                dvmOP = devices.ROLES_INSTR['DVM3'].Read()
                V = float(filter(self.filt, dvmOP))
                print'dvmOP =', V
                self.V3Data.append(V)
        if not (devices.ROLES_INSTR['DVM12'].demo and
                devices.ROLES_INSTR['DVM3'].demo):
                    time.sleep(0.1)
        return 1

    def WriteDataThisRow(self, row, node):
        stat_ev = evts.StatusEvent(msg='AqnThread.WriteDataThisRow():',
                                   field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Row '+str(row), field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        if devices.ROLES_INSTR['DVMT'].demo is True:
            TdvmOP = np.random.normal(108.0, 1.0e-2)
            self.TdvmOP = TdvmOP
        else:
            TdvmOP = devices.ROLES_INSTR['DVMT'].Read()  # .SendCmd('READ?')
            self.TdvmOP = float(filter(self.filt, TdvmOP))

        # Update run_dict:
        self.run_dict['Node'].append(node)
        self.run_dict['Nom_Vout'].append(self.Vout)
        self.run_dict['Date_time'].append(self.tm)
        self.run_dict['IP_V']['val'].append(self.V12m[node])
        self.run_dict['IP_V']['sd'].append(self.V12sd[node])
        self.run_dict['IPrange'].append(self.IPrange)
        self.run_dict['OP_V']['val'].append(self.V3m)
        self.run_dict['OP_V']['sd'].append(self.V3sd)
        self.run_dict['OPrange'].append(self.OPrange)
        self.run_dict['Pt_DVM'].append(self.TdvmOP)
        self.run_dict['T_GMH'].append(self.T)
        self.run_dict['Room_conds']['T'].append(self.Troom)
        self.run_dict['Room_conds']['P'].append(self.Proom)
        self.run_dict['Room_conds']['RH'].append(self.RHroom)

    def AbortRun(self):
        # prematurely end run, prompted by regular checks of _want_abort flag
        self.Standby()  # Set sources to 0V and leave system safe

        Update = {'progress': 100.0, 'end_flag': 1}
        stop_ev = evts.DataEvent(ud=Update)
        wx.PostEvent(self.RunPage, stop_ev)

        self.RunPage.StartBtn.Enable(True)
        self.RunPage.StopBtn.Enable(False)
        print'\nRun aborted.'
        logger.info('Run aborted.')

    def FinishRun(self):
        # Run complete - leave system safe and final data-save

        RunID = str(self.RunPage.run_id)
        self.RunPage.master_run_dict.update({RunID: self.run_dict})

        data_file = self.TopLevel.data_file
        correct_dir = self.TopLevel.directory
        displayed_dir = self.SetupPage.WorkingDir.GetValue()
        working_dir = os.path.dirname(data_file)
        assert displayed_dir == correct_dir, 'Working Directory display error!'
        assert working_dir == correct_dir, 'Working Directory error!'
        with open(data_file, 'w') as IVY_out:
            msg = 'SAVING ALL RUN DATA to {:s}'.format(data_file)
            logger.info(msg)
            json.dump(self.RunPage.master_run_dict, IVY_out)

        self.Standby()  # Set sources to 0V and leave system safe

        Update = {'progress': 100.0, 'end_flag': 1}
        stop_ev = evts.DataEvent(ud=Update)
        wx.PostEvent(self.RunPage, stop_ev)

        stat_ev = evts.StatusEvent(msg='RUN COMPLETED', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        self.RunPage.StartBtn.Enable(True)
        self.RunPage.StopBtn.Enable(False)

    def Standby(self):
        # Set sources to 0V and disable outputs
        self.RunPage.V1Setting.SetValue(str(0))
        devices.ROLES_INSTR['SRC'].Stby  # .SendCmd('R0=')

    def abort(self):
        """abort worker thread."""
        # Method for use by main thread to signal an abort
        stat_ev = evts.StatusEvent(msg='abort(): Run aborted', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        self._want_abort = 1
        time.sleep(1)

    def filt(self, char):
        '''
        A helper function to filter rubbish from DVM o/p unicode string
        and retain any number (as a string)
        '''
        accept_str = u'-0.12345678eE9'
        return char in accept_str  # Returns 'True' or 'False'


"""--------------End of Thread class definition-------------------"""
