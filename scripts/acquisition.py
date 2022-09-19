# -*- coding: utf-8 -*-
"""
PYTHON 3 DEVELOPMENT VERSION

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
# import logging
import json

# import IVY_events as evts
from scripts import devices, IVY_events as evts
# import devices

VSETDELAY = 180  # 180s - Delay after applying new voltage.
AZERODELAY = 5  # 5 - Delay after AZERO.
NREADS = 20
TEST_V_OUT = [0.1, 1, 10]  # O/P test voltage selection
NODES = ['V1', 'V2']  # Input node selection
TEST_V_MASK = [0, -1, 1, 0]  # Polarity / zero selection
I_MIN = 1e-11  # 10 pA (S/N issues, noise limit)
I_MAX = 0.01  # 10 mA (Opamp overheating limit)
V1_MIN = 0.01  # 10 mV (S/N issues, noise limit)
V1_MAX = 10  # 10 V (Opamp output limit)
P_MAX = 480  # Maximum progress (20 measurement-cycles * 24 rows)

# logger = logging.getLogger(__name__)


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
        self.Comment = self.RunPage.comment.GetValue()

        self._want_abort = 0

        self.V12Data = {'V1': [], 'V2': []}
        self.V3Data = []
        self.Times = []
        self.V12m = {'V1': 0.0, 'V2': 0.0}
        self.V12sd = {'V1': 0.0, 'V2': 0.0}
        self.V3m = 0.0
        self.V3sd = 0.0
        self.T = 0.0
        self.Troom = 0.0
        self.Proom = 0.0
        self.RHroom = 0.0

        # Dictionary for accumulating data for THIS RUN:
        self.run_dict = {'Comment': self.Comment,
                         'Rs': self.RunPage.Rs_val,
                         'DUC_G': float(self.RunPage.DUCgain_cb.GetValue()),
                         'Settle_delay': self.RunPage.settle_del_spinctrl.GetValue(),
                         'Vset_delay': VSETDELAY,
                         'Azero_delay': AZERODELAY,
                         'DVM12_init': devices.ROLES_INSTR['DVM12'].InitStr,
                         'DVM3_init': devices.ROLES_INSTR['DVM3'].InitStr,
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

        print('Role -> Instrument:')
        print('------------------------------')
        # logger.info('Role -> Instrument:')
        # Print all device objects
        for r in devices.ROLES_INSTR.keys():
            if r == 'IVbox':
                d = 'IV_box'
            else:
                d = devices.ROLES_INSTR[r].descr
            sub_dict = {r: d}
            self.run_dict['Instruments'].update(sub_dict)
            print(f'{devices.INSTR_DATA[d]["role"]} \t-> {d}')
            # logger.info(f'{devices.INSTR_DATA[d]["role"]} \t-> {d}')

        self.settle_time = self.RunPage.settle_del_spinctrl.GetValue()

        # Local record of GMH ports and addresses
        self.GMH1Demo_status = devices.ROLES_INSTR['GMH'].demo
        self.GMH1Port = devices.ROLES_INSTR['GMH'].port

        self.Rs = 0.0
        self.duc_gain = 0.0
        self.v1_nom = 0.0
        self.i_nom = 0.0
        self.v_out = 0.0
        self.v1_set = 0.0
        self.input_range = 0.0
        self.output_range = 0.0
        self.T_dvm_op = 0.0
        self.tm = 0.0

        self.start()  # Starts the thread running on creation

    def run(self):
        """
        Run Worker Thread.
        This is where all the important stuff goes, in a repeated cycle.
        """
        print('\nRUN START...\n')
        # logger.info('RUN START...')

        # Set button availability
        self.RunPage.stop_btn.Enable(True)
        self.RunPage.start_btn.Enable(False)

        # Clear plots
        clr_plot_ev = evts.ClearPlotEvent()
        wx.PostEvent(self.PlotPage, clr_plot_ev)

        stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Waiting to settle...', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        time.sleep(self.settle_time)

        # Initialise all instruments (doesn't open GMH sensors yet)
        if self._want_abort:
            self.abort_run()
            return
        self.initialise()

        stat_ev = evts.StatusEvent(msg='',
                                   field='b')  # write to both status fields
        wx.PostEvent(self.TopLevel, stat_ev)

        if self._want_abort:
            self.abort_run()
            return
        stat_ev = evts.StatusEvent(msg=' Post-initialise delay (3s)', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        row = 1
        pbar = 0

        """
        The following sequence progresses through three blocks with nominal
        output voltages of 0.1, 1 or 10 V. For each output voltage block
        the input DVM switches between the two nodes V1 and V2.
        For each input node, a mask is applied to the output voltage (and
        thus input V) causing the value to be set to 0 V, each polarity,
        then 0 V again.

        Nominal input voltage and current are calculated based on nominal
        DUC gain and Rs. If the nominal input current would be outside the
        range {1e-11 A < I < 1e-3 A}, that 8-row block is skipped. If the
        nominal calculated input voltage, required to result in the prescribed
        output, would be outside the range {0.01 V < V < 10 V}, that 8-row
        block is skipped.
        """
        self.Rs = self.RunPage.Rs_val
        self.duc_gain = float(self.RunPage.DUCgain_cb.GetValue())

        for abs_V3 in TEST_V_OUT:  # Loop over desired output voltages
            self.v1_nom = self.Rs * abs_V3 / self.duc_gain  # Nominal non-zero input
            # print(f'\n________________|V3| loop_______________________'
            #       f'\n acquisition.py, L181: abs_V3 = {abs_V3}, v1_nom = {self.v1_nom}\n')
            # logger.info('V3: {} V'.format(abs_V3))

            self.i_nom = self.v1_nom / self.Rs
            if abs(self.i_nom) <= I_MIN or abs(self.i_nom) >= I_MAX:
                i_nom_str = '(%.1g A)' % self.i_nom
                warning = '\nNominal I/P test-I outside scope! ' + i_nom_str
                print(warning)
                # logger.warning(warning)
                stat_ev = evts.StatusEvent(msg=warning, field=1)
                wx.PostEvent(self.TopLevel, stat_ev)
                pbar += 160
                update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-',
                          'row': row, 'progress': 100.0*pbar/P_MAX,
                          'end_flag': 0}

                update_ev = evts.DataEvent(ud=update)
                wx.PostEvent(self.RunPage, update_ev)
                continue

            if abs(self.v1_nom) < V1_MIN or abs(self.v1_nom) > V1_MAX:
                v_nom_str = '(%.1g V)' % self.v1_nom
                warning = 'Nom. I/P test-V outside scope! '+v_nom_str
                print(f'\n{warning}')
                # logger.warning(warning)
                stat_ev = evts.StatusEvent(msg=warning, field=1)
                wx.PostEvent(self.TopLevel, stat_ev)
                pbar += 160
                update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-',
                          'row': row, 'progress': 100.0*pbar/P_MAX,
                          'end_flag': 0}
                update_ev = evts.DataEvent(ud=update)
                wx.PostEvent(self.RunPage, update_ev)
                continue

            update = {'node': '-', 'Vm': 0, 'Vsc': 0, 'time': '-', 'row': row,
                      'progress': 100.0*pbar/P_MAX, 'end_flag': 0}
            update_ev = evts.DataEvent(ud=update)
            wx.PostEvent(self.RunPage, update_ev)
            for node in NODES:  # Select input node (V1 then V2)
                self.set_node(node)
                # print('\n_\t______________V1,V2 loop___________________')
                # print(f'\n\tacquisition.py, L223: node = {node}, (abs_V3 = {abs_V3}, v1_nom = {self.v1_nom})\n')
                for V3_mask in TEST_V_MASK:
                    '''
                    Loop over {0,+,-,0} test voltages (assumes negative gain)
                    '''
                    # print('\n_\t_\t________________V3 mask loop_________________')
                    # print(f'\n\t\tacquisition.py, L229: V3 polarity mask = {V3_mask} '
                    #       f'(node = {node}, abs_V3 = {abs_V3}, v1_nom = {self.v1_nom})\n')
                    self.v_out = abs_V3 * V3_mask  # Nominal output
                    self.v1_set = -1.0 * self.v_out * self.Rs / self.duc_gain

                    if abs(self.v1_set) == 0:
                        self.v1_set = 0.0
                    # print('I/P test-V = %f \tO/P test-V = %f' % (self.v1_set, self.v_out))
                    # logger.info('I/P test-V = %f \tO/P test-V = %f',
                    #             self.v1_set, self.v_out)

                    stat_ev = evts.StatusEvent(msg='AqnThread.run():', field=0)
                    wx.PostEvent(self.TopLevel, stat_ev)
                    stat_ev = evts.StatusEvent(msg='I/P test-V = ' +
                                                   str(self.v1_set) +
                                               '. O/P test-V = ' +
                                                   str(self.v_out), field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    self.set_up_meas_this_row(node)  # Clear F5520A errors & data
                    if self._want_abort:
                        self.abort_run()
                        return
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(3)  # Wait 3s after checking F5520A error

                    '''
                    Set DVM ranges to suit voltages that they're
                    about to be exposed to and ensure they're RANGE-LOCKED:
                    '''
                    cmd12 = f'DCV {self.v1_nom}'
                    # print(f'DVM12 command = "{cmd12}" - RANGE-LOCKING DVM12 TO {self.v1_nom}...')
                    devices.ROLES_INSTR['DVM12'].send_cmd(cmd12)
                    # rng12 = float(devices.ROLES_INSTR['DVM12'].send_cmd('RANGE?'))
                    # print(f'(acqn.L263) SANITY CHECK: IP-range = {rng12}.')

                    # print(f'RANGE-LOCKING DVM3 TO {abs_V3}...')
                    devices.ROLES_INSTR['DVM3'].send_cmd(f'DCV {abs_V3}')
                    # rng3 = float(devices.ROLES_INSTR['DVM3'].send_cmd('RANGE?'))
                    # print(f'(acqn.L268) SANITY CHECK: OP-range={rng3}')

                    if not (devices.ROLES_INSTR['DVM12'].demo and devices.ROLES_INSTR['DVM3'].demo):
                        time.sleep(0.5)  # Settle after setting range

                    # print('Aqn_thread.run(): masked V1_set = {}'.format(self.v1_set))
                    # logger.info('masked V1_set = %f', self.v1_set)
                    self.RunPage.V1_set_numctrl.SetValue(str(self.v1_set))
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(0.5)  # Settle after setting V
                    if self.v1_set == 0:
                        devices.ROLES_INSTR['SRC'].oper()  # Over-ride 0V STBY
                    if self._want_abort:
                        self.abort_run()
                        return
                    if not devices.ROLES_INSTR['SRC'].demo:
                        time.sleep(VSETDELAY)  # wait after applying V and OPER mode.

                    # Prepare DVMs...
                    stat_ev = evts.StatusEvent(msg='Preparing DVMs...',
                                               field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    devices.ROLES_INSTR['DVM12'].send_cmd('LFREQ LINE')
                    devices.ROLES_INSTR['DVM3'].send_cmd('LFREQ LINE')
                    if self._want_abort:
                        self.abort_run()
                        return
                    if not (devices.ROLES_INSTR['DVM12'].demo and devices.ROLES_INSTR['DVM3'].demo):
                        time.sleep(3)

                    devices.ROLES_INSTR['DVM12'].send_cmd('AZERO ONCE')
                    devices.ROLES_INSTR['DVM3'].send_cmd('AZERO ON')
                    if self._want_abort:
                        self.abort_run()
                        return
                    if not (devices.ROLES_INSTR['DVM12'].demo and devices.ROLES_INSTR['DVM3'].demo):
                        time.sleep(AZERODELAY)  # 30

                    stat_msg = f'Making {NREADS} measurements each of {node} and V3 (V1_nom = {self.v1_nom} V)'
                    print(stat_msg)
                    # logger.info(stat_msg)
                    stat_ev = evts.StatusEvent(msg=stat_msg, field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    for n in range(NREADS):  # Acquire all V and t readings
                        self.measure_v(node)
                        self.measure_v('V3')
                        pbar += 1
                        update = {'node': '-', 'Vm': 0, 'Vsd': 0, 'time': '-',
                                  'row': row, 'progress': 100.0*pbar/P_MAX,
                                  'end_flag': 0}

                        update_ev = evts.DataEvent(ud=update)
                        wx.PostEvent(self.RunPage, update_ev)
                        if self._want_abort:
                            self.abort_run()
                            return
                    print('\n')
                    time.sleep(1)

                    msg = 'Number of {0:s} readings != {1:d}!'.format(node, NREADS)
                    assert len(self.V12Data[node]) == NREADS, msg

                    msg = 'Number of timestamps != {0:d}!'.format(NREADS)
                    assert len(self.Times) == NREADS, msg

                    tm_raw = dt.datetime.fromtimestamp(float(np.mean(self.Times)))
                    self.tm = tm_raw.strftime("%d/%m/%Y %H:%M:%S")  # self.tm
                    self.V12m[node] = float(np.mean(self.V12Data[node]))
                    self.V12sd[node] = float(np.std(self.V12Data[node], ddof=1))

                    msg = 'V12m[{0:s}] = {1:.6f}'.format(node, self.V12m[node])

                    print(msg)
                    # logger.info(msg)
                    self.set_node(node)
                    update = {'node': node, 'Vm': self.V12m[node],
                              'Vsd': self.V12sd[node], 'time': self.tm,
                              'row': row, 'progress': 100.0*pbar/P_MAX,
                              'end_flag': 0}

                    update_ev = evts.DataEvent(ud=update)
                    wx.PostEvent(self.RunPage, update_ev)
                    # input range should be fixed at abs. nom. value.:
                    if not devices.ROLES_INSTR['DVM12'].demo:
                        input_range = float(devices.ROLES_INSTR['DVM12'].send_cmd('RANGE?'))
                        assert input_range >= self.v1_nom, f'Input range wrong!: range={input_range}, V1_nom={self.v1_nom}'
                        if not isinstance(input_range, float):
                            input_range = 0.0
                        self.input_range = input_range

                    time.sleep(2)  # Give user time to read vals before update

                    msg = 'Number of V3 readings != {0:d}!'.format(NREADS)
                    print(msg)
                    # logger.info(msg)
                    assert len(self.V3Data) == NREADS, msg + '(got {} instead)'.format(len(self.V3Data))
                    self.V3m = float(np.mean(self.V3Data))
                    self.V3sd = float(np.std(self.V3Data, ddof=1))
                    self.T = devices.ROLES_INSTR['GMH'].measure('T')
                    if not devices.ROLES_INSTR['DVM3'].demo:
                        output_range = float(devices.ROLES_INSTR['DVM3'].send_cmd('RANGE?'))
                        assert output_range >= abs_V3, f'Output range wrong!: range={output_range}, V3_nom={abs_V3}'

                        # if not isinstance(output_range, float):
                        #     output_range = 0.0
                        self.output_range = output_range

                    self.set_node('V3')
                    update = {'node': 'V3', 'Vm': self.V3m, 'Vsd': self.V3sd,
                              'time': self.tm, 'row': row, 'end_flag': 0}
                    update_ev = evts.DataEvent(ud=update)
                    wx.PostEvent(self.RunPage, update_ev)

                    if self._want_abort:
                        self.abort_run()
                        return
                    stat_ev = evts.StatusEvent(msg="Post-acqisn. delay (5s)",
                                               field=1)
                    wx.PostEvent(self.TopLevel, stat_ev)

                    # Record room conditions
                    self.Troom = devices.ROLES_INSTR['GMHroom'].measure('T')
                    self.Proom = devices.ROLES_INSTR['GMHroom'].measure('P')
                    self.RHroom = devices.ROLES_INSTR['GMHroom'].measure('RH')

                    self.write_data_this_row(row, node)
                    self.plot_this_row(row, node)
                    time.sleep(0.1)
                    row += 1

                    # Reset start row for next measurement
                # (end of V3_mask loop)
            # (end of node loop)
        # (end of abs_V3 loop)

        self.finish_run()
        return

    def set_node(self, node):
        """
        Update Node ComboBox and Change I/P node relays in IV-box
        """
        # print('AqnThread.SetNode(): ', node)
        # logger.info('Node = %s', node)
        self.RunPage.node_cb.SetValue(node)  # Update widget value
        s = node[1]
        if s in ('1', '2'):
            msg = 'Sending IVbox "{:s}"'.format(s)
            # print('AqnThread.SetNode():', msg)
            # logger.info(msg)
            devices.ROLES_INSTR['IVbox'].send_cmd(s)
        else:  # '3'
            msg = 'IGNORING IVbox cmd "{:s}"'.format(s)
            # print('AqnThread.SetNode():', msg)
            # logger.info(msg)
        if not devices.ROLES_INSTR['IVbox'].demo:
            time.sleep(1)

    def plot_this_row(self, row, node):
        # Plot data
        dates = []
        for d in self.Times:
            dates.append(dt.datetime.fromtimestamp(d))
        clear_plot = 0

        if row == 1:
            clear_plot = 1  # start each run with a clear plot

        plot_ev = evts.PlotEvent(t=dates, V12=self.V12Data[node],
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
            if 'GMH' not in devices.ROLES_INSTR[r].descr:
                msg = 'Opening {:s}'.format(d)
                # print('AqnThread.initialise():', msg)
                # logger.info(msg)
                devices.ROLES_INSTR[r].open()
            else:
                msg = '{:s} already open'.format(d)
                # print('AqnThread.initialise():', msg)
                # logger.info(msg)

            stat_ev = evts.StatusEvent(msg=d, field=1)
            wx.PostEvent(self.TopLevel, stat_ev)
            devices.ROLES_INSTR[r].init()
            if not devices.ROLES_INSTR[r].demo:
                time.sleep(1)
        stat_ev = evts.StatusEvent(msg='Done', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)

    def set_up_meas_this_row(self, node):
        d = devices.ROLES_INSTR['SRC'].descr
        if 'F5520A' in d:
            err = devices.ROLES_INSTR['SRC'].check_err()  # 'ERR?','*CLS'
            msg = 'Cleared F5520A error: "{}"'.format(err)
            # print(msg)
            # logger.info(msg)

        del self.V12Data[node][:]
        del self.V3Data[:]
        del self.Times[:]

    def measure_v(self, node):
        assert node in ('V1', 'V2', 'V3'), 'Unknown argument to MeasureV().'

        if node == 'V1':
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvm_op = np.random.normal(self.v1_set, 1.0e-5 * abs(self.v1_set) + 1e-6)
                # self.V12Data['V1'].append(dvmOP)
                rtn = {'node': node, 'value': dvm_op, 'demo_data': True}
            else:
                dvm_op = float(devices.ROLES_INSTR['DVM12'].read())
                # V = float(dvmOP)  # float(filter(self.filt, dvmOP))
                # self.V12Data['V1'].append(V)
                rtn = {'node': node, 'value': dvm_op, 'demo_data': False}
            self.V12Data['V1'].append(dvm_op)
            return rtn

        elif node == 'V2':
            if devices.ROLES_INSTR['DVM12'].demo is True:
                dvm_op = np.random.normal(0.0, 1.0e-5 * abs(self.v1_set) + 1e-6)
                # self.V12Data['V2'].append(dvmOP)
                rtn = {'node': node, 'value': dvm_op, 'demo_data': True}
            else:
                dvm_op = float(devices.ROLES_INSTR['DVM12'].read())
                # V = float(filter(self.filt, dvmOP))
                # self.V12Data['V2'].append(V)
                rtn = {'node': node, 'value': dvm_op, 'demo_data': False}
            self.V12Data['V2'].append(dvm_op)
            return rtn

        elif node == 'V3':
            """
            Just want ONE set of 20 timestamps -
            could have been either V1 or V2 instead.
            """
            self.Times.append(time.time())
            if devices.ROLES_INSTR['DVM3'].demo is True:
                dvm_op = np.random.normal(self.v_out,
                                          1.0e-5 * abs(self.v_out) + 1e-6)
                # self.V3Data.append(dvm_op)
                rtn = {'node': node, 'value': dvm_op, 'demo_data': True}
            else:
                dvm_op = float(devices.ROLES_INSTR['DVM3'].read())
                # V = float(filter(self.filt, dvm_op))
                rtn = {'node': node, 'value': dvm_op, 'demo_data': False}
            self.V3Data.append(dvm_op)
            # print('dvmOP =', dvm_op)
            return rtn

        # if not (devices.ROLES_INSTR['DVM12'].demo and devices.ROLES_INSTR['DVM3'].demo):
        #     time.sleep(0.1)
        # return 1

    def write_data_this_row(self, row, node):
        stat_ev = evts.StatusEvent(msg='AqnThread.WriteDataThisRow():',
                                   field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='Row '+str(row), field=1)
        wx.PostEvent(self.TopLevel, stat_ev)

        if devices.ROLES_INSTR['DVMT'].demo is True:
            T_dvm_op = np.random.normal(108.0, 1.0e-2)
            self.T_dvm_op = T_dvm_op
        else:
            T_dvm_op = devices.ROLES_INSTR['DVMT'].read()  # .SendCmd('READ?')
            self.T_dvm_op = float(T_dvm_op) # float(filter(self.filt, T_dvm_op))

        # Update run_dict:
        self.run_dict['Node'].append(node)
        self.run_dict['Nom_Vout'].append(self.v_out)
        self.run_dict['Date_time'].append(self.tm)
        self.run_dict['IP_V']['val'].append(self.V12m[node])
        self.run_dict['IP_V']['sd'].append(self.V12sd[node])
        self.run_dict['IPrange'].append(self.input_range)
        self.run_dict['OP_V']['val'].append(self.V3m)
        self.run_dict['OP_V']['sd'].append(self.V3sd)
        self.run_dict['OPrange'].append(self.output_range)
        self.run_dict['Pt_DVM'].append(self.T_dvm_op)
        self.run_dict['T_GMH'].append(self.T)
        self.run_dict['Room_conds']['T'].append(self.Troom)
        self.run_dict['Room_conds']['P'].append(self.Proom)
        self.run_dict['Room_conds']['RH'].append(self.RHroom)

    def abort_run(self):
        # prematurely end run, prompted by regular checks of _want_abort flag
        self.standby()  # Set sources to 0V and leave system safe

        update = {'progress': 100.0, 'end_flag': 1}
        stop_ev = evts.DataEvent(ud=update)
        wx.PostEvent(self.RunPage, stop_ev)

        self.RunPage.start_btn.Enable(True)
        self.RunPage.stop_btn.Enable(False)
        stat_ev = evts.StatusEvent(msg='RUN ABORTED', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        print('\nRun aborted.')
        # logger.info('Run aborted.')

    def finish_run(self):
        # Run complete - leave system safe and final data-save

        run_id = str(self.RunPage.run_id_txtctrl.GetValue())
        msg = f'Adding run "{run_id}" to master run dict.'
        print(msg)
        # logger.info(msg)
        self.RunPage.master_run_dict.update({run_id: self.run_dict})

        data_file = self.TopLevel.data_file
        # correct_dir = self.TopLevel.directory
        # displayed_dir = self.SetupPage.WorkingDir.GetValue()
        # working_dir = os.path.dirname(data_file)
        # assert displayed_dir == working_dir, 'Working Directory display error!'
        # assert working_dir == correct_dir, 'Working Directory error!'
        with open(data_file, 'w') as IVY_out:
            msg = f'SAVING ALL RUN DATA to {data_file}'
            # logger.info(msg)
            json.dump(self.RunPage.master_run_dict, IVY_out)

        self.standby()  # Set sources to 0V and leave system safe

        update = {'progress': 100.0, 'end_flag': 1}
        stop_ev = evts.DataEvent(ud=update)
        wx.PostEvent(self.RunPage, stop_ev)

        msg = '_________RUN COMPLETED_________'
        stat_ev = evts.StatusEvent(msg=msg, field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        stat_ev = evts.StatusEvent(msg='', field=1)
        wx.PostEvent(self.TopLevel, stat_ev)
        # logger.info(f'\n{msg}\n')


    def standby(self):
        # Set sources to 0V and disable outputs
        self.RunPage.V1_set_numctrl.SetValue(str(0))
        devices.ROLES_INSTR['SRC'].stby()  # .SendCmd('R0=')
        self.RunPage.start_btn.Enable(True)
        self.RunPage.stop_btn.Enable(False)

    def abort(self):
        """abort worker thread."""
        # Method for use by main thread to signal an abort
        stat_ev = evts.StatusEvent(msg='abort(): Run aborted', field=0)
        wx.PostEvent(self.TopLevel, stat_ev)
        self._want_abort = 1
        self.standby()  # Set sources to 0V and leave system safe
        time.sleep(1)

    def filt(self, char):
        """
        A helper function to filter rubbish from DVM o/p unicode string
        and retain any number (as a string)
        """
        accept_str = u'-0.12345678eE9'
        return char in accept_str  # Returns 'True' or 'False'

"""--------------End of Thread class definition-------------------"""
