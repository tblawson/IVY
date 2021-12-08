# -*- coding: utf-8 -*-
# run_page.py
""" Defines individual run page as a panel-like object,
for inclusion in a wx.Notebook object

PYTHON3 DEVELOPMENT VERSION

Created on Fri Mar 5 15:00:16 2021

@author: t.lawson
"""

import wx
from wx.lib.masked import NumCtrl
import logging
import time
import datetime as dt
import math
from scripts import acquisition as acq, devices, IVY_events as Evts

logger = logging.getLogger(__name__)


class RunPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.status = self.GetTopLevelParent().sb
        self.directory = self.GetTopLevelParent().directory
        self.version = self.GetTopLevelParent().version
        self.SetupPage = self.GetTopLevelParent().page1
        self.run_id = 'none'

        self.GAINS_CHOICE = ['1e3', '1e4', '1e5', '1e6',
                             '1e7', '1e8', '1e9', '1e10']
        self.Rs_CHOICE = ['1k', '10k', '100k', '1M', '10M', '100M', '1G']
        self.Rs_SWITCHABLE = self.Rs_CHOICE[: 4]
        self.Rs_VALUES = [1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9]
        self.Rs_choice_to_val = dict(zip(self.Rs_CHOICE, self.Rs_VALUES))
        self.VNODE_CHOICE = ['V1', 'V2', 'V3']

        # Event bindings
#        self.Bind(evts.EVT_UPDATE_COM_STR, self.UpdateComment)
        self.Bind(Evts.EVT_DATA, self.update_data)
#        self.Bind(evts.EVT_START_ROW, self.UpdateStartRow)

        self.RunThread = None

        # Comment widgets
        comment_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Comment:')
        self.comment = wx.TextCtrl(self, id=wx.ID_ANY, size=(600, 20))
#        self.Comment.Bind(wx.EVT_TEXT, self.OnComment)
        comtip = 'Use this field to add remarks and observations that may not'\
                 ' be recorded automatically.'
        # self.Comment.SetToolTipString(comtip)
        self.comment.SetToolTip(comtip)

        run_id_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Run ID:')
        self.new_run_id_btn = wx.Button(self, id=wx.ID_ANY,
                                        label='Create new run id')
        idcomtip = 'Create new id to uniquely identify this set of '\
                   'measurement data.'
        # self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.new_run_id_btn.SetToolTip(idcomtip)
        self.new_run_id_btn.Bind(wx.EVT_BUTTON, self.on_new_run_id)
        self.run_id_txtctrl = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

        # Run Setup widgets
        duc_gain_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                     style=wx.ALIGN_LEFT,
                                     label='DUC gain (V/A):')
        self.DUCgain_cb = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GAINS_CHOICE,
                                      style=wx.CB_DROPDOWN)
        rs_lbl = wx.StaticText(self, id=wx.ID_ANY,
                               style=wx.ALIGN_LEFT, label='I/P Rs:')
        self.Rs_cb = wx.ComboBox(self, wx.ID_ANY, choices=self.Rs_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.Rs_cb.Bind(wx.EVT_COMBOBOX, self.on_Rs)
        settle_del_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Settle delay:')
        self.settle_del_spinctrl = wx.SpinCtrl(self, id=wx.ID_ANY, value='1800',
                                               min=0, max=3600)
        src_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_LEFT,
                                label='V1 Setting:')
        self.V1_set_numctrl = NumCtrl(self, id=wx.ID_ANY, integerWidth=3,
                                      fractionWidth=8, groupDigits=True)
        self.V1_set_numctrl.Bind(wx.lib.masked.EVT_NUM, self.on_v1_set)
        zero_volts_btn = wx.Button(self, id=wx.ID_ANY, label='Set zero volts',
                                   size=(200, 20))
        zero_volts_btn.Bind(wx.EVT_BUTTON, self.on_zero_volts)

        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, style=wx.LI_HORIZONTAL)

        #  Run control and progress widgets
        self.start_btn = wx.Button(self, id=wx.ID_ANY, label='Start run')
        self.start_btn.Bind(wx.EVT_BUTTON, self.on_start)
        self.stop_btn = wx.Button(self, id=wx.ID_ANY, label='Abort run')
        self.stop_btn.Bind(wx.EVT_BUTTON, self.on_abort)
        self.stop_btn.Enable(False)
        node_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Node:')
        self.node_cb = wx.ComboBox(self, wx.ID_ANY, choices=self.VNODE_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.node_cb.Bind(wx.EVT_COMBOBOX, self.on_node)
        v_av_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Mean V:')
        self.Vav_numctrl = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                                   groupDigits=True)
        v_sd_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Stdev V:')
        self.Vsd_numctrl = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                                   groupDigits=True)
        time_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Timestamp:')
        self.time_txtctrl = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY,
                                        size=(200, 20))
        row_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                label='Current measurement:')
        self.Row_txtctrl = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        progress_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                     label='Run progress:')
        self.progress_gauge = wx.Gauge(self, id=wx.ID_ANY, range=100,
                                       name='Progress')

        gb_sizer = wx.GridBagSizer()

        # Comment widgets
        gb_sizer.Add(comment_lbl, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.comment, pos=(0, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(run_id_lbl, pos=(1, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.new_run_id_btn, pos=(3, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.run_id_txtctrl, pos=(1, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Run setup widgets
        gb_sizer.Add(duc_gain_lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DUCgain_cb, pos=(3, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(rs_lbl, pos=(2, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Rs_cb, pos=(3, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(settle_del_lbl, pos=(2, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.settle_del_spinctrl, pos=(3, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(src_lbl, pos=(2, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V1_set_numctrl, pos=(3, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(zero_volts_btn, pos=(3, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.h_sep1, pos=(4, 0), span=(1, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)

        #  Run control and progress widgets
        gb_sizer.Add(self.start_btn, pos=(5, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.stop_btn, pos=(6, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(node_lbl, pos=(5, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.node_cb, pos=(6, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_av_lbl, pos=(5, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vav_numctrl, pos=(6, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_sd_lbl, pos=(5, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vsd_numctrl, pos=(6, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(time_lbl, pos=(5, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.time_txtctrl, pos=(6, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(row_lbl, pos=(5, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Row_txtctrl, pos=(6, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(progress_lbl, pos=(7, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.progress_gauge, pos=(7, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gb_sizer)

        self.autocomstr = ''
        self.manstr = ''

        self.Rs_val = 0

        # Dictionary to hold ALL runs for this application session:
        self.master_run_dict = {}

    def on_new_run_id(self, e):
        self.version = self.GetTopLevelParent().version
        duc_name = self.SetupPage.DUCName.GetValue()
        gain = self.DUCgain_cb.GetValue()
        rs = self.Rs_cb.GetValue()
        timestamp = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.run_id = f'IVY.v{self.version} {duc_name} (Gain={gain}; Rs={rs}) {timestamp}'
        self.status.SetStatusText('Id for subsequent runs:', 0)
        self.status.SetStatusText(str(self.run_id), 1)
        self.run_id_txtctrl.SetValue(str(self.run_id))

    def update_data(self, e):
        # Triggered by an 'update data' event
        # event parameter is a dictionary:
        # ud{'node:,'Vm':,'Vsd':,'time':,'row':,'Prog':,'end_flag':[0,1]}
        if 'node' in e.ud:
            self.node_cb.SetValue(str(e.ud['node']))
        if 'Vm' in e.ud:
            self.Vav_numctrl.SetValue(str(e.ud['Vm']))
        if 'Vsd' in e.ud:
            self.Vsd_numctrl.SetValue(str(e.ud['Vsd']))
        if 'time' in e.ud:
            self.time_txtctrl.SetValue(str(e.ud['time']))
        if 'row' in e.ud:
            self.Row_txtctrl.SetValue(str(e.ud['row']))
        if 'progress' in e.ud:
            self.progress_gauge.SetValue(e.ud['progress'])
        if 'end_flag' in e.ud:  # Aborted or Finished
            self.RunThread = None
            self.start_btn.Enable(True)

    def on_Rs(self, e):
        self.Rs_val = self.Rs_choice_to_val[e.GetString()]  # an INT
        print(f'\nRunPage.OnRs(): Rs ={self.Rs_val}')
        logger.info(f'Rs = {self.Rs_val}')
        if e.GetString() in self.Rs_SWITCHABLE:  # a STRING
            s = str(int(math.log10(self.Rs_val)))  # '3','4','5' or '6'
            msg = f'\nSwitching Rs - Sending "{s}" to IVbox.'
            print(msg)
            logger.info(msg)
            devices.ROLES_INSTR['IVbox'].send_cmd(s)

    @staticmethod
    def on_node(e):
        node = e.GetString()  # 'V1', 'V2', or 'V3'
        print('\nRunPage.OnNode():', node)
        logger.info('\nRunPage.OnNode(): {}'.format(node))
        s = node[1]
        if s in ('1', '2'):
            print('\nRunPage.OnNode():Sending IVbox "{}".'.format(s))
            logger.info('\nRunPage.OnNode():Sending IVbox "{}"'.format(s))
            devices.ROLES_INSTR['IVbox'].send_cmd(s)
        else:  # '3'
            print('\nRunPage.OnNode():IGNORING IVbox cmd "{}".'.format(s))
            logger.info('IGNORING IVbox cmd "{}".'.format(s))

    @staticmethod
    def on_v1_set(e):
        # Called by change in value (manually OR by software!)
        v1 = e.GetValue()
        msg = 'V1 = {}'.format(v1)
        print('RunPage.OnV1Set(): ', msg)
        logger.info(msg)
        src = devices.ROLES_INSTR['SRC']
        src.set_v(v1)
        time.sleep(0.5)
        if v1 == 0:
            src.stby()
        else:
            src.oper()
        time.sleep(0.5)

    def on_zero_volts(self, e):
        # V1:
        src = devices.ROLES_INSTR['SRC']
        if self.V1_set_numctrl.GetValue() == 0:
            print('RunPage.OnZeroVolts(): Zero/Stby directly')
            logger.info('RunPage.OnZeroVolts(): Zero/Stby directly.')
            src.set_v(0)
            src.stby()
        else:
            self.V1_set_numctrl.SetValue('0')  # Calls OnV1Set() ONLY IF VAL CHANGES
            print('RunPage.OnZeroVolts():  Zero/Stby via V1 display')
            logger.info('RunPage.OnZeroVolts():  Zero/Stby via V1 display.')

    def on_start(self, e):
        self.progress_gauge.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('', 1)
        self.status.SetStatusText('Starting run', 0)
        if self.RunThread is None:
            self.stop_btn.Enable(True)  # Enable Stop button
            self.start_btn.Enable(False)  # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)

    def on_abort(self):
        self.start_btn.Enable(True)
        self.stop_btn.Enable(False)  # Disable Stop button
        self.RunThread._want_abort = 1  # .abort

