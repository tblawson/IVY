# -*- coding: utf-8 -*-
""" nbpages.py - Defines individual notebook pages as panel-like objects

DEVELOPMENT VERSION

Created on Tue Jun 30 10:10:16 2015

@author: t.lawson
"""

import os
import logging
import wx
from wx.lib.masked import NumCtrl
import datetime as dt
import time
import math
import json
import IVY_events as Evts
import acquisition as acq
import devices
import GTC
import matplotlib
matplotlib.use('WXAgg')  # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick

# from matplotlib.backends.backend_wx import NavigationToolbar2Wx
# from openpyxl import load_workbook, cell

matplotlib.rc('lines', linewidth=1, color='blue')

logger = logging.getLogger(__name__)

"""
------------------------
# Setup Page definition:
------------------------
"""


class SetupPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # Event bindings
        self.Bind(Evts.EVT_FILEPATH, self.update_dir)

        self.status = self.GetTopLevelParent().sb
        self.version = self.GetTopLevelParent().version

        self.SRC_COMBO_CHOICE = ['none']
        self.DVM_COMBO_CHOICE = ['none']
        self.GMH_COMBO_CHOICE = ['none']
        self.IVBOX_COMBO_CHOICE = {'node=V1': '1',
                                   'node=V2': '2',
                                   'Rs=10^3': '3',
                                   'Rs=10^4': '4',
                                   'Rs=10^5': '5',
                                   'Rs=10^6': '6'}  # '': None
        self.INSTRUMENT_CHOICE = {'SRC': 'SRC_F5520A',
                                  'DVM12': 'DVM_3458A:s/n518',
                                  'DVM3': 'DVM_3458A:s/n066',
                                  'DVMT': 'DVM_34401A:s/n976',
                                  'GMH': 'GMH:s/n627',
                                  'GMHroom': 'GMH:s/n367'}
        self.T_SENSOR_CHOICE = devices.T_Sensors
        self.cbox_addr_COM = []
        self.cbox_addr_GPIB = []
        self.cbox_instr_SRC = []
        self.cbox_instr_DVM = []
        self.cbox_instr_GMH = []

        self.GMH1Addr = self.GMH2Addr = 0  # invalid initial address as default

        self.ResourceList = []
        self.ComList = []
        self.GPIBList = []
        self.GPIBAddressList = ['addresses', 'GPIB0::0']  # dummy values
        self.COMAddressList = ['addresses', 'COM0']  # dummy values initially.

        self.test_btns = []  # list of test buttons

        # Instruments
        self.SrcLbl = wx.StaticText(self, label='V1 source (SRC):',
                                    id=wx.ID_ANY)
        self.Sources = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.SRC_COMBO_CHOICE,
                                   size=(150, 10), style=wx.CB_DROPDOWN)
        self.Sources.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_SRC.append(self.Sources)

        self.IP_DVM_Lbl = wx.StaticText(self, label='Input DVM (DVM12):',
                                        id=wx.ID_ANY)
        self.IP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.IP_Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.IP_Dvms)
        self.OP_DVM_Lbl = wx.StaticText(self, label='Output DVM (DVM3):',
                                        id=wx.ID_ANY)
        self.OP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.OP_Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.OP_Dvms)
        self.TDvmLbl = wx.StaticText(self, label='T-probe DVM (DVMT):',
                                     id=wx.ID_ANY)
        self.TDvms = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.DVM_COMBO_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.TDvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.TDvms)

        self.GMHLbl = wx.StaticText(self, label='GMH probe (GMH):',
                                    id=wx.ID_ANY)
        self.GMHProbes = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GMH_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.GMHProbes.Bind(wx.EVT_COMBOBOX, self.build_comm_str)
        self.cbox_instr_GMH.append(self.GMHProbes)

        self.GMHroomLbl = wx.StaticText(self,
                                        label='Room conds. GMH probe (GMHroom):',
                                        id=wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.GMH_COMBO_CHOICE,
                                         style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)

        self.IVboxLbl = wx.StaticText(self, label='IV_box (IVbox) setting:',
                                      id=wx.ID_ANY)
        self.IVbox = wx.ComboBox(self, wx.ID_ANY,
                                 choices=list(self.IVBOX_COMBO_CHOICE.keys()),
                                 style=wx.CB_DROPDOWN)

        # Addresses
        self.SrcAddr = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.GPIBAddressList,
                                   size=(150, 10), style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.SrcAddr)
        self.SrcAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.IP_DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GPIBAddressList,
                                      style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.IP_DvmAddr)
        self.IP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.OP_DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GPIBAddressList,
                                      style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.OP_DvmAddr)
        self.OP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.TDvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                    choices=self.GPIBAddressList,
                                    style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.TDvmAddr)
        self.TDvmAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.GMHPorts = wx.ComboBox(self, wx.ID_ANY,
                                    choices=self.COMAddressList,
                                    style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHPorts)
        self.GMHPorts.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.GMHroomPorts = wx.ComboBox(self, wx.ID_ANY,
                                        choices=self.COMAddressList,
                                        style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHroomPorts)
        self.GMHroomPorts.Bind(wx.EVT_COMBOBOX, self.update_addr)

        self.IVboxAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.COMAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.IVboxAddr)
        self.IVboxAddr.Bind(wx.EVT_COMBOBOX, self.update_addr)

        # Filename
        dir_lbl = wx.StaticText(self, label='Working directory:',
                                id=wx.ID_ANY)
        self.WorkingDir = wx.TextCtrl(self, id=wx.ID_ANY,
                                      value=self.GetTopLevelParent().directory,
                                      style=wx.TE_READONLY)

        # DUC
        self.DUCName = wx.TextCtrl(self, id=wx.ID_ANY, value='DUC Name')
        self.DUCName.Bind(wx.EVT_TEXT, self.build_comm_str)

        # Autopopulate btn
        self.AutoPop = wx.Button(self, id=wx.ID_ANY, label='AutoPopulate')
        self.AutoPop.Bind(wx.EVT_BUTTON, self.on_auto_pop)

        # Test buttons
        self.VisaList = wx.Button(self, id=wx.ID_ANY, label='List Visa res')
        self.VisaList.Bind(wx.EVT_BUTTON, self.on_visa_list)
        self.ResList = wx.TextCtrl(self, id=wx.ID_ANY,
                                   value='Available Visa resources',
                                   style=wx.TE_READONLY | wx.TE_MULTILINE)

        self.STest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.STest.Bind(wx.EVT_BUTTON, self.on_test)

        self.D12Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.D12Test.Bind(wx.EVT_BUTTON, self.on_test)

        self.D3Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.D3Test.Bind(wx.EVT_BUTTON, self.on_test)

        self.DTTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.DTTest.Bind(wx.EVT_BUTTON, self.on_test)

        self.GMHTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMHTest.Bind(wx.EVT_BUTTON, self.on_test)

        self.GMHroomTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMHroomTest.Bind(wx.EVT_BUTTON, self.on_test)

        self.IVboxTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.IVboxTest.Bind(wx.EVT_BUTTON, self.on_IV_box_test)

        response_lbl = wx.StaticText(self,
                                     label='Instrument Test Response:',
                                     id=wx.ID_ANY)
        self.Response = wx.TextCtrl(self, id=wx.ID_ANY, value='',
                                    style=wx.TE_READONLY)

        gb_sizer = wx.GridBagSizer()

        # Instruments
        gb_sizer.Add(self.SrcLbl, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Sources, pos=(0, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IP_DVM_Lbl, pos=(1, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IP_Dvms, pos=(1, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.OP_DVM_Lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.OP_Dvms, pos=(2, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.TDvmLbl, pos=(3, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.TDvms, pos=(3, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHLbl, pos=(4, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHProbes, pos=(4, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomLbl, pos=(5, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomProbes, pos=(5, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IVboxLbl, pos=(6, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IVbox, pos=(6, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Addresses
        gb_sizer.Add(self.SrcAddr, pos=(0, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IP_DvmAddr, pos=(1, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.OP_DvmAddr, pos=(2, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.TDvmAddr, pos=(3, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHPorts, pos=(4, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomPorts, pos=(5, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IVboxAddr, pos=(6, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # DUC Name
        gb_sizer.Add(self.DUCName, pos=(6, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Filename
        gb_sizer.Add(dir_lbl, pos=(8, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.WorkingDir, pos=(8, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Test buttons
        gb_sizer.Add(self.STest, pos=(0, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.D12Test, pos=(1, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.D3Test, pos=(2, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DTTest, pos=(3, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHTest, pos=(4, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomTest, pos=(5, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IVboxTest, pos=(6, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        gb_sizer.Add(response_lbl, pos=(3, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Response, pos=(4, 4), span=(1, 3),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.VisaList, pos=(0, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.ResList, pos=(0, 4), span=(3, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Autopopulate btn
        gb_sizer.Add(self.AutoPop, pos=(2, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gb_sizer)

        # Roles and corresponding comboboxes/test btns are associated here:
        devices.ROLES_WIDGETS = {'SRC': {'lbl': self.SrcLbl,
                                         'icb': self.Sources,
                                         'acb': self.SrcAddr,
                                         'tbtn': self.STest}}
        devices.ROLES_WIDGETS.update({'DVM12': {'lbl': self.IP_DVM_Lbl,
                                                'icb': self.IP_Dvms,
                                                'acb': self.IP_DvmAddr,
                                                'tbtn': self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVM3': {'lbl': self.OP_DVM_Lbl,
                                               'icb': self.OP_Dvms,
                                               'acb': self.OP_DvmAddr,
                                               'tbtn': self.D3Test}})
        devices.ROLES_WIDGETS.update({'DVMT': {'lbl': self.TDvmLbl,
                                               'icb': self.TDvms,
                                               'acb': self.TDvmAddr,
                                               'tbtn': self.DTTest}})
        devices.ROLES_WIDGETS.update({'GMH': {'lbl': self.GMHLbl,
                                              'icb': self.GMHProbes,
                                              'acb': self.GMHPorts,
                                              'tbtn': self.GMHTest}})
        devices.ROLES_WIDGETS.update({'GMHroom': {'lbl': self.GMHroomLbl,
                                                  'icb': self.GMHroomProbes,
                                                  'acb': self.GMHroomPorts,
                                                  'tbtn': self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'IVbox': {'lbl': self.IVboxLbl,
                                                'icb': self.IVbox,
                                                'acb': self.IVboxAddr,
                                                'tbtn': self.IVboxTest}})

        # Create IV-box instrument once:
        d = 'IV_box'  # Description
        r = 'IVbox'  # Role
        self.create_instr(d, r)

        self.build_combo_choices()

    def build_combo_choices(self):
        for d in devices.INSTR_DATA.keys():
            if 'SRC' in d:
                self.SRC_COMBO_CHOICE.append(d)
            elif 'DVM' in d:
                self.DVM_COMBO_CHOICE.append(d)
            elif 'GMH' in d:
                self.GMH_COMBO_CHOICE.append(d)

        # Re-build combobox choices from list of SRC's
        for cbox in self.cbox_instr_SRC:
            current_val = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.SRC_COMBO_CHOICE)
            cbox.SetValue(current_val)

        # Re-build combobox choices from list of DVM's
        for cbox in self.cbox_instr_DVM:
            current_val = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.DVM_COMBO_CHOICE)
            cbox.SetValue(current_val)

        # Re-build combobox choices from list of GMH's
        for cbox in self.cbox_instr_GMH:
            current_val = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.GMH_COMBO_CHOICE)
            cbox.SetValue(current_val)

        # No choices for IV box - there's only one

    def update_dir(self, e):
        """
        Display working directory once selected.
        """
        self.WorkingDir.SetValue(e.Dir)

    def on_auto_pop(self, e):
        """
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        """
        for r in self.INSTRUMENT_CHOICE.keys():
            d = self.INSTRUMENT_CHOICE[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d)  # Update instr. cbox
            self.create_instr(d, r)
        if self.DUCName.GetValue() == u'DUC Name':
            self.DUCName.SetForegroundColour((255, 0, 0))  # red
            self.DUCName.SetValue('CHANGE_THIS!')

    def update_instr(self, e):
        """
        An instrument was selected for a role.
        Find description d and role r, then pass to CreatInstr()
        """
        r = ''
        d = e.GetString()
        for r in devices.ROLES_WIDGETS.keys():  # Cycle through roles
            if devices.ROLES_WIDGETS[r]['icb'] == e.GetEventObject():
                break  # stop looking on finding the right instr & role
        self.create_instr(d, r)

    def create_instr(self, d, r):
        """
         Called by both OnAutoPop() and UpdateInstr().
         Create each instrument in software & open visa session
         (GPIB and IVbox only).
         For GMH instruments, use GMH dll, not visa.
        """
        msg_head = 'CreateInstr(): {}'
        print('\nCreateInstr({},{})...'.format(d, r))
        logger.info('\nCreateInstr({0:s},{1:s})...'.format(d, r))
        if 'GMH' in r:  # Changed from d to r
            # create and open a GMH instrument instance
            msg = 'Creating GMH device ({0:s} -> {1:s})'.format(d, r)
            print('\n', msg_head.format(msg))
            logger.info(msg_head.format(msg))
            devices.ROLES_INSTR.update({r: devices.GMHSensor(d, r)})
        else:
            # create a visa instrument instance
            msg = 'Creating VISA device ({0:s} -> {1:s}).'.format(d, r)
            print('\n', msg_head.format(msg))
            logger.info(msg_head.format(msg))
            devices.ROLES_INSTR.update({r: devices.Instrument(d, r)})
            devices.ROLES_INSTR[r].open()
        self.set_instr(d, r)

    @staticmethod
    def set_instr(d, r):
        """
        Update internal info (INSTR_DATA), updates the addresses
        and Enables/disables testbuttons as necessary.
        Called by CreateInstr().
        """
        msg_head = 'SetInstr(): {}'
        assert d in devices.INSTR_DATA, 'Unknown instrument: %s' % d
        assert_msg = 'Unknown parameter ("role") for %s: .' % d
        assert 'role' in devices.INSTR_DATA[d], assert_msg
        devices.INSTR_DATA[d]['role'] = r  # update default role

        # Set the address cb to correct value (refer to devices.INSTR_DATA)
        a_cb = devices.ROLES_WIDGETS[r]['acb']
        msg = 'Address = {}'.format(devices.INSTR_DATA[d]['str_addr'])
        print(msg_head.format(msg))
        logger.info(msg_head.format(msg))
        a_cb.SetValue((devices.INSTR_DATA[d]['str_addr']))
        if d == 'none':
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(False)
        else:
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(True)

    def update_addr(self, e):
        # An address was manually selected
        # 1st, we'll need instrument description d...
        msg_head = 'UpdateAddr(): {}'
        d = 'none'
        r = ''
        addr = 0
        acb = e.GetEventObject()  # 'a'ddress 'c'ombo 'b'ox
        for r in devices.ROLES_WIDGETS.keys():
            if devices.ROLES_WIDGETS[r]['acb'] == acb:
                if r == 'IVbox':
                    d = 'IV_box'
                else:
                    d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when we've found the instr descr

        # ...Now change INSTR_DATA...
        a = e.GetString()  # address string, eg 'COM5' or 'GPIB0::23'
        # Ignore dummy values, like 'NO_ADDRESS':
        if (a not in self.GPIBAddressList) or (a not in self.COMAddressList):
            devices.INSTR_DATA[d]['str_addr'] = a
            devices.ROLES_INSTR[r].str_addr = a
            addr = int(a.lstrip('COMGPIB0:'))  # leave only numeric part
            devices.INSTR_DATA[d]['addr'] = addr
            devices.ROLES_INSTR[r].addr = addr
        msg = '{0:s} using {1:s} set to '\
              'addr {2:d} ({3:s})'.format(r, d, addr, a)
        print(msg_head.format(msg))
        logger.info(msg_head.format(msg))
        self.create_instr(d, r)

    def on_test(self, e):
        # Called when a 'test' button is clicked
        msg_head = 'nbpages.SetupPage.OnTest(): {}'
        r = ''
        d = 'none'
        for r in devices.ROLES_WIDGETS.keys():  # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when we've found right instr descr
        print('\n', msg_head.format(d))
        logger.info(msg_head.format(d))
        assert_msg = '{} has no "test" parameter'.format(d)
        assert 'test' in devices.INSTR_DATA[d], assert_msg
        test = devices.INSTR_DATA[d]['test']  # test string
        print('\tTest string:', test)
        logger.info('Test string: {}'.format(test))
        self.Response.SetValue(str(devices.ROLES_INSTR[r].test(test)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d, test), 0)

    def on_IV_box_test(self):
        # NOTE: config is the configuration description NOT the test string:
        config = devices.ROLES_WIDGETS['IVbox']['icb'].GetValue()
        test = self.IVBOX_COMBO_CHOICE[config]
        if test is not None:
            try:
                devices.ROLES_INSTR['IVbox'].test(test)
                self.Response.SetValue(config)
            except devices.visa.VisaIOError:
                self.Response.SetValue('IV_box test failed!')
        else:
            self.Response.SetValue('IV_box: empty test!')

    def build_comm_str(self, e):
        # Called by a change in GMH probe selection, or DUC name
        self.version = self.GetTopLevelParent().version
        d = e.GetString()
        r = ''
        if 'GMH' in d:  # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.create_instr(d, r)
        else:  # DUC name has been set or changed
            if d in ('CHANGE_THIS!', 'DUC Name'):  # DUC not yet specified!
                self.DUCName.SetForegroundColour((255, 0, 0))  # red
            else:
                self.DUCName.SetForegroundColour((0, 127, 0))  # green
#        RunPage = self.GetParent().GetPage(1)
#        params = {'DUC': self.DUCName.GetValue(),
#                  'GMH': self.GMHProbes.GetValue()}
#        commstr = 'IVY v.{0:s}. DUC: {1:s} monitored by '\
#                  '{2:s}'.format(self.version, params['DUC'], params['GMH'])
#        evt = evts.UpdateCommentEvent(str=commstr)
#        wx.PostEvent(RunPage, evt)

    def on_visa_list(self, e):
        res_list = devices.RM.list_resources()
        del self.ResourceList[:]  # list of COM ports ('COM X') & GPIB addr's
        del self.ComList[:]  # list of COM ports (numbers only)
        del self.GPIBList[:]  # list of GPIB addresses (numbers only)
        for item in res_list:
            self.ResourceList.append(item.replace('ASRL', 'COM'))
        for item in self.ResourceList:
            addr = item.replace('::INSTR', '')
            if 'COM' in item:
                self.ComList.append(addr)
            elif 'GPIB' in item:
                self.GPIBList.append(addr)

        # Re-build combobox choices from list of COM ports
        for cbox in self.cbox_addr_COM:
            current_port = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.ComList)
            cbox.SetValue(current_port)

        # Re-build combobox choices from list of GPIB addresses
        for cbox in self.cbox_addr_GPIB:
            current_port = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.GPIBList)
            cbox.SetValue(current_port)

        # Add resources to ResList TextCtrl widget
        res_addr_list = '\n'.join(self.ResourceList)
        self.ResList.SetValue(res_addr_list)


'''
____________________________________________
#-------------- End of Setup Page -----------
____________________________________________
'''
'''
----------------------
# Run Page definition:
----------------------
'''


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
        self.Comment = wx.TextCtrl(self, id=wx.ID_ANY, size=(600, 20))
#        self.Comment.Bind(wx.EVT_TEXT, self.OnComment)
        comtip = 'Use this field to add remarks and observations that may not'\
                 ' be recorded automatically.'
        # self.Comment.SetToolTipString(comtip)
        self.Comment.SetToolTip(comtip)

        self.NewRunIDBtn = wx.Button(self, id=wx.ID_ANY,
                                     label='Create new run id')
        idcomtip = 'Create new id to uniquely identify this set of '\
                   'measurement data.'
        # self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.NewRunIDBtn.SetToolTip(idcomtip)
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.on_new_run_id)
        self.RunID = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

        # Run Setup widgets
        duc_gain_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                     style=wx.ALIGN_LEFT,
                                     label='DUC gain (V/A):')
        self.DUCgain = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.GAINS_CHOICE,
                                   style=wx.CB_DROPDOWN)
        rs_lbl = wx.StaticText(self, id=wx.ID_ANY,
                               style=wx.ALIGN_LEFT, label='I/P Rs:')
        self.Rs = wx.ComboBox(self, wx.ID_ANY, choices=self.Rs_CHOICE,
                              style=wx.CB_DROPDOWN)
        self.Rs.Bind(wx.EVT_COMBOBOX, self.on_Rs)
        settle_del_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Settle delay:')
        self.SettleDel = wx.SpinCtrl(self, id=wx.ID_ANY, value='0',
                                     min=0, max=3600)
        src_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_LEFT,
                                label='V1 Setting:')
        self.V1Setting = NumCtrl(self, id=wx.ID_ANY, integerWidth=3,
                                 fractionWidth=8, groupDigits=True)
        self.V1Setting.Bind(wx.lib.masked.EVT_NUM, self.on_v1_set)
        zero_volts_btn = wx.Button(self, id=wx.ID_ANY, label='Set zero volts',
                                   size=(200, 20))
        zero_volts_btn.Bind(wx.EVT_BUTTON, self.on_zero_volts)

        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, style=wx.LI_HORIZONTAL)

        #  Run control and progress widgets
        self.StartBtn = wx.Button(self, id=wx.ID_ANY, label='Start run')
        self.StartBtn.Bind(wx.EVT_BUTTON, self.on_start)
        self.StopBtn = wx.Button(self, id=wx.ID_ANY, label='Abort run')
        self.StopBtn.Bind(wx.EVT_BUTTON, self.on_abort)
        self.StopBtn.Enable(False)
        node_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Node:')
        self.Node = wx.ComboBox(self, wx.ID_ANY, choices=self.VNODE_CHOICE,
                                style=wx.CB_DROPDOWN)
        self.Node.Bind(wx.EVT_COMBOBOX, self.on_node)
        v_av_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Mean V:')
        self.Vav = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)
        v_sd_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Stdev V:')
        self.Vsd = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)
        time_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Timestamp:')
        self.Time = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY,
                                size=(200, 20))
        row_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                label='Current measurement:')
        self.Row = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        progress_lbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                     label='Run progress:')
        self.Progress = wx.Gauge(self, id=wx.ID_ANY, range=100,
                                 name='Progress')

        gb_sizer = wx.GridBagSizer()

        # Comment widgets
        gb_sizer.Add(comment_lbl, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Comment, pos=(0, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.NewRunIDBtn, pos=(1, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.RunID, pos=(1, 1), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Run setup widgets
        gb_sizer.Add(duc_gain_lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.DUCgain, pos=(3, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(rs_lbl, pos=(2, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Rs, pos=(3, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(settle_del_lbl, pos=(2, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.SettleDel, pos=(3, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(src_lbl, pos=(2, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.V1Setting, pos=(3, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(zero_volts_btn, pos=(3, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.h_sep1, pos=(4, 0), span=(1, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)

        #  Run control and progress widgets
        gb_sizer.Add(self.StartBtn, pos=(5, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.StopBtn, pos=(6, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(node_lbl, pos=(5, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Node, pos=(6, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_av_lbl, pos=(5, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vav, pos=(6, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(v_sd_lbl, pos=(5, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Vsd, pos=(6, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(time_lbl, pos=(5, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Time, pos=(6, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(row_lbl, pos=(5, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Row, pos=(6, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(progress_lbl, pos=(7, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Progress, pos=(7, 1), span=(1, 5),
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
        gain = self.DUCgain.GetValue()
        rs = self.Rs.GetValue()
        timestamp = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.run_id = str('IVY.v{} {} (Gain={}; Rs={}) {}'.format(self.version,
                                                                  duc_name,
                                                                  gain, rs,
                                                                  timestamp))
        self.status.SetStatusText('Id for subsequent runs:', 0)
        self.status.SetStatusText(str(self.run_id), 1)
        self.RunID.SetValue(str(self.run_id))

    def update_data(self, e):
        # Triggered by an 'update data' event
        # event parameter is a dictionary:
        # ud{'node:,'Vm':,'Vsd':,'time':,'row':,'Prog':,'end_flag':[0,1]}
        if 'node' in e.ud:
            self.Node.SetValue(str(e.ud['node']))
        if 'Vm' in e.ud:
            self.Vav.SetValue(str(e.ud['Vm']))
        if 'Vsd' in e.ud:
            self.Vsd.SetValue(str(e.ud['Vsd']))
        if 'time' in e.ud:
            self.Time.SetValue(str(e.ud['time']))
        if 'row' in e.ud:
            self.Row.SetValue(str(e.ud['row']))
        if 'progress' in e.ud:
            self.Progress.SetValue(e.ud['progress'])
        if 'end_flag' in e.ud:  # Aborted or Finished
            self.RunThread = None
            self.StartBtn.Enable(True)

    def on_Rs(self, e):
        self.Rs_val = self.Rs_choice_to_val[e.GetString()]  # an INT
        print('\nRunPage.OnRs(): Rs =', self.Rs_val)
        logger.info('Rs = {}'.format(self.Rs_val))
        if e.GetString() in self.Rs_SWITCHABLE:  # a STRING
            s = str(int(math.log10(self.Rs_val)))  # '3','4','5' or '6'
            print('\nSwitching Rs - Sending "{}" to IVbox.'.format(s))
            logger.info('Switching Rs - Sending "{}" to IVbox.'.format(s))
            devices.ROLES_INSTR['IVbox'].send_cmd(s)

    def on_node(self, e):
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

    def on_v1_set(self, e):
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
        if self.V1Setting.GetValue() == 0:
            print('RunPage.OnZeroVolts(): Zero/Stby directly')
            logger.info('RunPage.OnZeroVolts(): Zero/Stby directly.')
            src.set_v(0)
            src.stby()
        else:
            self.V1Setting.SetValue('0')  # Calls OnV1Set() ONLY IF VAL CHANGES
            print('RunPage.OnZeroVolts():  Zero/Stby via V1 display')
            logger.info('RunPage.OnZeroVolts():  Zero/Stby via V1 display.')

    def on_start(self, e):
        self.Progress.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('', 1)
        self.status.SetStatusText('Starting run', 0)
        if self.RunThread is None:
            self.StopBtn.Enable(True)  # Enable Stop button
            self.StartBtn.Enable(False)  # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)

    def on_abort(self):
        self.StartBtn.Enable(True)
        self.StopBtn.Enable(False)  # Disable Stop button
        self.RunThread._want_abort = 1  # .abort


'''
__________________________________________
#-------------- End of Run Page ----------
__________________________________________
'''
'''
-----------------------
# Plot Page definition:
-----------------------
'''


class PlotPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.Bind(Evts.EVT_PLOT, self.update_plot)
        self.Bind(Evts.EVT_CLEARPLOT, self.clear_plot)

        self.figure = Figure()

        self.figure.subplots_adjust(hspace=0.3)  # 0.3" v-space b/tw subplots

        # 3high x 1wide, 3rd plot down:
        self.V3ax = self.figure.add_subplot(3, 1, 3)
        # Auto offset to centre on data:
        self.V3ax.ticklabel_format(style='sci', useOffset=False, axis='y',
                                   scilimits=(2, -2))

        # Scientific notation:
        fmt = mtick.ScalarFormatter(useMathText=True, useOffset=False)
        self.V3ax.yaxis.set_major_formatter(fmt)
        # Autoscale with 'buffer' around data extents:
        self.V3ax.autoscale(enable=True, axis='y',
                            tight=False)
        self.V3ax.set_xlabel('time')
        self.V3ax.set_ylabel('V3')

        # 3high x 1wide, 1st plot down:
        self.V1ax = self.figure.add_subplot(3, 1, 1, sharex=self.V3ax)
        self.V1ax.ticklabel_format(useOffset=False,
                                   axis='y')  # Auto o/set to centre on data
        self.V1ax.autoscale(enable=True,
                            axis='y',
                            tight=False)  # Autoscale with data 'buffer'
        plt.setp(self.V1ax.get_xticklabels(),
                 visible=False)  # Hide x-axis labels
        self.V1ax.set_ylabel('V1')
        self.V1ax.set_ylim(auto=True)
        v1_y_ost = self.V1ax.get_xaxis().get_offset_text()
        v1_y_ost.set_visible(False)

        # 3high x 1wide, 2nd plot down:
        self.V2ax = self.figure.add_subplot(3, 1, 2, sharex=self.V3ax)
        self.V2ax.ticklabel_format(useOffset=False,
                                   axis='y')  # Auto offset to centre on data

        # Autoscale with 'buffer' around data extents
        self.V2ax.autoscale(enable=True, axis='y', tight=False)
        plt.setp(self.V2ax.get_xticklabels(),
                 visible=False)  # Hide x-axis labels
        self.V2ax.set_ylabel('V2')
        self.V2ax.set_ylim(auto=True)
        v2_y_ost = self.V2ax.get_xaxis().get_offset_text()
        v2_y_ost.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)

    def update_plot(self, e):
        msg_head = 'PlotPage.UpdatePlot(): {}'
        print(msg_head.format('len(t)={}'.format(len(e.t))))
        logger.info(msg_head.format('len(t)={}'.format(len(e.t))))
        print(msg_head.format('{} len(V1)={}', 'len(V3)={}'.format(e.node, len(e.V12), len(e.V3))))
        logger.info(msg_head.format('{} len(V1)={}', 'len(V3)={}'.format(e.node, len(e.V12), len(e.V3))))
        if e.node == 'V1':
            self.V1ax.plot_date(e.t, e.V12, 'bo')
        else:  # V2 data
            self.V2ax.plot_date(e.t, e.V12, 'go')
        self.V3ax.plot_date(e.t, e.V3, 'ro')
        self.figure.autofmt_xdate()  # default settings
        self.V3ax.fmt_xdata = mdates.DateFormatter('%d-%m-%Y, %H:%M:%S')
        self.canvas.draw()
        self.canvas.Refresh()

    def clear_plot(self, e):
        self.V1ax.cla()
        self.V2ax.cla()
        self.V3ax.cla()
        self.V1ax.set_ylabel('V1')
        self.V2ax.set_ylabel('V2')
        self.V3ax.set_ylabel('V3')
        self.canvas.draw()
        self.canvas.Refresh()


'''
__________________________________________
-------------- End of Plot Page ----------
__________________________________________
'''

'''
---------------------------
# Analysis Page definition:
---------------------------
'''
DELTA = u'\N{GREEK CAPITAL LETTER DELTA}'


class CalcPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.version = self.GetTopLevelParent().version
#        self.data_file = self.GetTopLevelParent().data_file
#        self.results_file = os.path.join(os.path.dirname(self.data_file),
#                                         'IVY_Results.json')
        self.Results = {}  # Accumulate run-analyses here

        self.Rs_VALUES = self.GetTopLevelParent().page2.Rs_VALUES
        self.Rs_NAMES = ['IV1k 1k', 'IV10k 10k',
                         'IV100k 100k', 'IV1M 1M', 'IV10M 10M',
                         'IV100M 100M', 'IV1G 1G']
        self.Rs_VAL_NAME = dict(zip(self.Rs_VALUES, self.Rs_NAMES))

        gb_sizer = wx.GridBagSizer()

        # Analysis set-up:
        self.ListRuns = wx.Button(self, id=wx.ID_ANY, label='List run IDs')
        self.ListRuns.Bind(wx.EVT_BUTTON, self.on_list_runs)
        gb_sizer.Add(self.ListRuns, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.RunID = wx.ComboBox(self, id=wx.ID_ANY,
                                 style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.RunID.Bind(wx.EVT_COMBOBOX, self.on_run_choice)
        self.RunID.Bind(wx.EVT_TEXT, self.on_run_choice)
        gb_sizer.Add(self.RunID, pos=(0, 1), span=(1, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.Analyze = wx.Button(self, id=wx.ID_ANY, label='Analyze')
        self.Analyze.Bind(wx.EVT_BUTTON, self.on_analyze)
        gb_sizer.Add(self.Analyze, pos=(0, 7), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)  #

        # -----------------------------------------------------------------
        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, size=(720, 1),
                                    style=wx.LI_HORIZONTAL)
        gb_sizer.Add(self.h_sep1, pos=(1, 0), span=(1, 8),
                     flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------

        # Run summary:
        run_info_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Run Summary:')
        gb_sizer.Add(run_info_lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.RunInfo = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_MULTILINE |
                                   wx.TE_READONLY | wx.HSCROLL,
                                   size=(250, 1))
        gb_sizer.Add(self.RunInfo, pos=(3, 0), span=(20, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Analysis results:
        # Headings & environment readings
        quantity_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Quantity')
        gb_sizer.Add(quantity_lbl, pos=(2, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        ureal_summary_lbl = wx.StaticText(self, id=wx.ID_ANY, size=(200, 1),
                                          label='Val, Unc, DoF')
        gb_sizer.Add(ureal_summary_lbl, pos=(2, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        k_lbl = wx.StaticText(self, id=wx.ID_ANY, label='k(95%)')
        gb_sizer.Add(k_lbl, pos=(2, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        exp_u_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Exp Uncert')
        gb_sizer.Add(exp_u_lbl, pos=(2, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        p_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Pressure:')
        gb_sizer.Add(p_lbl, pos=(3, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.PSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.PSummary, pos=(3, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.Pk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.Pk, pos=(3, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.PExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.PExpU, pos=(3, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        rh_lbl = wx.StaticText(self, id=wx.ID_ANY, label='%RH:')
        gb_sizer.Add(rh_lbl, pos=(4, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHSummary, pos=(4, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHk, pos=(4, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHExpU, pos=(4, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        tgmh_lbl = wx.StaticText(self, id=wx.ID_ANY, label='T (GMH):')
        gb_sizer.Add(tgmh_lbl, pos=(5, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHSummary = wx.TextCtrl(self, id=wx.ID_ANY,
                                       style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHSummary, pos=(5, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHk, pos=(5, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHExpU, pos=(5, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND
        # -----------------------------------------------------------------
        self.h_sep3 = wx.StaticLine(self, id=wx.ID_ANY, size=(480, 1),
                                    style=wx.LI_HORIZONTAL)
        gb_sizer.Add(self.h_sep3, pos=(6, 2), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------
        vout_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                 label='Nom. ' + DELTA + 'Vout:')
        gb_sizer.Add(vout_lbl, pos=(7, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.NomVout = wx.ComboBox(self, id=wx.ID_ANY,
                                   style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.NomVout.Bind(wx.EVT_COMBOBOX, self.on_nom_vout_choice)
        gb_sizer.Add(self.NomVout, pos=(7, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        i_in_lbl = wx.StaticText(self, id=wx.ID_ANY, label=DELTA + 'I_in:')
        gb_sizer.Add(i_in_lbl, pos=(8, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.IinSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.IinSummary, pos=(8, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.Iink = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.Iink, pos=(8, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.IinExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.IinExpU, pos=(8, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        self.Budget = wx.TextCtrl(self, style=wx.TE_MULTILINE |
                                  wx.TE_READONLY | wx.HSCROLL, id=wx.ID_ANY,)
        budget_font = wx.Font(8, wx.MODERN, wx.NORMAL, wx.NORMAL)
        self.Budget.SetFont(budget_font)
        gb_sizer.Add(self.Budget, pos=(9, 2), span=(14, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)  #

        self.SetSizerAndFit(gb_sizer)

        self.data_file = None
        self.run_data = None
        self.run_IDs = []
        self.runstr = ''
        self.run_ID = ''
        self.ThisResult = {}
        self.nom_Vout = {}
        self.budget_table_sorted = {'pos': [], 'neg': []}
        self.result_row = []
        self.results_file = ''

    def on_list_runs(self, e):
        """
        Open the .json data file and de-serialize into the dictionary
        self.run_data for later access. Update the choices in the 'Run ID'
        widget to the run ids (used as primary keys in self.run_data).
        """
        self.data_file = self.GetTopLevelParent().data_file
        with open(self.data_file, 'r') as in_file:
            self.run_data = json.load(in_file)
            self.run_IDs = list(self.run_data.keys())

        self.RunID.Clear()
        self.RunID.AppendItems(self.run_IDs)
        self.RunID.SetSelection(0)

    def on_run_choice(self, e):
        id = e.GetString()
        self.runstr = json.dumps(self.run_data[id], indent=4)
        self.RunInfo.SetValue(self.runstr)
        self.PSummary.Clear()
        self.Pk.Clear()
        self.PExpU.Clear()
        self.RHSummary.Clear()
        self.RHk.Clear()
        self.RHExpU.Clear()
        self.TGMHSummary.Clear()
        self.TGMHk.Clear()
        self.TGMHExpU.Clear()
        self.IinSummary.Clear()
        self.Iink.Clear()
        self.IinExpU.Clear()
        self.Budget.Clear()

    def on_analyze(self, e):
        self.run_ID = self.RunID.GetValue()
        this_run = self.run_data[self.run_ID]

        logger.info('STARTING ANALYSIS...')

        # Correction for Pt-100 sensor DVM:
        dvmt = this_run['Instruments']['DVMT']
        dvmt_cor = self.build_ureal(devices.INSTR_DATA[dvmt]['correction_100r'])

        """
        Pt sensor is a few cm away from input resistors, so assume a
        fairly large type B Tdef of 0.1 deg C:
        """
        pt_t_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                             3, label='Pt_T_def')

        pt_alpha = self.build_ureal(devices.RES_DATA['Pt 100r']['alpha'])
        pt_beta = self.build_ureal(devices.RES_DATA['Pt 100r']['beta'])
        pt_r0 = self.build_ureal(devices.RES_DATA['Pt 100r']['R0_LV'])
        pt_t_ref = self.build_ureal(devices.RES_DATA['Pt 100r']['TRef_LV'])

        """
        GMH sensor is a few cm away from DUC which, itself, has a size of
        several cm, so assume a fairly large type B Tdef of 0.1 deg C:
        """
        gmh_t_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                              3, label='GMH_T_def')

        comment = this_run['Comment']
        duc_name = self.get_duc_name_from_run_id(self.run_ID)
        duc_gain = this_run['DUC_G']
        mean_date = self.get_mean_date()
        processed_date = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        print('Comment:', comment)
        print('Run_Id:', self.run_ID)
        print('gain ={}'.format(duc_gain))
        print('Mean_date:', mean_date)
        logger.info('Comment: %s\nRun_ID: %s\ngain = %s\nMean_date: %s',
                    comment, self.run_ID, duc_gain, mean_date)

        # Determine mean env. conditions
        gmh_temps = []
        gmh_room_rhs = []
        gmh_room_ps = []
        for T in this_run['T_GMH']:
            gmh_temps.append(T[0])
        for RH in this_run['Room_conds']['RH']:
            gmh_room_rhs.append(RH[0])
        for P in this_run['Room_conds']['P']:
            gmh_room_ps.append(P[0])

        d = this_run['Instruments']['GMH']
        t_gmh_cor = self.build_ureal(devices.INSTR_DATA[d]['T_correction'])
        t_gmh_raw = GTC.ta.estimate_digitized(gmh_temps, 0.01)
        t_gmh = t_gmh_raw + t_gmh_cor + gmh_t_def
        t_gmh_k = GTC.rp.k_factor(t_gmh.df)
        t_gmh_eu = t_gmh.u*t_gmh_k

        d = this_run['Instruments']['GMHroom']
        rh_cor = self.build_ureal(devices.INSTR_DATA[d]['RH_correction'])
        rh_raw = GTC.ta.estimate_digitized(gmh_room_rhs, 0.1)
        rh = rh_raw*(1 + rh_cor)
        rh_k = GTC.rp.k_factor(rh.df)
        rh_eu = rh.u*rh_k

        # Re-use d (same instrument description)
        press_cor = self.build_ureal(devices.INSTR_DATA[d]['P_correction'])
        press_raw = GTC.ta.estimate_digitized(gmh_room_ps, 0.1)
        press = press_raw * (1 + press_cor)
        press_k = GTC.rp.k_factor(press.df)
        press_eu = press.u * press_k

        self.Results.update({self.run_ID: {}})
        self.ThisResult = self.Results[self.run_ID]
        self.ThisResult.update({'Comment': comment,
                                'Date': mean_date,
                                'Processed date': processed_date,
                                'DUC name': duc_name,
                                'DUC gain': duc_gain,
                                'T_GMH': {},
                                'RH': {},
                                'P': {},
                                'Nom_dV': {}})

        self.ThisResult['T_GMH'].update({'value': t_gmh.x,
                                         'uncert': t_gmh.u,
                                         'dof': t_gmh.df,
                                         'label': t_gmh.label,
                                         'k': t_gmh_k,
                                         'ExpU': t_gmh_eu})

        self.ThisResult['RH'].update({'value': rh.x,
                                      'uncert': rh.u,
                                      'dof': rh.df,
                                      'label': rh.label,
                                      'k': rh_k,
                                      'ExpU': rh_eu})

        self.ThisResult['P'].update({'value': press.x,
                                     'uncert': press.u,
                                     'dof': press.df,
                                     'label': press.label,
                                     'k': press_k,
                                     'ExpU': press_eu})

        self.PSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(press.x, press.u, press.df))  # str(press.s)
        self.Pk.SetValue('{0:.1f}'.format(press_k))
        self.PExpU.SetValue('{0:.2f}'.format(press_eu))
        self.RHSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(rh.x, rh.u, rh.df))  # str(rh.s)
        self.RHk.SetValue('{0:.1f}'.format(rh_k))
        self.RHExpU.SetValue('{0:.2f}'.format(rh_eu))
        self.TGMHSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(t_gmh.x, t_gmh.u, t_gmh.df))  # str(t_gmh.s)
        self.TGMHk.SetValue('{0:.1f}'.format(t_gmh_k))
        self.TGMHExpU.SetValue('{0:.2f}'.format(t_gmh_eu))

        influencies = []
        v1s = []
        v2s = []
        v3s = []

        num_rows = len(this_run['Nom_Vout'])
        for row in range(0, num_rows, 8):  # 0, [8, [16]]
            gains = []  # gains = set()
            # 'neg' and 'pos' refer to polarity of OUTPUT VOLTAGE, not
            # input current! nom_Vout = +/-( 0.1,[1,[10]] ):
            self.nom_Vout = {'pos': this_run['Nom_Vout'][row+2],
                             'neg': this_run['Nom_Vout'][row+1]}
            abs_nom_vout = self.nom_Vout['pos']

            # Construct ureals from raw voltage data, including gain correction
            for n in range(4):
                label_suffix_1 = this_run['Node'][row+n]+'_'+str(n)
                label_suffix_2 = this_run['Node'][row+4+n]+'_'+str(n)
                label_suffix_3 = 'V3' + '_' + str(n)

                v1_v = this_run['IP_V']['val'][row+n]
                v1_u = this_run['IP_V']['sd'][row+n]
                v1_dof = this_run['Nreads'] - 1
                v1_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_1

                d1 = this_run['Instruments']['DVM12']
                gain_param = self.get_gain_err_param(v1_v)
                print(devices.INSTR_DATA[d1].keys())
                gain = self.build_ureal(devices.INSTR_DATA[d1][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v1_raw = GTC.ureal(v1_v, v1_u, v1_dof, label=v1_label)
                v1s.append(GTC.result(v1_raw/gain))

                v2_v = this_run['IP_V']['val'][row+4+n]
                v2_u = this_run['IP_V']['sd'][row+4+n]
                v2_dof = this_run['Nreads'] - 1
                v2_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_2

                d2 = d1  # Same DVM
                gain_param = self.get_gain_err_param(v2_v)
                gain = self.build_ureal(devices.INSTR_DATA[d2][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v2_raw = GTC.ureal(v2_v, v2_u, v2_dof, label=v2_label)
                v2s.append(GTC.result(v2_raw/gain))

                v3_v = this_run['OP_V']['val'][row+n]
                v3_u = this_run['OP_V']['sd'][row+n]
                v3_dof = this_run['Nreads'] - 1
                v3_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_3

                d3 = this_run['Instruments']['DVM3']
                gain_param = self.get_gain_err_param(v3_v)
                gain = self.build_ureal(devices.INSTR_DATA[d3][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v3_raw = GTC.ureal(v3_v, v3_u, v3_dof, label=v3_label)
                v3s.append(GTC.result(v3_raw/gain))

                influencies.extend([v1_raw, v2_raw, v3_raw])

            influencies.extend(gains)  # List of unique gain corrections.
            print('list of gains:')
            for g in gains:
                print('{} +/- {}, dof={}'.format(g.x, g.u, g.df))

            # Offset-adjustment
            v1_pos = GTC.result(v1s[2] - (v1s[0] + v1s[3]) / 2)
            v1_neg = GTC.result(v1s[1] - (v1s[0] + v1s[3]) / 2)
            v2_pos = GTC.result(v2s[2] - (v2s[0] + v2s[3]) / 2)
            v2_neg = GTC.result(v2s[1] - (v2s[0] + v2s[3]) / 2)
            v3_pos = GTC.result(v3s[2] - (v3s[0] + v3s[3]) / 2)
            v3_neg = GTC.result(v3s[1] - (v3s[0] + v3s[3]) / 2)

            # V-drop across Rs
            v_rs_pos = GTC.result(v1_pos - v2_pos)
            v_rs_neg = GTC.result(v1_neg - v2_neg)
            assert v_rs_pos.x * v_rs_neg.x < 0, 'V_Rs polarity error!'

            # Rs Temperature
            t_rs = []
            pt_r_cor = []
            for n, R_raw in enumerate(this_run['Pt_DVM'][row:row+8]):
                pt_r_cor.append(GTC.result(R_raw * (1 + dvmt_cor),
                                           label='Pt_Rcor'+str(n)))
                t_rs.append(GTC.result(self.R_to_T(pt_alpha, pt_beta,
                                                   pt_r_cor[n],
                                                   pt_r0, pt_t_ref)))

            av_t_rs = GTC.result(GTC.fn.mean(t_rs),
                                 label='av_T_Rs'+str(abs_nom_vout))
            influencies.extend(pt_r_cor)
            influencies.extend([pt_alpha, pt_beta, pt_r0, pt_t_ref,
                                dvmt_cor, pt_t_def])  # av_T_Rs
            assert pt_alpha in influencies, 'Influencies missing Pt_alpha!'
            assert pt_beta in influencies, 'Influencies missing Pt_beta!'
            assert pt_r0 in influencies, 'Influencies missing Pt_R0!'
            assert pt_t_ref in influencies, 'Influencies missing Pt_TRef!'
            assert dvmt_cor in influencies, 'Influencies missing DVMT_cor!'
            assert pt_t_def in influencies, 'Influencies missing Pt_T_def!'

            # Value of Rs
            nom_Rs = this_run['Rs']
            msg = 'Nominal Rs value: {0}, '\
                  'Abs. Nom. Vout: {1:.1f}'.format(nom_Rs, abs_nom_vout)
            print( '\n', msg, '\n')
            logger.info(msg)
            rs_name = self.Rs_VAL_NAME[nom_Rs]
            rs_0 = self.build_ureal(devices.RES_DATA[rs_name]['R0_LV'])
            rs_t_ref = self.build_ureal(devices.RES_DATA[rs_name]['TRef_LV'])
            rs_alpha = self.build_ureal(devices.RES_DATA[rs_name]['alpha'])
            rs_beta = self.build_ureal(devices.RES_DATA[rs_name]['beta'])

            # Correct Rs value for temperature
            delta_t = GTC.result(av_t_rs - rs_t_ref + pt_t_def)
            rs = GTC.result(rs_0 * (1 + rs_alpha * delta_t + rs_beta * delta_t ** 2))

            influencies.extend([rs_0, rs_alpha, rs_beta, rs_t_ref])

            '''
            Finally, calculate current-change in,
            for nominal voltage-change out:
            REMEMBER: POSITIVE 'Vout' is caused by 'I_pos'
            (which may be NEGATIVE for inverting device)!
            '''
            i_in_pos = GTC.result(v_rs_pos / rs)
            i_in_neg = GTC.result(v_rs_neg / rs)
            assert i_in_pos.x * i_in_neg.x < 0, 'Iin polarity error!'

            # Calculate I_in that would produce nominal Vout:
            i_pos = GTC.result(i_in_pos * self.nom_Vout['pos'] / v3_pos)
            i_neg = GTC.result(i_in_neg * self.nom_Vout['neg'] / v3_neg)
            assert i_pos.x * i_neg.x < 0, 'I polarity error!'

            this_result = {'pos': i_pos, 'neg': i_neg}

            # build uncertainty budget table
            budget_table = {'pos': [], 'neg': []}
            for i in influencies:
                print('Working through influence variables: {}.'.format(i.label))
                logger.info('Working through influence variables: {}.'.format(i.label))
                if i.u == 0:
                    sensitivity = {'pos': 0, 'neg': 0}
                else:
                    sensitivity = {'pos': GTC.component(i_pos, i)/i.u,
                                   'neg': GTC.component(i_neg, i)/i.u}

                # Only include non-zero influencies:
                if abs(GTC.component(i_pos, i)) > 0:
                    print('Included component of I+: {}'.format(GTC.component(i_pos, i)))
                    logger.info('Included component of I+: {}'.format(GTC.component(i_pos, i)))
                    budget_table['pos'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['pos'],
                                                GTC.component(i_pos, i)])
                else:
                    print('ZERO COMPONENT of I+')
                    logger.info('ZERO COMPONENT of I+')

                if abs(GTC.component(i_neg, i)) > 0:
                    print('Included component of I-: {}'.format(GTC.component(i_neg, i)))
                    logger.info('Included component of I-: {}'.format(GTC.component(i_neg, i)))
                    budget_table['neg'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['neg'],
                                                GTC.component(i_neg, i)])
                else:
                    print('ZERO COMPONENT of I-')
                    logger.info('ZERO COMPONENT of I-')

            # self.budget_table_sorted = {'pos': [], 'neg': []}
            self.budget_table_sorted['pos'] = sorted(budget_table['pos'],
                                                     key=self.by_u_cont,
                                                     reverse=True)
            self.budget_table_sorted['neg'] = sorted(budget_table['neg'],
                                                     key=self.by_u_cont,
                                                     reverse=True)

            # Write results (including budgets)
            self.result_row = self.write_this_result(this_result)
            time.sleep(0.1)
            row += 8
            del influencies[:]
            del v1s[:]
            del v2s[:]
            del v3s[:]
        # <-- End of analysis loop for this run

        # Save analysis result
        self.results_file = os.path.join(os.path.dirname(self.data_file),
                                         'IVY_Results.json')
        with open(self.results_file, 'w') as results_fp:
            json.dump(self.Results, results_fp, indent=4)

        # Ensure Vout c-box is cleared before updating choices.
        # Vout choices should be chosen from: (0.1, -0.1, 1, -1, 10, -10):
        self.NomVout.Clear()
        for V in self.ThisResult['Nom_dV'].keys():
            self.NomVout.Append(str(V))
        # ______________END OF OnAnalyze()________________

    def on_nom_vout_choice(self, e):
        """
        Select which analysed data sub-set (by Vout) to display
        - Delta_Iin and
        - Uncert budget
        """
        v_out = float(e.GetString())
        i_in = self.build_ureal(self.ThisResult['Nom_dV'][v_out]['Delta_Iin'])
        i_in_str = '{0:.5e}, u={1:.1e}, df={2:.1f}'.format(i_in.x, i_in.u, i_in.df)
        i_in_k = GTC.rp.k_factor(i_in.df)
        i_in_EU = i_in.u * i_in_k
        self.IinSummary.SetValue(i_in_str)
        self.Iink.SetValue('{0:.1f}'.format(i_in_k))
        self.IinExpU.SetValue('{0:.2e}'.format(i_in_EU))

        # Build uncertainty budget table (as a string for display):
        budget_str = '{:28}{:<14}{:<10}{:<6}{:<10}{:<13}\n'\
            .format('Quantity', 'Value', 'Std u.', 'dof',
                    'Sens. Co.', 'U. Cont.')
        u_budget_dict = self.ThisResult['Nom_dV'][v_out]['U_Budget']
        for n, q in enumerate(u_budget_dict['quantity(label)']):
            v = u_budget_dict['value'][n]
            u = u_budget_dict['std. uncert'][n]
            d = float(u_budget_dict['dof'][n])
            s = u_budget_dict['sens. co.'][n]
            c = u_budget_dict['uncert. cont.'][n]
            line = '{0:<28}{1:<16.{2}g}{3:<10.2g}{4:<6.1f}{5:<10.1e}{6:<13.1e}\n'\
                .format(q, v, self.set_precision(v, u), u, d, s, c)
            budget_str += line
        self.Budget.SetValue(budget_str)

    @staticmethod
    def set_precision(v, u):
        if v == 0:
            v_lg = 0
        else:
            v_lg = math.log10(abs(v))
        if u == 0:
            u_lg = 0
        else:
            u_lg = math.log10(u)
        if v > u:
            return int(round(v_lg - u_lg)) + 2
        else:
            return 2

    def get_duc_name_from_run_id(self, runid):
        start = 'IVY.v{} '.format(self.version)
        end = ' (Gain='
        return runid[len(start): runid.find(end)]

    def get_mean_date(self):
        '''
        Accept a list of times (as strings).
        Return a mean time (as a string).
        '''
        time_fmt = '%d/%m/%Y %H:%M:%S'
        t_av = 0.0
        t_lst = self.run_data[self.run_ID]['Date_time']

        for t_str in t_lst:
            t_dt = dt.datetime.strptime(t_str, time_fmt)
            t_tup = dt.datetime.timetuple(t_dt)  # A Python time tuple object
            t_av += time.mktime(t_tup)  # time as float (seconds from epoch)

        t_av /= len(t_lst)
        t_av_dt = dt.datetime.fromtimestamp(t_av)
        return t_av_dt.strftime(time_fmt)  # av. time as string

    @staticmethod
    def build_ureal(d):
        """
        Accept a dictionary containing keys of 'value', 'uncert', 'dof' and
        'label' and use corresponding values as input to GTC.ureal().
        Return resulting ureal. Ignore any other keys.
        """
        try:
            assert isinstance(d, dict), 'Not a dictionary!'
            return GTC.ureal(float(d['value']),
                             float(d['uncert']),
                             float(d['dof']),
                             str(d['label']))
        except (AssertionError, KeyError) as msg:
            # raise TypeError('Non-ureal input')
            print('build_ureal():', msg)
            return 0

    @staticmethod
    def get_gain_err_param(v):
        """
        Return the key (a string) identifying the correct gain parameter
        for V (and appropriate range).
        """
        if abs(v) < 0.001:
            nom_v = '0.0001'
            nom_range = '0.1'
        elif abs(v) < 0.022:
            nom_v = '0.01'
            nom_range = '0.1'
        elif abs(v) < 0.071:
            nom_v = '0.05'
            nom_range = '0.1'
        elif abs(v) < 0.22:
            nom_v = '0.1'
            nom_range = '0.1'
        elif abs(v) < 0.71:
            nom_v = '0.5'
            nom_range = '1'
        elif abs(v) < 2.2:
            nom_v = nom_range = '1'
        else:
            nom_v = nom_range = str(int(abs(round(v))))  # '1' or '10'
        gain_param = 'Vgain_{0}r{1}'.format(nom_v, nom_range)
        return gain_param

    def R_to_T(self, alpha, beta, R, R0, T0):
        """
        Convert a resistive T-sensor reading from resistance to temperature.
        All arguments and return value are ureals.
        """
        if beta.x == 0 and beta.u == 0:  # No 2nd-order T-Co
            T = GTC.result((R/R0 - 1)/alpha + T0)
        else:
            a = GTC.result(beta)
            b = GTC.result(alpha - 2*beta*T0, True)
            c = GTC.result(1 - alpha*T0 + beta*T0**2 - (R/R0))
            T = GTC.result((-b + GTC.sqrt(b**2 - 4*a*c))/(2*a))
        return T

    @staticmethod
    def by_u_cont(line):
        """
        A function required for budget-table sorting
        by uncertainty contribution.
        """
        return line[5]

    def write_this_result(self, result):
        """
        Write results and uncert. budget for THIS NOM_VOUT (BOTH polarities)
        """
        result_dict = {'pos': {'Delta_Iin': {}, 'U_Budget': {}},
                       'neg': {'Delta_Iin': {}, 'U_Budget': {}}}

        # Summary for each polarity:
        for polarity in result.keys():  # 'pos' or 'neg'
            value = result[polarity].x
            uncert = result[polarity].u
            dof = result[polarity].df
            lbl = result[polarity].label
            k = GTC.rp.k_factor(dof)
            exp_u = k * uncert

            # Delta_I_in:
            result_dict[polarity]['Delta_Iin'].update({'value': value,
                                                       'uncert': uncert,
                                                       'dof': dof,
                                                       'label': lbl,
                                                       'k': k, 'EU': exp_u})
            # Uncert Budget:
            quantity = []
            v = []
            u = []
            df = []
            sens = []
            cont = []
            for line in self.budget_table_sorted[polarity]:
                for i, heading in enumerate((quantity, v, u, df, sens, cont)):
                    heading.append(line[i])
            result_dict[polarity]['U_Budget'].update({'quantity(label)': quantity,
                                                      'value': v,
                                                      'std. uncert': u,
                                                      'dof': df,
                                                      'sens. co.': sens,
                                                      'uncert. cont.': cont})
            d = {self.nom_Vout[polarity]: result_dict[polarity]}
            self.ThisResult['Nom_dV'].update(d)

        return 1

    @staticmethod
    def add_if_unique(item, lst):
        """
        Append 'item' to 'lst' only if it is not already present.
        """
        if item not in lst:
            lst.append(item)
        return lst
