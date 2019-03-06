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

import matplotlib
matplotlib.use('WXAgg')  # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
# from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick

#from openpyxl import load_workbook, cell
import IVY_events as evts
import acquisition as acq
import devices

import GTC

matplotlib.rc('lines', linewidth=1, color='blue')

logger = logging.getLogger(__name__)

'''
------------------------
# Setup Page definition:
------------------------
'''


class SetupPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # Event bindings
        self.Bind(evts.EVT_FILEPATH, self.UpdateDir)

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
                                   'Rs=10^6': '6',
                                   }  # '': None
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
        self.Sources.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_SRC.append(self.Sources)

        self.IP_DVM_Lbl = wx.StaticText(self, label='Input DVM (DVM12):',
                                        id=wx.ID_ANY)
        self.IP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.IP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.IP_Dvms)
        self.OP_DVM_Lbl = wx.StaticText(self, label='Output DVM (DVM3):',
                                        id=wx.ID_ANY)
        self.OP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.OP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.OP_Dvms)
        self.TDvmLbl = wx.StaticText(self, label='T-probe DVM (DVMT):',
                                     id=wx.ID_ANY)
        self.TDvms = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.DVM_COMBO_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.TDvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.TDvms)

        self.GMHLbl = wx.StaticText(self, label='GMH probe (GMH):',
                                    id=wx.ID_ANY)
        self.GMHProbes = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GMH_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.GMHProbes.Bind(wx.EVT_COMBOBOX, self.BuildCommStr)
        self.cbox_instr_GMH.append(self.GMHProbes)

        self.GMHroomLbl = wx.StaticText(self,
                                        label='Room conds. GMH probe (GMHroom):',
                                        id=wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.GMH_COMBO_CHOICE,
                                         style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)

        self.IVboxLbl = wx.StaticText(self, label='IV_box (IVbox) setting:',
                                      id=wx.ID_ANY)
        self.IVbox = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.IVBOX_COMBO_CHOICE.keys(),
                                 style=wx.CB_DROPDOWN)

        # Addresses
        self.SrcAddr = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.GPIBAddressList,
                                   size=(150, 10), style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.SrcAddr)
        self.SrcAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.IP_DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GPIBAddressList,
                                      style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.IP_DvmAddr)
        self.IP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.OP_DvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                      choices=self.GPIBAddressList,
                                      style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.OP_DvmAddr)
        self.OP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.TDvmAddr = wx.ComboBox(self, wx.ID_ANY,
                                    choices=self.GPIBAddressList,
                                    style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.TDvmAddr)
        self.TDvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.GMHPorts = wx.ComboBox(self, wx.ID_ANY,
                                    choices=self.COMAddressList,
                                    style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHPorts)
        self.GMHPorts.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.GMHroomPorts = wx.ComboBox(self, wx.ID_ANY,
                                        choices=self.COMAddressList,
                                        style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHroomPorts)
        self.GMHroomPorts.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        self.IVboxAddr = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.COMAddressList,
                                     style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.IVboxAddr)
        self.IVboxAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        # Filename
        DirLbl = wx.StaticText(self, label='Working directory:',
                               id=wx.ID_ANY)
        self.WorkingDir = wx.TextCtrl(self, id=wx.ID_ANY,
                                      value=self.GetTopLevelParent().directory,
                                      style=wx.TE_READONLY)

        # DUC
        self.DUCName = wx.TextCtrl(self, id=wx.ID_ANY, value='DUC Name')
        self.DUCName.Bind(wx.EVT_TEXT, self.BuildCommStr)

        # Autopopulate btn
        self.AutoPop = wx.Button(self, id=wx.ID_ANY, label='AutoPopulate')
        self.AutoPop.Bind(wx.EVT_BUTTON, self.OnAutoPop)

        # Test buttons
        self.VisaList = wx.Button(self, id=wx.ID_ANY, label='List Visa res')
        self.VisaList.Bind(wx.EVT_BUTTON, self.OnVisaList)
        self.ResList = wx.TextCtrl(self, id=wx.ID_ANY,
                                   value='Available Visa resources',
                                   style=wx.TE_READONLY | wx.TE_MULTILINE)

        self.STest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.STest.Bind(wx.EVT_BUTTON, self.OnTest)

        self.D12Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.D12Test.Bind(wx.EVT_BUTTON, self.OnTest)

        self.D3Test = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.D3Test.Bind(wx.EVT_BUTTON, self.OnTest)

        self.DTTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.DTTest.Bind(wx.EVT_BUTTON, self.OnTest)

        self.GMHTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMHTest.Bind(wx.EVT_BUTTON, self.OnTest)

        self.GMHroomTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.GMHroomTest.Bind(wx.EVT_BUTTON, self.OnTest)

        self.IVboxTest = wx.Button(self, id=wx.ID_ANY, label='Test')
        self.IVboxTest.Bind(wx.EVT_BUTTON, self.OnIVBoxTest)

        ResponseLbl = wx.StaticText(self,
                                    label='Instrument Test Response:',
                                    id=wx.ID_ANY)
        self.Response = wx.TextCtrl(self, id=wx.ID_ANY, value='',
                                    style=wx.TE_READONLY)

        gbSizer = wx.GridBagSizer()

        # Instruments
        gbSizer.Add(self.SrcLbl, pos=(0, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Sources, pos=(0, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IP_DVM_Lbl, pos=(1, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IP_Dvms, pos=(1, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.OP_DVM_Lbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.OP_Dvms, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.TDvmLbl, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.TDvms, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHLbl, pos=(4, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHProbes, pos=(4, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomLbl, pos=(5, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomProbes, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IVboxLbl, pos=(6, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IVbox, pos=(6, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Addresses
        gbSizer.Add(self.SrcAddr, pos=(0, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IP_DvmAddr, pos=(1, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.OP_DvmAddr, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.TDvmAddr, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHPorts, pos=(4, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomPorts, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IVboxAddr, pos=(6, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # DUC Name
        gbSizer.Add(self.DUCName, pos=(6, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Filename
        gbSizer.Add(DirLbl, pos=(8, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.WorkingDir, pos=(8, 1), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Test buttons
        gbSizer.Add(self.STest, pos=(0, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.D12Test, pos=(1, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.D3Test, pos=(2, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.DTTest, pos=(3, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHTest, pos=(4, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomTest, pos=(5, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IVboxTest, pos=(6, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        gbSizer.Add(ResponseLbl, pos=(3, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Response, pos=(4, 4), span=(1, 3),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.VisaList, pos=(0, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.ResList, pos=(0, 4), span=(3, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Autopopulate btn
        gbSizer.Add(self.AutoPop, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gbSizer)

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
        self.CreateInstr(d, r)

        self.BuildComboChoices()

    def BuildComboChoices(self):
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

    def UpdateDir(self, e):
        '''
        Display working directory once selected.
        '''
        self.WorkingDir.SetValue(e.Dir)

    def OnAutoPop(self, e):
        '''
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        '''
        self.instrument_choice = {'SRC': 'SRC_F5520A',
                                  'DVM12': 'DVM_3458A:s/n518',
                                  'DVM3': 'DVM_3458A:s/n066',
                                  'DVMT': 'DVM_34401A:s/n976',
                                  'GMH': 'GMH:s/n627',
                                  'GMHroom': 'GMH:s/n367'}
        for r in self.instrument_choice.keys():
            d = self.instrument_choice[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d)  # Update instr. cbox
            self.CreateInstr(d, r)
        if self.DUCName.GetValue() == u'DUC Name':
            self.DUCName.SetForegroundColour((255, 0, 0))
            self.DUCName.SetValue('CHANGE_THIS!')

    def UpdateInstr(self, e):
        '''
        An instrument was selected for a role.
        Find description d and role r, then pass to CreatInstr()
        '''
        d = e.GetString()
        for r in devices.ROLES_WIDGETS.keys():  # Cycle through roles
            if devices.ROLES_WIDGETS[r]['icb'] == e.GetEventObject():
                break  # stop looking on finding the right instr & role
        self.CreateInstr(d, r)

    def CreateInstr(self, d, r):
        '''
         Called by both OnAutoPop() and UpdateInstr().
         Create each instrument in software & open visa session
         (GPIB and IVbox only).
         For GMH instruments, use GMH dll, not visa.
        '''
        print '\nCreateInstr(%s,%s)...' % (d, r)
        logger.info('Instr = %s; role = %s)...', d, r)
        if 'GMH' in r:  # Changed from d to r
            # create and open a GMH instrument instance
            msg = 'Creating GMH device ({0:s} -> {1:s})'.format(d, r)
            print'\nnbpages.SetupPage.CreateInstr():', msg
            logger.info(msg)
            devices.ROLES_INSTR.update({r: devices.GMH_Sensor(d, r)})
        else:
            # create a visa instrument instance
            msg = 'Creating VISA device ({0:s} -> {1:s}).'.format(d, r)
            print'\nnbpages.SetupPage.CreateInstr():', msg
            logger.info(msg)
            devices.ROLES_INSTR.update({r: devices.instrument(d, r)})
            devices.ROLES_INSTR[r].Open()
        self.SetInstr(d, r)

    def SetInstr(self, d, r):
        """
        Called by CreateInstr().
        Updates internal info (INSTR_DATA), updates the addresses
        and Enables/disables testbuttons as necessary.
        """
        assert d in devices.INSTR_DATA, 'Unknown instrument: %s' % d
        assert_msg = 'Unknown parameter ("role") for %s: .' % d
        assert 'role' in devices.INSTR_DATA[d], assert_msg
        devices.INSTR_DATA[d]['role'] = r  # update default role

        # Set the address cb to correct value (refer to devices.INSTR_DATA)
        a_cb = devices.ROLES_WIDGETS[r]['acb']
        msg = 'Address = {}'.format(devices.INSTR_DATA[d]['str_addr'])
        print 'SetInstr():', msg
        logger.info(msg)
        a_cb.SetValue((devices.INSTR_DATA[d]['str_addr']))
        if d == 'none':
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(False)
        else:
            devices.ROLES_WIDGETS[r]['tbtn'].Enable(True)

    def UpdateAddr(self, e):
        # An address was manually selected
        # 1st, we'll need instrument description d...
        d = 'none'
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
            addr = a.lstrip('COMGPIB0:')  # leave only numeric part
            devices.INSTR_DATA[d]['addr'] = int(addr)
            devices.ROLES_INSTR[r].addr = int(addr)
        msg = '{0:s} using {1:s} set to '\
              'addr {2:s} ({3:s})'.format(r, d, addr, a)
        print'UpdateAddr():', msg
        logger.info(msg)

    def OnTest(self, e):
        # Called when a 'test' button is clicked
        d = 'none'
        for r in devices.ROLES_WIDGETS.keys():  # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when we've found right instr descr
        print'\nnbpages.SetupPage.OnTest():', d
        logger.info('%s', d)
        assert_msg = '%s has no "test" parameter' % d
        assert 'test' in devices.INSTR_DATA[d], assert_msg
        test = devices.INSTR_DATA[d]['test']  # test string
        print '\tTest string:', test
        logger.info('Test string: %s', test)
        self.Response.SetValue(str(devices.ROLES_INSTR[r].Test(test)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d, test), 0)

    def OnIVBoxTest(self, e):
        # NOTE: config is the configuration description NOT the test string:
        config = devices.ROLES_WIDGETS['IVbox']['icb'].GetValue()
        test = self.IVBOX_COMBO_CHOICE[config]
        if test is not None:
            try:
                devices.ROLES_INSTR['IVbox'].Test(test)
                self.Response.SetValue(config)
            except devices.visa.VisaIOError:
                self.Response.SetValue('IV_box test failed!')
        else:
            self.Response.SetValue('IV_box: empty test!')

    def BuildCommStr(self, e):
        # Called by a change in GMH probe selection, or DUC name
        self.version = self.GetTopLevelParent().version
        d = e.GetString()
        if 'GMH' in d:  # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.CreateInstr(d, r)
        else:  # DUC name has been set or changed
            if d in ('CHANGE_THIS!', 'DUC Name'):  # DUC not yet specified!
                self.DUCName.SetForegroundColour((255, 0, 0))
            else:
                self.DUCName.SetForegroundColour((0, 127, 0))
#        RunPage = self.GetParent().GetPage(1)
#        params = {'DUC': self.DUCName.GetValue(),
#                  'GMH': self.GMHProbes.GetValue()}
#        commstr = 'IVY v.{0:s}. DUC: {1:s} monitored by '\
#                  '{2:s}'.format(self.version, params['DUC'], params['GMH'])
#        evt = evts.UpdateCommentEvent(str=commstr)
#        wx.PostEvent(RunPage, evt)

    def OnVisaList(self, e):
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
            current_COM = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.ComList)
            cbox.SetValue(current_COM)

        # Re-build combobox choices from list of GPIB addresses
        for cbox in self.cbox_addr_GPIB:
            current_COM = cbox.GetValue()
            cbox.Clear()
            cbox.AppendItems(self.GPIBList)
            cbox.SetValue(current_COM)

        # Add resources to ResList TextCtrl widget
        self.res_addr_list = '\n'.join(self.ResourceList)
        self.ResList.SetValue(self.res_addr_list)

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
        self.Bind(evts.EVT_DATA, self.UpdateData)
#        self.Bind(evts.EVT_START_ROW, self.UpdateStartRow)

        self.RunThread = None

        # Comment widgets
        CommentLbl = wx.StaticText(self, id=wx.ID_ANY, label='Comment:')
        self.Comment = wx.TextCtrl(self, id=wx.ID_ANY, size=(600, 20))
#        self.Comment.Bind(wx.EVT_TEXT, self.OnComment)
        comtip = 'Use this field to add remarks and observations that may not'\
                 ' be recorded automatically.'
        self.Comment.SetToolTipString(comtip)

        self.NewRunIDBtn = wx.Button(self, id=wx.ID_ANY,
                                     label='Create new run id')
        idcomtip = 'Create new id to uniquely identify this set of '\
                   'measurement data.'
        self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.OnNewRunID)
        self.RunID = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

        # Run Setup widgets
        DUCgainLbl = wx.StaticText(self, id=wx.ID_ANY,
                                   style=wx.ALIGN_LEFT,
                                   label='DUC gain (V/A):')
        self.DUCgain = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.GAINS_CHOICE,
                                   style=wx.CB_DROPDOWN)
        RsLbl = wx.StaticText(self, id=wx.ID_ANY,
                              style=wx.ALIGN_LEFT, label='I/P Rs:')
        self.Rs = wx.ComboBox(self, wx.ID_ANY, choices=self.Rs_CHOICE,
                              style=wx.CB_DROPDOWN)
        self.Rs.Bind(wx.EVT_COMBOBOX, self.OnRs)
        SettleDelLbl = wx.StaticText(self, id=wx.ID_ANY, label='Settle delay:')
        self.SettleDel = wx.SpinCtrl(self, id=wx.ID_ANY, value='0',
                                     min=0, max=3600)
        SrcLbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_LEFT,
                               label='V1 Setting:')
        self.V1Setting = NumCtrl(self, id=wx.ID_ANY, integerWidth=3,
                                 fractionWidth=8, groupDigits=True)
        self.V1Setting.Bind(wx.lib.masked.EVT_NUM, self.OnV1Set)
        ZeroVoltsBtn = wx.Button(self, id=wx.ID_ANY, label='Set zero volts',
                                 size=(200, 20))
        ZeroVoltsBtn.Bind(wx.EVT_BUTTON, self.OnZeroVolts)

        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, style=wx.LI_HORIZONTAL)

        #  Run control and progress widgets
        self.StartBtn = wx.Button(self, id=wx.ID_ANY, label='Start run')
        self.StartBtn.Bind(wx.EVT_BUTTON, self.OnStart)
        self.StopBtn = wx.Button(self, id=wx.ID_ANY, label='Abort run')
        self.StopBtn.Bind(wx.EVT_BUTTON, self.OnAbort)
        self.StopBtn.Enable(False)
        NodeLbl = wx.StaticText(self, id=wx.ID_ANY, label='Node:')
        self.Node = wx.ComboBox(self, wx.ID_ANY, choices=self.VNODE_CHOICE,
                                style=wx.CB_DROPDOWN)
        self.Node.Bind(wx.EVT_COMBOBOX, self.OnNode)
        VavLbl = wx.StaticText(self, id=wx.ID_ANY, label='Mean V:')
        self.Vav = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)
        VsdLbl = wx.StaticText(self, id=wx.ID_ANY, label='Stdev V:')
        self.Vsd = NumCtrl(self, id=wx.ID_ANY, integerWidth=3, fractionWidth=9,
                           groupDigits=True)
        TimeLbl = wx.StaticText(self, id=wx.ID_ANY, label='Timestamp:')
        self.Time = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY,
                                size=(200, 20))
        RowLbl = wx.StaticText(self, id=wx.ID_ANY,
                               label='Current measurement:')
        self.Row = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        ProgressLbl = wx.StaticText(self, id=wx.ID_ANY, style=wx.ALIGN_RIGHT,
                                    label='Run progress:')
        self.Progress = wx.Gauge(self, id=wx.ID_ANY, range=100,
                                 name='Progress')

        gbSizer = wx.GridBagSizer()

        # Comment widgets
        gbSizer.Add(CommentLbl, pos=(0, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Comment, pos=(0, 1), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.NewRunIDBtn, pos=(1, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.RunID, pos=(1, 1), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Run setup widgets
        gbSizer.Add(DUCgainLbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.DUCgain, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(RsLbl, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Rs, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(SettleDelLbl, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.SettleDel, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(SrcLbl, pos=(2, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.V1Setting, pos=(3, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(ZeroVoltsBtn, pos=(3, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.h_sep1, pos=(4, 0), span=(1, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)

        #  Run control and progress widgets
        gbSizer.Add(self.StartBtn, pos=(5, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.StopBtn, pos=(6, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(NodeLbl, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Node, pos=(6, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(VavLbl, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Vav, pos=(6, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(VsdLbl, pos=(5, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Vsd, pos=(6, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(TimeLbl, pos=(5, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Time, pos=(6, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(RowLbl, pos=(5, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Row, pos=(6, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(ProgressLbl, pos=(7, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Progress, pos=(7, 1), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gbSizer)

        self.autocomstr = ''
        self.manstr = ''

        # Dictionary to hold ALL runs for this application session:
        self.master_run_dict = {}

    def OnNewRunID(self, e):
        self.version = self.GetTopLevelParent().version
        DUCname = self.SetupPage.DUCName.GetValue()
        self.run_id = str('IVY.v' + self.version + ' ' + DUCname + ' (Gain=' +
                          self.DUCgain.GetValue() + '; Rs=' +
                          self.Rs.GetValue() + ') ' +
                          dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.status.SetStatusText('Id for subsequent runs:', 0)
        self.status.SetStatusText(str(self.run_id), 1)
        self.RunID.SetValue(str(self.run_id))

    def UpdateData(self, e):
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

    def OnRs(self, e):
        self.Rs_val = self.Rs_choice_to_val[e.GetString()]  # an INT
        print '\nRunPage.OnRs(): Rs =', self.Rs_val
        logger.info('Rs = %d', self.Rs_val)
        if e.GetString() in self.Rs_SWITCHABLE:  # a STRING
            s = str(int(math.log10(self.Rs_val)))  # '3','4','5' or '6'
            print '\nSwitching Rs - Sending "%s" to IVbox' % s
            logger.info('Switching Rs - Sending "%s" to IVbox', s)
            devices.ROLES_INSTR['IVbox'].SendCmd(s)

    def OnNode(self, e):
        node = e.GetString()  # 'V1', 'V2', or 'V3'
        print'\nRunPage.OnNode():', node
        logger.info('Node = %s', node)
        s = node[1]
        if s in ('1', '2'):
            print'\nRunPage.OnNode():Sending IVbox "', s, '"'
            logger.info('Sending IVbox "%s"', s)
            devices.ROLES_INSTR['IVbox'].SendCmd(s)
        else:  # '3'
            print'\nRunPage.OnNode():IGNORING IVbox cmd "', s, '"'
            logger.info('IGNORING IVbox cmd "%s"', s)

    def OnV1Set(self, e):
        # Called by change in value (manually OR by software!)
        V1 = e.GetValue()
        print'RunPage.OnV1Set(): V1 =', V1, '(', type(V1), ')'
        logger.info('V1 = %s', V1)
        src = devices.ROLES_INSTR['SRC']
        src.SetV(V1)
        time.sleep(0.5)
        if V1 == 0:
            src.Stby()
        else:
            src.Oper()
        time.sleep(0.5)

    def OnZeroVolts(self, e):
        # V1:
        src = devices.ROLES_INSTR['SRC']
        if self.V1Setting.GetValue() == 0:
            print'RunPage.OnZeroVolts(): Zero/Stby directly'
            logger.info('Zero/Stby directly')
            src.SetV(0)
            src.Stby()
        else:
            self.V1Setting.SetValue('0')  # Calls OnV1Set() ONLY IF VAL CHANGES
            print'RunPage.OnZeroVolts():  Zero/Stby via V1 display'
            logger.info('Zero/Stby via V1 display')

    def OnStart(self, e):
        self.Progress.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('', 1)
        self.status.SetStatusText('Starting run', 0)
        if self.RunThread is None:
            self.StopBtn.Enable(True)  # Enable Stop button
            self.StartBtn.Enable(False)  # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)

    def OnAbort(self, e):
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

        self.Bind(evts.EVT_PLOT, self.UpdatePlot)
        self.Bind(evts.EVT_CLEARPLOT, self.ClearPlot)

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
        V1_y_ost = self.V1ax.get_xaxis().get_offset_text()
        V1_y_ost.set_visible(False)

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
        V2_y_ost = self.V2ax.get_xaxis().get_offset_text()
        V2_y_ost.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)

    def UpdatePlot(self, e):
        print'PlotPage.UpdatePlot(): len(t)=', len(e.t)
        logger.info('len(t) = %d', len(e.t))
        print e.node, 'len(V1)=', len(e.V12), 'len(V3)=', len(e.V3)
        logger.info('len(V1) = %d, len(V3) = %d', len(e.V12), len(e.V3))
        if e.node == 'V1':
            self.V1ax.plot_date(e.t, e.V12, 'bo')
        else:  # V2 data
            self.V2ax.plot_date(e.t, e.V12, 'go')
        self.V3ax.plot_date(e.t, e.V3, 'ro')
        self.figure.autofmt_xdate()  # default settings
        self.V3ax.fmt_xdata = mdates.DateFormatter('%d-%m-%Y, %H:%M:%S')
        self.canvas.draw()
        self.canvas.Refresh()

    def ClearPlot(self, e):
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
        self.data_file = self.GetTopLevelParent().data_file
        self.results_file = os.path.join(os.path.dirname(self.data_file),
                                         'IVY_Results.json')
        self.Results = {}  # Accumulate run-analyses here

        self.Rs_VALUES = self.GetTopLevelParent().page2.Rs_VALUES
        self.Rs_NAMES = ['IV1k 1k', 'IV10k 10k',
                         'IV100k 100k', 'IV1M 1M', 'IV10M 10M',
                         'IV100M 100M', 'IV1G 1G']
        self.Rs_VAL_NAME = dict(zip(self.Rs_VALUES, self.Rs_NAMES))

#        self.RunID_choices = []

        gbSizer = wx.GridBagSizer()

        # Analysis set-up:
        self.ListRuns = wx.Button(self, id=wx.ID_ANY, label='List run IDs')
        self.ListRuns.Bind(wx.EVT_BUTTON, self.OnListRuns)
        gbSizer.Add(self.ListRuns, pos=(0, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.RunID = wx.ComboBox(self, id=wx.ID_ANY,
                                 style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.RunID.Bind(wx.EVT_COMBOBOX, self.OnRunChoice)
        self.RunID.Bind(wx.EVT_TEXT, self.OnRunChoice)
        gbSizer.Add(self.RunID, pos=(0, 1), span=(1, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.Analyze = wx.Button(self, id=wx.ID_ANY, label='Analyze')
        self.Analyze.Bind(wx.EVT_BUTTON, self.OnAnalyze)
        gbSizer.Add(self.Analyze, pos=(0, 7), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)  #

        # -----------------------------------------------------------------
        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, size=(720, 1),
                                    style=wx.LI_HORIZONTAL)
        gbSizer.Add(self.h_sep1, pos=(1, 0), span=(1, 8),
                    flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------

        # Run summary:
        RunInfoLbl = wx.StaticText(self, id=wx.ID_ANY, label='Run Summary:')
        gbSizer.Add(RunInfoLbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.RunInfo = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_MULTILINE |
                                   wx.TE_READONLY | wx.HSCROLL,
                                   size=(250, 1))
        gbSizer.Add(self.RunInfo, pos=(3, 0), span=(20, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Analysis results:
        # Headings & environment readings
        QuantityLbl = wx.StaticText(self, id=wx.ID_ANY, label='Quantity')
        gbSizer.Add(QuantityLbl, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        UrealSummaryLbl = wx.StaticText(self, id=wx.ID_ANY, size=(200, 1),
                                        label='Val, Unc, DoF')
        gbSizer.Add(UrealSummaryLbl, pos=(2, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        KLbl = wx.StaticText(self, id=wx.ID_ANY, label='k(95%)')
        gbSizer.Add(KLbl, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        ExpULbl = wx.StaticText(self, id=wx.ID_ANY, label='Exp Uncert')
        gbSizer.Add(ExpULbl, pos=(2, 6), span=(1, 1),
                    flag=wx.ALL, border=5)  # | wx.EXPAND

        PLbl = wx.StaticText(self, id=wx.ID_ANY, label='Pressure:')
        gbSizer.Add(PLbl, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.PSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.PSummary, pos=(3, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.Pk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.Pk, pos=(3, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.PExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.PExpU, pos=(3, 6), span=(1, 1),
                    flag=wx.ALL, border=5)  # | wx.EXPAND

        RHLbl = wx.StaticText(self, id=wx.ID_ANY, label='%RH:')
        gbSizer.Add(RHLbl, pos=(4, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.RHSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.RHSummary, pos=(4, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.RHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.RHk, pos=(4, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.RHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.RHExpU, pos=(4, 6), span=(1, 1),
                    flag=wx.ALL, border=5)  # | wx.EXPAND

        TGMHLbl = wx.StaticText(self, id=wx.ID_ANY, label='T (GMH):')
        gbSizer.Add(TGMHLbl, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHSummary = wx.TextCtrl(self, id=wx.ID_ANY,
                                       style=wx.TE_READONLY)
        gbSizer.Add(self.TGMHSummary, pos=(5, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.TGMHk, pos=(5, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.TGMHExpU, pos=(5, 6), span=(1, 1),
                    flag=wx.ALL, border=5)  # | wx.EXPAND
        # -----------------------------------------------------------------
        self.h_sep3 = wx.StaticLine(self, id=wx.ID_ANY, size=(480, 1),
                                    style=wx.LI_HORIZONTAL)
        gbSizer.Add(self.h_sep3, pos=(6, 2), span=(1, 5),
                    flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------
        VoutLbl = wx.StaticText(self, id=wx.ID_ANY,
                                label='Nom. ' + DELTA + 'Vout:')
        gbSizer.Add(VoutLbl, pos=(7, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.NomVout = wx.ComboBox(self, id=wx.ID_ANY,
                                   style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.NomVout.Bind(wx.EVT_COMBOBOX, self.OnNomVoutChoice)
        gbSizer.Add(self.NomVout, pos=(7, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        IinLbl = wx.StaticText(self, id=wx.ID_ANY, label=DELTA+'I_in:')
        gbSizer.Add(IinLbl, pos=(8, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.IinSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.IinSummary, pos=(8, 3), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.Iink = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.Iink, pos=(8, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.IinExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.IinExpU, pos=(8, 6), span=(1, 1),
                    flag=wx.ALL, border=5)  # | wx.EXPAND

        self.Budget = wx.TextCtrl(self, style=wx.TE_MULTILINE |
                                  wx.TE_READONLY | wx.HSCROLL, id=wx.ID_ANY,)
        budget_font = wx.Font(8, wx.MODERN, wx.NORMAL, wx.NORMAL)
        self.Budget.SetFont(budget_font)
        gbSizer.Add(self.Budget, pos=(9, 2), span=(14, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)  #

        self.SetSizerAndFit(gbSizer)

    def OnListRuns(self, e):
        '''
        Open the .json data file and de-serialize into the dictionary
        self.run_data for later access. Update the choices in the 'Run ID'
        widget to the run ids (used as primary keys in self.run_data).
        '''
        self.data_file = self.GetTopLevelParent().data_file
        with open(self.data_file, 'r') as in_file:
            self.run_data = json.load(in_file)
            self.run_IDs = self.run_data.keys()

        self.RunID.Clear()
        self.RunID.AppendItems(self.run_IDs)
        self.RunID.SetSelection(0)

    def OnRunChoice(self, e):
        ID = e.GetString()
        self.runstr = json.dumps(self.run_data[ID], indent=4)
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

    def OnAnalyze(self, e):
        self.run_ID = self.RunID.GetValue()
        self.this_run = self.run_data[self.run_ID]

        logger.info('STARTING ANALYSIS...')

        # Correction for Pt-100 sensor DVM:
        DVMT = self.this_run['Instruments']['DVMT']
        DVMT_cor = self.BuildUreal(devices.INSTR_DATA[DVMT]['correction_100r'])

        '''
        Pt sensor is a few cm away from input resistors, so assume a
        fairly large type B Tdef of 0.1 deg C:
        '''
        Pt_T_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                             3, label='Pt_T_def')

        Pt_alpha = self.BuildUreal(devices.RES_DATA['Pt 100r']['alpha'])
        Pt_beta = self.BuildUreal(devices.RES_DATA['Pt 100r']['beta'])
        Pt_R0 = self.BuildUreal(devices.RES_DATA['Pt 100r']['R0_LV'])
        Pt_TRef = self.BuildUreal(devices.RES_DATA['Pt 100r']['TRef_LV'])

        '''
        GMH sensor is a few cm away from DUC which, itself, has a size of
        several cm, so assume a fairly large type B Tdef of 0.1 deg C:
        '''
        GMH_T_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                              3, label='GMH_T_def')

        Comment = self.this_run['Comment']
        DUC_name = self.GetDUCNamefromRunID(self.run_ID)
        DUC_gain = self.this_run['DUC_G']
        Mean_date = self.GetMeanDate()
        Proc_date = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        print'Comment:', Comment
        print'Run_Id:', self.run_ID
        print'gain =', DUC_gain
        print 'Mean_date:', Mean_date
        logger.info('Comment: %s\nRun_ID: %s\ngain = %s\nMean_date: %s',
                    Comment, self.run_ID, DUC_gain, Mean_date)

        # Determine mean env. conditions
        GMH_Ts = self.this_run['T_GMH']
        GMHroom_RHs = self.this_run['Room_conds']['RH']
        GMHroom_Ps = self.this_run['Room_conds']['P']

        d = self.this_run['Instruments']['GMH']
        T_GMH_cor = self.BuildUreal(devices.INSTR_DATA[d]['T_correction'])
        T_GMH_raw = GTC.ta.estimate_digitized(GMH_Ts, 0.01)
        T_GMH = T_GMH_raw + T_GMH_cor + GMH_T_def
        T_GMH_k = GTC.rp.k_factor(T_GMH.df)
        T_GMH_EU = T_GMH.u*T_GMH_k

        d = self.this_run['Instruments']['GMHroom']
        RH_cor = self.BuildUreal(devices.INSTR_DATA[d]['RH_correction'])
        RH_raw = GTC.ta.estimate_digitized(GMHroom_RHs, 0.1)
        RH = RH_raw*(1 + RH_cor)
        RH_k = GTC.rp.k_factor(RH.df)
        RH_EU = RH.u*RH_k

        # Re-use d (same instrument description)
        P_cor = self.BuildUreal(devices.INSTR_DATA[d]['P_correction'])
        P_raw = GTC.ta.estimate_digitized(GMHroom_Ps, 0.1)
        P = P_raw*(1 + P_cor)
        P_k = GTC.rp.k_factor(P.df)
        P_EU = P.u*P_k

        self.Results.update({self.run_ID: {}})
        self.ThisResult = self.Results[self.run_ID]
        self.ThisResult.update({'Comment': Comment,
                                'Date': Mean_date,
                                'Processed date': Proc_date,
                                'DUC name': DUC_name,
                                'DUC gain': DUC_gain,
                                'T_GMH': {},
                                'RH': {},
                                'P': {},
                                'Nom_dV': {}})

        self.ThisResult['T_GMH'].update({'value': T_GMH.x,
                                         'uncert': T_GMH.u,
                                         'dof': T_GMH.df,
                                         'label': T_GMH.label,
                                         'k': T_GMH_k,
                                         'ExpU': T_GMH_EU})

        self.ThisResult['RH'].update({'value': RH.x,
                                      'uncert': RH.u,
                                      'dof': RH.df,
                                      'label': RH.label,
                                      'k': RH_k,
                                      'ExpU': RH_EU})

        self.ThisResult['P'].update({'value': P.x,
                                     'uncert': P.u,
                                     'dof': P.df,
                                     'label': P.label,
                                     'k': P_k,
                                     'ExpU': P_EU})

        self.PSummary.SetValue(str(P.s))
        self.Pk.SetValue('{0:.1f}'.format(P_k))
        self.PExpU.SetValue('{0:.2f}'.format(P_EU))
        self.RHSummary.SetValue(str(RH.s))
        self.RHk.SetValue('{0:.1f}'.format(RH_k))
        self.RHExpU.SetValue('{0:.2f}'.format(RH_EU))
        self.TGMHSummary.SetValue(str(T_GMH.s))
        self.TGMHk.SetValue('{0:.1f}'.format(T_GMH_k))
        self.TGMHExpU.SetValue('{0:.2f}'.format(T_GMH_EU))

        influencies = []
        V1s = []
        V2s = []
        V3s = []

        num_rows = len(self.this_run['Nom_Vout'])
        for row in range(0, num_rows, 8):  # 0, [8, [16]]
            gains = set()
            # 'neg' and 'pos' refer to polarity of OUTPUT VOLTAGE, not
            # input current! nom_Vout = +/-( 0.1,[1,[10]] ):
            self.nom_Vout = {'pos': self.this_run['Nom_Vout'][row+2],
                             'neg': self.this_run['Nom_Vout'][row+1]}
            abs_nom_Vout = self.nom_Vout['pos']

            # Construct ureals from raw voltage data, including gain correction
            for n in range(4):
                label_suffix_1 = self.this_run['Node'][row+n]+'_'+str(n)
                label_suffix_2 = self.this_run['Node'][row+4+n]+'_'+str(n)
                label_suffix_3 = 'V3' + '_' + str(n)

                V1_v = self.this_run['IP_V']['val'][row+n]
                V1_u = self.this_run['IP_V']['sd'][row+n]
                V1_d = self.this_run['Nreads'] - 1
                V1_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_1

                d1 = self.this_run['Instruments']['DVM12']
                gain_param = self.get_gain_err_param(V1_v)
                gain = self.BuildUreal(devices.INSTR_DATA[d1][gain_param])
                gains.add(gain)
                V1_raw = GTC.ureal(V1_v, V1_u, V1_d, label=V1_l)
                V1s.append(GTC.result(V1_raw/gain))

                V2_v = self.this_run['IP_V']['val'][row+4+n]
                V2_u = self.this_run['IP_V']['sd'][row+4+n]
                V2_d = self.this_run['Nreads'] - 1
                V2_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_2

                d2 = d1  # Same DVM
                gain_param = self.get_gain_err_param(V2_v)
                gain = self.BuildUreal(devices.INSTR_DATA[d2][gain_param])
                gains.add(gain)
                V2_raw = GTC.ureal(V2_v, V2_u, V2_d, label=V2_l)
                V2s.append(GTC.result(V2_raw/gain))

                V3_v = self.this_run['OP_V']['val'][row+n]
                V3_u = self.this_run['OP_V']['sd'][row+n]
                V3_d = self.this_run['Nreads'] - 1
                V3_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_3

                d3 = self.this_run['Instruments']['DVM3']
                gain_param = self.get_gain_err_param(V3_v)
                gain = self.BuildUreal(devices.INSTR_DATA[d3][gain_param])
                gains.add(gain)
                V3_raw = GTC.ureal(V3_v, V3_u, V3_d, label=V3_l)
                V3s.append(GTC.result(V3_raw/gain))

                influencies.extend([V1_raw, V2_raw, V3_raw])

            influencies.extend(list(gains))  # List of unique gain corrections.
            print 'list of gains:'
            for g in list(gains):
                print g.s

            # Offset-adjustment
            V1_pos = GTC.result(V1s[2] - (V1s[0] + V1s[3]) / 2)
            V1_neg = GTC.result(V1s[1] - (V1s[0] + V1s[3]) / 2)
            V2_pos = GTC.result(V2s[2] - (V2s[0] + V2s[3]) / 2)
            V2_neg = GTC.result(V2s[1] - (V2s[0] + V2s[3]) / 2)
            V3_pos = GTC.result(V3s[2] - (V3s[0] + V3s[3]) / 2)
            V3_neg = GTC.result(V3s[1] - (V3s[0] + V3s[3]) / 2)

            # V-drop across Rs
            V_Rs_pos = GTC.result(V1_pos - V2_pos)
            V_Rs_neg = GTC.result(V1_neg - V2_neg)
            assert V_Rs_pos.x * V_Rs_neg.x < 0, 'V_Rs polarity error!'

            # Rs Temperature
            T_Rs = []
            Pt_R_cor = []
            for n, R_raw in enumerate(self.this_run['Pt_DVM'][row:row+8]):
                Pt_R_cor.append(GTC.result(R_raw * (1 + DVMT_cor),
                                           label='Pt_Rcor'+str(n)))
                T_Rs.append(GTC.result(self.R_to_T(Pt_alpha, Pt_beta,
                                                   Pt_R_cor[n],
                                                   Pt_R0, Pt_TRef)))

            av_T_Rs = GTC.result(GTC.fn.mean(T_Rs),
                                 label='av_T_Rs'+str(abs_nom_Vout))
            influencies.extend(Pt_R_cor)
            influencies.extend([Pt_alpha, Pt_beta, Pt_R0, Pt_TRef,
                                DVMT_cor, Pt_T_def])  # av_T_Rs
            assert Pt_alpha in influencies, 'Influencies missing Pt_alpha!'
            assert Pt_beta in influencies, 'Influencies missing Pt_beta!'
            assert Pt_R0 in influencies, 'Influencies missing Pt_R0!'
            assert Pt_TRef in influencies, 'Influencies missing Pt_TRef!'
            assert DVMT_cor in influencies, 'Influencies missing DVMT_cor!'
            assert Pt_T_def in influencies, 'Influencies missing Pt_T_def!'

            # Value of Rs
            nom_Rs = self.this_run['Rs']
            msg = 'Nominal Rs value: {0}, '\
                  'Abs. Nom. Vout: {1:.1f}'.format(nom_Rs, abs_nom_Vout)
            print '\n', msg, '\n'
            logger.info(msg)
            Rs_name = self.Rs_VAL_NAME[nom_Rs]
            Rs_0 = self.BuildUreal(devices.RES_DATA[Rs_name]['R0_LV'])
            Rs_TRef = self.BuildUreal(devices.RES_DATA[Rs_name]['TRef_LV'])
            Rs_alpha = self.BuildUreal(devices.RES_DATA[Rs_name]['alpha'])
            Rs_beta = self.BuildUreal(devices.RES_DATA[Rs_name]['beta'])

            # Correct Rs value for temperature
            dT = GTC.result(av_T_Rs - Rs_TRef + Pt_T_def)
            Rs = GTC.result(Rs_0*(1 + Rs_alpha*dT + Rs_beta*dT**2))

            influencies.extend([Rs_0, Rs_alpha, Rs_beta, Rs_TRef])

            '''
            Finally, calculate current-change in,
            for nominal voltage-change out:
            REMEMBER: POSITIVE 'Vout' is caused by 'I_pos'
            (which may be NEGATIVE for inverting device)!
            '''
            Iin_pos = GTC.result(V_Rs_pos/Rs)
            Iin_neg = GTC.result(V_Rs_neg/Rs)
            assert Iin_pos.x * Iin_neg.x < 0, 'Iin polarity error!'

            # Calculate I_in that would produce nominal Vout:
            I_pos = GTC.result(Iin_pos*self.nom_Vout['pos']/V3_pos)
            I_neg = GTC.result(Iin_neg*self.nom_Vout['neg']/V3_neg)
            assert I_pos.x * I_neg.x < 0, 'I polarity error!'

            this_result = {'pos': I_pos, 'neg': I_neg}

            # build uncertainty budget table
            budget_table = {'pos': [], 'neg': []}
            for i in influencies:
                print'Working through influence variables:', i.label
                logger.info('Working through influence variables: %s', i.label)
                if i.u == 0:
                    sensitivity = {'pos': 0, 'neg': 0}
                else:
                    sensitivity = {'pos': GTC.component(I_pos, i)/i.u,
                                   'neg': GTC.component(I_neg, i)/i.u}

                # Only include non-zero influencies:
                if abs(GTC.component(I_pos, i)) > 0:
                    print 'Included component of I+:', GTC.component(I_pos, i)
                    logger.info('Included component of I+: %d',
                                GTC.component(I_pos, i))
                    budget_table['pos'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['pos'],
                                                GTC.component(I_pos, i)])
                else:
                    print'ZERO COMPONENT of I+'
                    logger.info('ZERO COMPONENT of I+')

                if abs(GTC.component(I_neg, i)) > 0:
                    print 'Included component of I-:', GTC.component(I_neg, i)
                    logger.info('Included component of I-: %d',
                                GTC.component(I_neg, i))
                    budget_table['neg'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['neg'],
                                                GTC.component(I_neg, i)])
                else:
                    print'ZERO COMPONENT of I-'
                    logger.info('ZERO COMPONENT of I-')

            self.budget_table_sorted = {'pos': [], 'neg': []}
            self.budget_table_sorted['pos'] = sorted(budget_table['pos'],
                                                     key=self.by_u_cont,
                                                     reverse=True)
            self.budget_table_sorted['neg'] = sorted(budget_table['neg'],
                                                     key=self.by_u_cont,
                                                     reverse=True)

            # Write results (including budgets)
            self.result_row = self.WriteThisResult(this_result)
            time.sleep(0.1)
            row += 8
            del influencies[:]
            del V1s[:]
            del V2s[:]
            del V3s[:]
        # <-- End of analysis loop for this run

        # Save analysis result
        with open(self.results_file, 'w') as results_fp:
            json.dump(self.Results, results_fp, indent=4)

        # Ensure Vout c-box is cleared before updating choices.
        # Vout choices should be chosen from: (0.1, -0.1, 1, -1, 10, -10):
        self.NomVout.Clear()
        for V in self.ThisResult['Nom_dV'].keys():
            self.NomVout.Append(str(V))
        # ______________END OF OnAnalyze()________________

    def OnNomVoutChoice(self, e):
        '''
        Select which analysed data sub-set (by Vout) to display
        - Delta_Iin and
        - Uncert budget
        '''
        Vout = float(e.GetString())
        Iin = self.BuildUreal(self.ThisResult['Nom_dV'][Vout]['Delta_Iin'])
        IinStr = '{0:.5e}, u={1:.1e}, df={2:.1f}'.format(Iin.x, Iin.u, Iin.df)
        Iin_k = GTC.rp.k_factor(Iin.df)
        Iin_EU = Iin.u*Iin_k
        self.IinSummary.SetValue(IinStr)
        self.Iink.SetValue('{0:.1f}'.format(Iin_k))
        self.IinExpU.SetValue('{0:.2e}'.format(Iin_EU))

        # Build uncertainty budget table (as a string for display):
        budget_str = '{:28}{:<14}{:<10}{:<6}{:<10}{:<13}\n'\
            .format('Quantity', 'Value', 'Std u.', 'dof',
                    'Sens. Co.', 'U. Cont.')
        u_budget_dict = self.ThisResult['Nom_dV'][Vout]['U_Budget']
        for n, q in enumerate(u_budget_dict['quantity(label)']):
            v = u_budget_dict['value'][n]
            u = u_budget_dict['std. uncert'][n]
            d = float(u_budget_dict['dof'][n])
            s = u_budget_dict['sens. co.'][n]
            c = u_budget_dict['uncert. cont.'][n]
            line = '{:<28}{:<14.5g}{:<10.1g}{:<6.1f}{:<10.1e}{:<13.1e}\n'\
                .format(q, v, u, d, s, c)
            budget_str += line
        self.Budget.SetValue(budget_str)

#    def GetNamefromComment(self, c):
#        return c[c.find('DUC: ') + 5: c.find(' monitored by GMH')]
    def GetDUCNamefromRunID(self, runid):
        start = 'IVY.v' + self.version + ' '
        end = ' (Gain='
        return runid[len(start)-1: runid.find(end)]

    def GetMeanDate(self):
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

    def BuildUreal(self, d):
        '''
        Accept a dictionary contaning keys of 'value', 'uncert', 'dof' and
        'label' and use corresponding values as input to GTC.ureal().
        Return resulting ureal. Ignore any other keys.
        '''
        assert isinstance(d, dict), 'Not a dictionary!'
        try:
            return GTC.ureal(float(d['value']),
                             float(d['uncert']),
                             float(d['dof']),
                             str(d['label']))
        except:
            raise TypeError('Non-ureal input')
            return 0

    def get_gain_err_param(self, V):
        '''
        Return the key (a string) identifying the correct gain parameter
        for V (and appropriate range).
        '''
        if abs(V) < 0.001:
            nomV = '0.0001'
            nomRange = '0.1'
        elif abs(V) < 0.022:
            nomV = '0.01'
            nomRange = '0.1'
        elif abs(V) < 0.071:
            nomV = '0.05'
            nomRange = '0.1'
        elif abs(V) < 0.22:
            nomV = '0.1'
            nomRange = '0.1'
        elif abs(V) < 0.71:
            nomV = '0.5'
            nomRange = '1'
        elif abs(V) < 2.2:
            nomV = nomRange = '1'
        else:
            nomV = nomRange = str(int(abs(round(V))))  # '1' or '10'
        gain_param = 'Vgain_{0}r{1}'.format(nomV, nomRange)
        return gain_param

    def R_to_T(self, alpha, beta, R, R0, T0):
        '''
        Convert a resistive T-sensor reading from resistance to temperature.
        All arguments and return value are ureals.
        '''
        if (beta.x == 0 and beta.u == 0):  # No 2nd-order T-Co
            T = GTC.result((R/R0 - 1)/alpha + T0)
        else:
            a = GTC.result(beta)
            b = GTC.result(alpha - 2*beta*T0, True)
            c = GTC.result(1 - alpha*T0 + beta*T0**2 - (R/R0))
            T = GTC.result((-b + GTC.sqrt(b**2 - 4*a*c))/(2*a))
        return T

    def by_u_cont(self, line):
        '''
        A function required for budget-table sorting
        by uncertainty contribution.
        '''
        return line[5]

    def WriteThisResult(self, result):
        '''
        Write results and uncert. budget for THIS NOM_VOUT (BOTH polarities)
        '''
        result_dict = {'pos': {'Delta_Iin': {}, 'U_Budget': {}},
                       'neg': {'Delta_Iin': {}, 'U_Budget': {}}}

        # Summary for each polarity:
        for polarity in result.keys():  # 'pos' or 'neg'
            value = result[polarity].x
            uncert = result[polarity].u
            dof = result[polarity].df
            lbl = result[polarity].label
            k = GTC.rp.k_factor(dof)
            EU = k*uncert

            # Delta_I_in:
            result_dict[polarity]['Delta_Iin'].update({'value': value,
                                                       'uncert': uncert,
                                                       'dof': dof,
                                                       'label': lbl,
                                                       'k': k, 'EU': EU})
            # Uncert Budget:
            Q = []
            v = []
            u = []
            df = []
            sens = []
            cont = []
            for line in self.budget_table_sorted[polarity]:
                for i, heading in enumerate((Q, v, u, df, sens, cont)):
                    heading.append(line[i])
            result_dict[polarity]['U_Budget'].update({'quantity(label)': Q,
                                                      'value': v,
                                                      'std. uncert': u,
                                                      'dof': df,
                                                      'sens. co.': sens,
                                                      'uncert. cont.': cont})
            d = {self.nom_Vout[polarity]: result_dict[polarity]}
            self.ThisResult['Nom_dV'].update(d)

        return 1
