# -*- coding: utf-8 -*-
""" nbpages.py - Defines individual notebook pages as panel-like objects

DEVELOPMENT VERSION

Created on Tue Jun 30 10:10:16 2015

@author: t.lawson
"""

import os

import wx
from wx.lib.masked import NumCtrl
import datetime as dt
import time
import math

import matplotlib
matplotlib.use('WXAgg')  # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
# from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick

from openpyxl import load_workbook, cell
import IVY_events as evts
import acquisition as acq
import devices

import GTC
from numbers import Number

matplotlib.rc('lines', linewidth=1, color='blue')


'''
------------------------
# Setup Page definition:
------------------------
'''


class SetupPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # Event bindings
        self.Bind(evts.EVT_FILEPATH, self.UpdateFilepath)

        self.status = self.GetTopLevelParent().sb
        self.version = self.GetTopLevelParent().version

        self.SRC_COMBO_CHOICE = ['none']
        self.DVM_COMBO_CHOICE = ['none']
        self.GMH_COMBO_CHOICE = ['none']
        self.IVBOX_COMBO_CHOICE = ['IV_box', 'none']
        self.T_SENSOR_CHOICE = devices.T_Sensors
        self.cbox_addr_COM = []
        self.cbox_addr_GPIB = []
        self.cbox_instr_SRC = []
        self.cbox_instr_DVM = []
        self.cbox_instr_GMH = []

        self.BuildComboChoices()

        self.GMH1Addr = self.GMH2Addr = 0  # invalid initial address as default

        self.ResourceList = []
        self.ComList = []
        self.GPIBList = []
        self.GPIBAddressList = ['addresses', 'GPIB0::0']  # dummy values
        self.COMAddressList = ['addresses', 'COM0']  # dummy values initially.

        self.test_btns = []  # list of test buttons

        # Instruments
        SrcLbl = wx.StaticText(self, label='V1 source (SRC):', id=wx.ID_ANY)
        self.Sources = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.SRC_COMBO_CHOICE,
                                   size=(150, 10), style=wx.CB_DROPDOWN)
        self.Sources.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_SRC.append(self.Sources)

        IP_DVM_Lbl = wx.StaticText(self, label='Input DVM (DVM12):',
                                   id=wx.ID_ANY)
        self.IP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.IP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.IP_Dvms)
        OP_DVM_Lbl = wx.StaticText(self, label='Output DVM (DVM3):',
                                   id=wx.ID_ANY)
        self.OP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.OP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.OP_Dvms)
        TDvmLbl = wx.StaticText(self, label='T-probe DVM (DVMT):',
                                id=wx.ID_ANY)
        self.TDvms = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.DVM_COMBO_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.TDvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.TDvms)

        GMHLbl = wx.StaticText(self, label='GMH probe (GMH):', id=wx.ID_ANY)
        self.GMHProbes = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GMH_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.GMHProbes.Bind(wx.EVT_COMBOBOX, self.BuildCommStr)
        self.cbox_instr_GMH.append(self.GMHProbes)

        GMHroomLbl = wx.StaticText(self,
                                   label='Room conds. GMH probe (GMHroom):',
                                   id=wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.GMH_COMBO_CHOICE,
                                         style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)

        IVboxLbl = wx.StaticText(self, label='IV_box (IVbox):', id=wx.ID_ANY)
        self.IVbox = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.IVBOX_COMBO_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.IVbox.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)

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
        FileLbl = wx.StaticText(self, label='Excel file full path:',
                                id=wx.ID_ANY)
        self.XLFile = wx.TextCtrl(self, id=wx.ID_ANY,
                                  value=self.GetTopLevelParent().ExcelPath,
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
        gbSizer.Add(SrcLbl, pos=(0, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.Sources, pos=(0, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(IP_DVM_Lbl, pos=(1, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.IP_Dvms, pos=(1, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(OP_DVM_Lbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.OP_Dvms, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(TDvmLbl, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.TDvms, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(GMHLbl, pos=(4, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHProbes, pos=(4, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(GMHroomLbl, pos=(5, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomProbes, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(IVboxLbl, pos=(6, 0), span=(1, 1),
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
        gbSizer.Add(FileLbl, pos=(8, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.XLFile, pos=(8, 1), span=(1, 5),
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
        devices.ROLES_WIDGETS = {'SRC': {'icb': self.Sources,
                                         'acb': self.SrcAddr,
                                         'tbtn': self.STest}}
        devices.ROLES_WIDGETS.update({'DVM12': {'icb': self.IP_Dvms,
                                                'acb': self.IP_DvmAddr,
                                                'tbtn': self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVM3': {'icb': self.OP_Dvms,
                                               'acb': self.OP_DvmAddr,
                                               'tbtn': self.D3Test}})
        devices.ROLES_WIDGETS.update({'DVMT': {'icb': self.TDvms,
                                               'acb': self.TDvmAddr,
                                               'tbtn': self.DTTest}})
        devices.ROLES_WIDGETS.update({'GMH': {'icb': self.GMHProbes,
                                              'acb': self.GMHPorts,
                                              'tbtn': self.GMHTest}})
        devices.ROLES_WIDGETS.update({'GMHroom': {'icb': self.GMHroomProbes,
                                                  'acb': self.GMHroomPorts,
                                                  'tbtn': self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'IVbox': {'icb': self.IVbox,
                                                'acb': self.IVboxAddr,
                                                'tbtn': self.IVboxTest}})

    def BuildComboChoices(self):
        for d in devices.INSTR_DATA.keys():
            if 'SRC:' in d:
                self.SRC_COMBO_CHOICE.append(d)
            elif 'DVM:' in d:
                self.DVM_COMBO_CHOICE.append(d)
            elif 'GMH:' in d:
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

    def UpdateFilepath(self, e):
        '''
        Called when a new Excel file has been selected.
        '''
        self.XLFile.SetValue(e.XLpath)

        # Open logfile
        logname = 'IVYv'+str(e.v)+'_'+str(dt.date.today())+'.log'
        logfile = os.path.join(e.d, logname)
        self.GetTopLevelParent().log = open(logfile, 'a')
        self.log = self.GetTopLevelParent().log

        # Read parameters sheet - gather instrument info:
        self.GetTopLevelParent().wb = load_workbook(self.XLFile.GetValue(), data_only = True) # Need cell VALUE, not FORMULA, so set data_only = True
        self.wb = self.GetTopLevelParent().wb
        self.ws_params = self.wb.get_sheet_by_name('Parameters')

        headings = (None, u'description', u'Instrument Info:',
                    u'parameter', u'value', u'uncert', u'dof', u'label')

        # Determine colummn indices from column letters:
        col_I = cell.cell.column_index_from_string('I') - 1
        col_J = cell.cell.column_index_from_string('J') - 1
        col_K = cell.cell.column_index_from_string('K') - 1
        col_L = cell.cell.column_index_from_string('L') - 1
        col_M = cell.cell.column_index_from_string('M') - 1
        col_N = cell.cell.column_index_from_string('N') - 1

        params = []
        values = []

        for r in self.ws_params.rows:  # a tuple of row objects
            descr = r[col_I].value  # cell.value
            param = r[col_J].value  # cell.value
            v_u_d_l = [r[col_K].value, r[col_L].value, r[col_M].value, r[col_N].value]  # value,uncert,dof,label

            if descr in headings and param in headings:
                continue  # Skip this row
            else:  # not header
                params.append(param)
                if v_u_d_l[1] is None:  # single-valued (no uncert)
                    values.append(v_u_d_l[0])  # append value as next item
                    print descr, ' : ', param, ' = ', v_u_d_l[0]
                    print >>self.log, descr, ' : ', param, ' = ', v_u_d_l[0]
                else:  # multi-valued
                    while v_u_d_l[-1] is None:  # remove empty cells
                        del v_u_d_l[-1]  # v_u_d_l.pop()
                    values.append(v_u_d_l)  # append value-list as next item
                    print descr, ' : ', param, ' = ', v_u_d_l
                    print >>self.log, descr, ' : ', param, ' = ', v_u_d_l

                if param == u'test':  # last parameter for this description
                    devices.DESCR.append(descr)  # build description list
                    devices.sublist.append(dict(zip(params, values)))
                    del params[:]
                    del values[:]

        print '----END OF PARAMETER LIST----'
        print >>self.log, '----END OF PARAMETER LIST----'

        # Compile into a dictionary that lives in devices.py...
        devices.INSTR_DATA = dict(zip(devices.DESCR, devices.sublist))
        self.BuildComboChoices()
        self.OnAutoPop(wx.EVT_BUTTON)  # Populate combo boxes immediately

    def OnAutoPop(self, e):
        '''
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        '''
        self.instrument_choice = {'SRC': 'SRC: F5520A',
                                  'DVM12': 'DVM: HP3458A, s/n518',
                                  'DVM3': 'DVM: HP3458A, s/n066',
                                  'DVMT': 'DVM: HP34401A, s/n976',
                                  'GMH': 'GMH: s/n627',
                                  'GMHroom': 'GMH: s/n367',
                                  'IVbox': 'IV_box'}
        for r in self.instrument_choice.keys():
            d = self.instrument_choice[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d)  # Update i_cb
            self.CreateInstr(d, r)
        if self.DUCName.GetValue() == u'DUC Name':
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
        # Called by both OnAutoPop() and UpdateInstr()
        # Create each instrument in software & open visa session (GPIB only)
        # For GMH instruments, use GMH dll, not visa

        if 'GMH' in r:  # Changed from d to r
            # create and open a GMH instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): \
            Creating GMH device (%s -> %s).' % (d, r)
            devices.ROLES_INSTR.update({r: devices.GMH_Sensor(d)})
        else:
            # create a visa instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): \
            Creating VISA device (%s -> %s).' % (d, r)
            devices.ROLES_INSTR.update({r: devices.instrument(d)})
            devices.ROLES_INSTR[r].Open()
        self.SetInstr(d, r)

    def SetInstr(self, d, r):
        """
        Called by CreateInstr().
        Updates internal info (INSTR_DATA) and Enables/disables testbuttons
        as necessary.
        """
        assert d in devices.INSTR_DATA, 'Unknown instrument: %s - \
        check Excel file is loaded.' % d
        assert 'role' in devices.INSTR_DATA[d], 'Unknown instrument parameter - \
        check Excel Parameters sheet is populated.'
        devices.INSTR_DATA[d]['role'] = r  # update default role

        # Set the address cb to correct value (according to devices.INSTR_DATA)
        a_cb = devices.ROLES_WIDGETS[r]['acb']
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
        print'UpdateAddr():', r, 'using', d, 'set to addr', addr, '(', a, ')'

    def OnTest(self, e):
        # Called when a 'test' button is clicked
        d = 'none'
        for r in devices.ROLES_WIDGETS.keys():  # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break  # stop looking when we've found right instr descr
        print'\nnbpages.SetupPage.OnTest():', d
        assert 'test' in devices.INSTR_DATA[d], 'No test exists \
        for this device.'
        test = devices.INSTR_DATA[d]['test']  # test string
        print '\tTest string:', test
        self.Response.SetValue(str(devices.ROLES_INSTR[r].Test(test)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d, test), 0)

    def OnIVBoxTest(self, e):
        resource = self.IVboxAddr.GetValue()
        config = str(devices.IVBOX_CONFIGS['V1'])
        try:
            instr = devices.RM.open_resource(resource)
            instr.write(config)
        except devices.visa.VisaIOError:
            self.Response.SetValue('Couldn\'t open visa resource for IV_box!')

    def BuildCommStr(self, e):
        # Called by a change in GMH probe selection, or DUC name
        d = e.GetString()
        if 'GMH' in d:  # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.CreateInstr(d, r)
        RunPage = self.GetParent().GetPage(1)
        params = {'DUC': self.DUCName.GetValue(), 'GMH': self.GMHProbes.GetValue()}
        joinstr = ' monitored by '
        commstr = 'IVY v.' + self.version + '. DUC: ' + params['DUC'] + joinstr + params['GMH']
        evt = evts.UpdateCommentEvent(str=commstr)
        wx.PostEvent(RunPage, evt)

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
        self.version = self.GetTopLevelParent().version
        self.run_id = 'none'

        self.GAINS_CHOICE = ['1e3', '1e4', '1e5', '1e6',
                             '1e7', '1e8', '1e9', '1e10']
        self.Rs_CHOICE = ['1k', '10k', '100k', '1M', '10M', '100M', '1G']
        self.Rs_SWITCHABLE = self.Rs_CHOICE[: 4]
        self.Rs_VALUES = [1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9]
        self.Rs_choice_to_val = dict(zip(self.Rs_CHOICE, self.Rs_VALUES))
        self.VNODE_CHOICE = ['V1', 'V2', 'V3']

        # Event bindings
        self.Bind(evts.EVT_UPDATE_COM_STR, self.UpdateComment)
        self.Bind(evts.EVT_DATA, self.UpdateData)
        self.Bind(evts.EVT_START_ROW, self.UpdateStartRow)

        self.RunThread = None

        # Comment widgets
        CommentLbl = wx.StaticText(self, id=wx.ID_ANY, label='Comment:')
        self.Comment = wx.TextCtrl(self, id=wx.ID_ANY, size=(600, 20))
        self.Comment.Bind(wx.EVT_TEXT, self.OnComment)
        comtip = 'This string is auto-generated from data on the Setup page.\
        Other notes may be added manually at the end.'
        self.Comment.SetToolTipString(comtip)

        self.NewRunIDBtn = wx.Button(self, id=wx.ID_ANY,
                                     label='Create new run id')
        idcomtip = 'Create new id to uniquely identify this set of \
        measurement data.'
        self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.OnNewRunID)
        self.RunID = wx.TextCtrl(self, id=wx.ID_ANY)  # size=(500,20)

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
        StartRowLbl = wx.StaticText(self, id=wx.ID_ANY, label='Start row:')
        self.StartRow = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)

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
        RowLbl = wx.StaticText(self, id=wx.ID_ANY, label='Current row:')
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
        gbSizer.Add(StartRowLbl, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        gbSizer.Add(self.StartRow, pos=(3, 5), span=(1, 1),
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

    def OnNewRunID(self, e):
        start = self.fullstr.find('DUC: ')
        end = self.fullstr.find(' monitored', start)
        DUCname = self.fullstr[start+4: end]
        self.run_id = str('IVY.v' + self.version + ' ' + DUCname + ' ' +
                          dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.status.SetStatusText('Id for subsequent runs:', 0)
        self.status.SetStatusText(str(self.run_id), 1)
        self.RunID.SetValue(str(self.run_id))

    def UpdateComment(self, e):
        # writes combined auto-comment and manual comment when
        # auto-generated comment is re-built
        self.autocomstr = e.str  # store copy of auto-generated comment
        self.Comment.SetValue(e.str+self.manstr)

    def OnComment(self, e):
        # Called when comment emits EVT_TEXT (i.e. whenever it's changed)
        # Prevent overwriting comment field (plus manually-entered notes)
        self.fullstr = self.Comment.GetValue()  # store a copy of full comment
        # Extract last part of comment (the manually-inserted bit)
        # - assume we manually added extra notes to END
        self.manstr = self.fullstr[len(self.autocomstr):]

    def UpdateData(self, e):
        # Triggered by an 'update data' event
        # event parameter is a dictionary:
        #  ud{'node:,'Vm':,'Vsd':,'time':,'row':,'Prog':,'end_flag':[0,1]}
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

    def UpdateStartRow(self, e):
        # Triggered by an 'update startrow' event
        self.StartRow.SetValue(str(e.row))

    def OnRs(self, e):
        self.Rs_val = self.Rs_choice_to_val[e.GetString()]  # an INT
        print '\nRunPage.OnRs(): Rs =', self.Rs_val
        if e.GetString() in self.Rs_SWITCHABLE:  # a STRING
            s = str(int(math.log10(self.Rs_val)))  # '3','4','5' or '6'
            print '\nSwitching Rs - Sending "%s" to IVbox' % s
            devices.ROLES_INSTR['IVbox'].SendCmd(s)

    def OnNode(self, e):
        node = e.GetString()  # 'V1', 'V2', or 'V3'
        print'\nRunPage.OnNode():', node
        s = node[1]
        if s in ('1', '2'):
            print'\nRunPage.OnNode():Sending IVbox "', s, '"'
            devices.ROLES_INSTR['IVbox'].SendCmd(s)
        else:  # '3'
            print'\nRunPage.OnNode():IGNORING IVbox cmd "', s, '"'

    def OnV1Set(self, e):
        # Called by change in value (manually OR by software!)
        V1 = e.GetValue()
        print'RunPage.OnV1Set(): V1 =',V1,'(',type(V1),')'
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
            src.SetV(0)
            src.Stby()
        else:
            self.V1Setting.SetValue('0')  # Calls OnV1Set() ONLY IF VAL CHANGES
            print'RunPage.OnZeroVolts():  Zero/Stby via V1 display'

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

        self.V3ax = self.figure.add_subplot(3, 1, 3)  # 3high x 1wide, 3rd plot down
        self.V3ax.ticklabel_format(style='sci', useOffset=False, axis='y',
                                   scilimits=(2, -2))  # Auto o/set to centre on data
        self.V3ax.yaxis.set_major_formatter(mtick.ScalarFormatter(useMathText=True, useOffset=False))  # Scientific notation.
        self.V3ax.autoscale(enable=True, axis='y',
                            tight=False)  # Autoscale with 'buffer' around data extents
        self.V3ax.set_xlabel('time')
        self.V3ax.set_ylabel('V3')

        self.V1ax = self.figure.add_subplot(3, 1, 1, sharex=self.V3ax)  # 3high x 1wide, 1st plot down
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

        self.V2ax = self.figure.add_subplot(3, 1, 2, sharex=self.V3ax)  # 3high x 1wide, 2nd plot down                            
        self.V2ax.ticklabel_format(useOffset=False,
                                   axis='y')  # Auto offset to centre on data
        self.V2ax.autoscale(enable=True, axis='y', tight=False)  # Autoscale with 'buffer' around data extents
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
        print e.node, 'len(V1)=', len(e.V12), 'len(V3)=', len(e.V3)
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
SEARCH_LIMIT = 24


class CalcPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.version = self.GetTopLevelParent().version
        self.Run_id = self.GetTopLevelParent().page2.run_id

        self.Rs_VALUES = self.GetTopLevelParent().page2.Rs_VALUES
        self.Rs_NAMES = ['Auto 1k', 'Auto 10k',
                         'Auto 100k', 'Auto 1M', 'Auto 10M',
                         'Auto 100M', 'Auto 1G']
        self.Rs_VAL_NAME = dict(zip(self.Rs_VALUES, self.Rs_NAMES))

        gbSizer = wx.GridBagSizer()

        # Analysis set-up:
        StartRowLbl = wx.StaticText(self,
                                    id=wx.ID_ANY, label='Data Start row:')
        gbSizer.Add(StartRowLbl, pos=(0, 0),
                    span=(1, 1), flag=wx.ALL | wx.EXPAND, border=5)
        self.StartRow = wx.TextCtrl(self, id=wx.ID_ANY,
                                    style=wx.TE_READONLY)  # TE_PROCESS_ENTER
        self.StartRow.Bind(wx.EVT_TEXT_ENTER, self.OnStartRow)
#        self.StartRow.SetToolTipString("Enter start row here BEFORE \
# clicking 'Analyze' button.")
        gbSizer.Add(self.StartRow, pos=(0, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        StopRowLbl = wx.StaticText(self, id=wx.ID_ANY,
                                   label='Data Stop row:')
        gbSizer.Add(StopRowLbl, pos=(0, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.StopRow = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.StopRow, pos=(0, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.Analyze = wx.Button(self, id=wx.ID_ANY, label='Analyze')
        self.Analyze.Bind(wx.EVT_BUTTON, self.OnAnalyze)
        gbSizer.Add(self.Analyze, pos=(0, 4), span=(1, 2),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, style=wx.LI_HORIZONTAL)
        gbSizer.Add(self.h_sep1, pos=(1, 0), span=(1, 6),
                    flag=wx.ALL | wx.EXPAND, border=5)

        # Analysis results:
        RangeLbl = wx.StaticText(self, id=wx.ID_ANY,
                                 label='Range or Gain (V/A):')
        gbSizer.Add(RangeLbl, pos=(2, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.Range = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.Range, pos=(3, 0), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        DeltaVLbl = wx.StaticText(self, id=wx.ID_ANY, label='O/P Delta-V:')
        gbSizer.Add(DeltaVLbl, pos=(2, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.DeltaV_01 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.DeltaV_01, pos=(3, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.DeltaV_1 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.DeltaV_1, pos=(4, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.DeltaV_10 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.DeltaV_10, pos=(5, 1), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        PosILbl = wx.StaticText(self, id=wx.ID_ANY, label='+I in (A):')
        gbSizer.Add(PosILbl, pos=(2, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.PosI_01 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.PosI_01, pos=(3, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.PosI_1 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.PosI_1, pos=(4, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.PosI_10 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.PosI_10, pos=(5, 2), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        NegILbl = wx.StaticText(self, id=wx.ID_ANY, label='-I in (A):')
        gbSizer.Add(NegILbl, pos=(2, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.NegI_01 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.NegI_01, pos=(3, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.NegI_1 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.NegI_1, pos=(4, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.NegI_10 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.NegI_10, pos=(5, 3), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        ExpULbl = wx.StaticText(self, id=wx.ID_ANY, label='Exp. U (A):')
        gbSizer.Add(ExpULbl, pos=(2, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.ExpU_01 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.ExpU_01, pos=(3, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.ExpU_1 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.ExpU_1, pos=(4, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.ExpU_10 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.ExpU_10, pos=(5, 4), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        CovFactLbl = wx.StaticText(self, id=wx.ID_ANY, label='k:')
        gbSizer.Add(CovFactLbl, pos=(2, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.CovFact_01 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.CovFact_01, pos=(3, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.CovFact_1 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.CovFact_1, pos=(4, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)
        self.CovFact_10 = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gbSizer.Add(self.CovFact_10, pos=(5, 5), span=(1, 1),
                    flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gbSizer)
        self.V01_widgets = [self.DeltaV_01, self.PosI_01, self.NegI_01,
                            self.ExpU_01, self.CovFact_01]
        self.V1_widgets = [self.DeltaV_1, self.PosI_1, self.NegI_1,
                           self.ExpU_1, self.CovFact_1]
        self.V10_widgets = [self.DeltaV_10, self.PosI_10, self.NegI_10,
                            self.ExpU_10, self.CovFact_10]
        self.Vout_widgets = {0.1: self.V01_widgets,
                             1: self.V1_widgets,
                             10: self.V10_widgets}

    def OnAnalyze(self, e):
        self.GetXL()

        self.Data_start_row = self.ws_Data['B1'].value
        print'Start row =', self.Data_start_row
        self.StartRow.SetValue(str(self.Data_start_row))

        self.Data_stop_row = self.GetStopRow()
        print'Stop row =', self.Data_stop_row
        self.StopRow.SetValue(str(self.Data_stop_row))

        # Set start row for next acquisition run:
        self.ws_Data['B1'].value = self.Data_stop_row + 4

        self.Results_start_row = self.ws_Results['B1'].value

        for V in [0.1, 1, 10]:
            for i in range(5):
                self.Vout_widgets[V][i].SetValue('')

        self.GetInstrAssignments()  # Result: self.role_descr
        self.GetParams()  # Result: self.I_INFO, self.R_INFO

        # Correction for Pt-100 sensor DVM:
        DVMT_cor = self.I_INFO[self.role_descr['DVMT']]['correction_100r']

        '''
        Pt sensor is a few cm away from input resistors, so assume a
        fairly large type B Tdef of 0.1 deg C:
        '''
        Pt_T_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                             3, label='Pt_T_def')
        Pt_alpha = self.R_INFO['Pt 100r']['alpha']
        Pt_beta = self.R_INFO['Pt 100r']['beta']
        Pt_R0 = self.R_INFO['Pt 100r']['R0_LV']
        Pt_TRef = self.R_INFO['Pt 100r']['TRef_LV']

        '''
        GMH sensor is a few cm away from DUC which, itself, has a size of
        several cm, so assume a fairly large type B Tdef of 0.1 deg C:
        '''
        GMH_T_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](0.1),
                              3, label='GMH_T_def')

        Comment = self.ws_Data['A'+str(self.Data_start_row)].value
        Run_Id = self.ws_Data['B'+str(self.Data_start_row-2)].value
        DUC_name = self.GetNamefromComment(Comment)
        DUC_gain = self.ws_Data['B'+str(self.Data_start_row)].value
        self.Range.SetValue(str('{0:.2e}'.format(DUC_gain)))
        Mean_date = self.GetMeanDate()

        print'Comment:', Comment
        print'Run_Id:', Run_Id
        print'gain =', DUC_gain
        print 'Mean_date:', Mean_date

        # Determine mean env. conditions
        GMH_Ts = []
        GMHroom_RHs = []
        GMHroom_Ps = []
        row = self.Data_start_row
        del GMH_Ts[:]
        del GMHroom_RHs[:]
        del GMHroom_Ps[:]
        while row <= self.Data_stop_row:
            GMH_Ts.append(self.ws_Data['L'+str(row)].value)
            GMHroom_RHs.append(self.ws_Data['R'+str(row)].value)
            GMHroom_Ps.append(self.ws_Data['Q'+str(row)].value)
            row += 1

        d = self.role_descr['GMH']
        T_GMH_cor = self.I_INFO[d]['T_correction']  # ppm, multiplicative, ureal
        T_GMH_raw = GTC.ta.estimate_digitized(GMH_Ts, 0.01)
        T_GMH = T_GMH_raw*(1 + T_GMH_cor) + GMH_T_def

        d = self.role_descr['GMHroom']
        RH_cor = self.I_INFO[d]['RH_correction']
        RH_raw = GTC.ta.estimate_digitized(GMHroom_RHs, 0.1)
        RH = RH_raw*(1 + RH_cor)

        # Re-use d (same instrument description)
        P_cor = self.I_INFO[d]['P_correction']
        P_raw = GTC.ta.estimate_digitized(GMHroom_Ps, 0.1)
        P = P_raw*(1 + P_cor)

        self.result_row = self.Write_Summary(Comment, Run_Id, DUC_name,
                                             DUC_gain, Mean_date, T_GMH, RH, P)

        influencies = []
        V1s = []
        V2s = []
        V3s = []
        row = self.Data_start_row
        while row < self.Data_stop_row:
            gains = set()
            # 'neg' and 'pos' refer to polarity of OUTPUT VOLTAGE, not
            # input current!
            neg_nom_Vout = self.ws_Data['G'+str(row+1)].value
            pos_nom_Vout = self.ws_Data['G'+str(row+2)].value
            abs_nom_Vout = pos_nom_Vout

            # Construct ureals from raw voltage data, including gain correction
            for n in range(4):
                label_suffix_1 = self.ws_Data['D'+str(row+n)].value+'_'+str(n)
                label_suffix_2 = self.ws_Data['D'+str(row+4+n)].value+'_'+str(n)
                label_suffix_3 = 'V3' + '_' + str(n)

                V1_v = self.ws_Data['J'+str(row+n)].value
                V1_u = self.ws_Data['K'+str(row+n)].value
                V1_d = self.ws_Data['F'+str(row+n)].value - 1
                V1_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_1
                d1 = self.role_descr['DVM12']
                gain_param = self.get_gain_err_param(V1_v)
                gain = self.I_INFO[d1][gain_param]
                gains.add(gain)
                V1_raw = GTC.ureal(V1_v, V1_u, V1_d, label=V1_l)
                V1s.append(GTC.result(V1_raw/gain))

                V2_v = self.ws_Data['J'+str(row+4+n)].value
                V2_u = self.ws_Data['K'+str(row+4+n)].value
                V2_d = self.ws_Data['F'+str(row+4+n)].value - 1
                V2_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_2
                d2 = self.role_descr['DVM12']
                gain_param = self.get_gain_err_param(V2_v)
                gain = self.I_INFO[d2][gain_param]
                gains.add(gain)
                V2_raw = GTC.ureal(V2_v, V2_u, V2_d, label=V2_l)
                V2s.append(GTC.result(V2_raw/gain))

                V3_v = self.ws_Data['H'+str(row+n)].value
                V3_u = self.ws_Data['I'+str(row+n)].value
                V3_d = self.ws_Data['F'+str(row+n)].value - 1
                V3_l = 'OP'+str(abs_nom_Vout)+'_'+label_suffix_3
                d3 = self.role_descr['DVM3']
                gain_param = self.get_gain_err_param(V3_v)
                gain = self.I_INFO[d3][gain_param]
                gains.add(gain)
                V3_raw = GTC.ureal(V3_v, V3_u, V3_d, label=V3_l)
                V3s.append(GTC.result(V3_raw/gain))

                GMH_Ts.append(self.ws_Data['L'+str(row)].value)
                GMH_Ts.append(self.ws_Data['L'+str(row+4)].value)
                influencies.extend([V1_raw, V2_raw, V3_raw])

            influencies.extend(list(gains))  # A list of unique gain corrections - no copies.
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

            # Rs Temperature
            T_Rs = []
            Pt_R_cor = []
            for r in range(8):
                Pt_R_raw = self.ws_Data['M'+str(row+r)].value
                Pt_R_cor.append(GTC.result(Pt_R_raw * (1 + DVMT_cor),
                                           label='Pt_Rcor'+str(r)))
                T_Rs.append(GTC.result(self.R_to_T(Pt_alpha, Pt_beta,
                                                      Pt_R_cor[r],
                                                      Pt_R0, Pt_TRef)))

            av_T_Rs = GTC.result(GTC.fn.mean(T_Rs),
                                    label='av_T_Rs'+str(abs_nom_Vout))
            influencies.extend(Pt_R_cor)
            influencies.extend([Pt_alpha, Pt_beta, Pt_R0, Pt_TRef,
                                DVMT_cor, Pt_T_def])  # av_T_Rs
            assert Pt_alpha in influencies,'Pt_alpha missing from influencies!'
            assert Pt_beta in influencies,'Pt_beta missing from influencies!'
            assert Pt_R0 in influencies,'Pt_R0 missing from influencies!'
            assert Pt_TRef in influencies,'Pt_TRef missing from influencies!'
            assert DVMT_cor in influencies,'DVMT_cor missing from influencies!'
            assert Pt_T_def in influencies,'Pt_T_def missing from influencies!'

            # Value of Rs
            nom_Rs = self.ws_Data['C'+str(row)].value
            print '\nNominal Rs value:', nom_Rs, 'Abs. Nom. Vout:\
', abs_nom_Vout, '\n'
            Rs_name = self.Rs_VAL_NAME[nom_Rs]
            Rs_0 = self.R_INFO[Rs_name]['R0_LV']  # a ureal
            Rs_TRef = self.R_INFO[Rs_name]['TRef_LV']  # a ureal
            Rs_alpha = self.R_INFO[Rs_name]['alpha']
            Rs_beta = self.R_INFO[Rs_name]['beta']

            # Correct Rs value for temperature
            dT = GTC.result(av_T_Rs - Rs_TRef + Pt_T_def)
            Rs = GTC.result(Rs_0*(1 + Rs_alpha*dT + Rs_beta*dT**2))

            influencies.extend([Rs_0, Rs_alpha, Rs_beta, Rs_TRef])

            '''
            Finally, calculate current-change in,
            for nominal voltage-change out:
            '''
            Iin_pos = GTC.result(V_Rs_pos/Rs)
            Iin_neg = GTC.result(V_Rs_neg/Rs)


            I_pos = GTC.result(Iin_pos*pos_nom_Vout / V3_pos)
            I_pos_k = GTC.rp.k_factor(I_pos.df)  # P = 95% by default
            I_pos_EU = I_pos_k * I_pos.u
            I_neg = GTC.result(Iin_neg*neg_nom_Vout/V3_neg)

            self.Vout_widgets[abs_nom_Vout][0].SetValue(str(abs_nom_Vout))
            self.Vout_widgets[abs_nom_Vout][1].SetValue('{0:.8g}'.format(I_pos.x))
            self.Vout_widgets[abs_nom_Vout][2].SetValue('{0:.8g}'.format(I_neg.x))
            # Just display positive value for now:
            self.Vout_widgets[abs_nom_Vout][3].SetValue('{0:.3g}'.format(I_pos_EU))
            self.Vout_widgets[abs_nom_Vout][4].SetValue(str(round(I_pos_k)))

            this_result = {'Vout': abs_nom_Vout, 'I_pos': I_pos,
                           'I_neg': I_neg}

            # build uncertainty budget table
            budget_table_pos = []
            budget_table_neg = []
            for i in influencies:
                print'Working through influence variables:', i.label
                if i.u == 0:
                    sensitivity_pos = sensitivity_neg = 0
                else:
                    sensitivity_pos = GTC.component(I_pos, i)/i.u
                    sensitivity_neg = GTC.component(I_neg, i)/i.u
                # Only include non-zero influencies:
                if abs(GTC.component(I_pos, i)) > 0:
                    print 'Included component of I+:',GTC.component(I_pos, i)
                    budget_table_pos.append([i.label, i.x, i.u, i.df,
                                             sensitivity_pos,
                                             GTC.component(I_pos, i)])
                else:
                    print'ZERO COMPONENT of I+'
                if abs(GTC.component(I_neg, i)) > 0:
                    print 'Included component of I_neg',GTC.component(I_neg, i)
                    budget_table_neg.append([i.label, i.x, i.u, i.df,
                                             sensitivity_neg,
                                             GTC.component(I_neg, i)])
                else:
                    print'ZERO COMPONENT of I-'
            self.budget_table_pos_sorted = sorted(budget_table_pos,
                                                  key=self.by_u_cont,
                                                  reverse=True)
            self.budget_table_neg_sorted = sorted(budget_table_neg,
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

    def GetXL(self):
        '''
        NOTE: Details of the Excel file are not available
        until the user has opened it!
        '''
        self.XLPath = self.GetTopLevelParent().ExcelPath
        print '\n', self.XLPath
        assert self.XLPath is not "", 'No data file open yet!'

        self.ws_Data = self.GetTopLevelParent().wb.get_sheet_by_name('Data')
        self.ws_Params = self.GetTopLevelParent().wb.get_sheet_by_name('Parameters')
        self.ws_Results = self.GetTopLevelParent().wb.get_sheet_by_name('Results')

    def OnStartRow(self, e):
        self.GetXL()
        self.ws_Data['B1'].value = int(e.GetString())

    def GetStopRow(self):
        row = self.Data_start_row
        self.Test_Vs = []
        '''
        Don't search forever and
        ignore final row:
        '''
        while row < self.Data_start_row + SEARCH_LIMIT - 1:
            NomVOP = self.ws_Data['G'+str(row)].value
            if NomVOP in (None, 'Nom. Vout '):  # Ran out of data
                break
            elif NomVOP in (0.1, 1, 10):
                self.Test_Vs.append(NomVOP)
                row += 1
                continue
            else:  # in (0,-0.1, -1, -10)
                row += 1
                continue
        Test_V_set = set(self.Test_Vs)
        if len(Test_V_set) < 1:
            print'GetStopRow(): Incomplete data! - ', self.Test_Vs
            return self.Data_start_row
        else:
            print 'GetStopRow(): Test Vs:', Test_V_set
            return self.Data_start_row + 4 * len(self.Test_Vs) - 1

    def GetNamefromComment(self, c):
        return c[c.find('DUC: ') + 5: c.find(' monitored by GMH')]

    def GetMeanDate(self):
        r = self.Data_start_row
        n = 0
        t_av = 0.0
        while r <= self.Data_stop_row:
            s = self.ws_Data['E'+str(r)].value  # A unicode str
            # Convert s to a Python datetime object:
            t_dt = dt.datetime.strptime(s, '%d/%m/%Y %H:%M:%S')
            t_tup = dt.datetime.timetuple(t_dt)  # A Python time tuple object
            t_av += time.mktime(t_tup)  # time as float (seconds from epoch)
            r += 1
            n += 1
        t_av /= n
        t_av_dt = dt.datetime.fromtimestamp(t_av)
        return t_av_dt.strftime('%d/%m/%Y %H:%M:%S')  # av. time as string

    def Write_Summary(self, Comment, Run_Id, DUC_name,
                      DUC_gain, date, T, RH, P):
        '''
        Write Run summary and result column-headings.
        Return next row
        '''
        DELTA = u'\N{GREEK CAPITAL LETTER DELTA}'
        row = self.Results_start_row
        proc_date = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        proc_string = 'Processesed by IVY v{} on {}'.format(self.version, proc_date)
        self.ws_Results['H'+str(row)].value = proc_string
        self.ws_Results['A'+str(row)].value = 'Comment:'
        self.ws_Results['B'+str(row)].value = Comment
        self.ws_Results['A'+str(row+1)].value = 'Run Id:'
        self.ws_Results['B'+str(row+1)].value = Run_Id
        self.ws_Results['A'+str(row+2)].value = 'Date:'
        self.ws_Results['B'+str(row+2)].value = date
        self.ws_Results['A'+str(row+3)].value = 'DUC Name:'
        self.ws_Results['B'+str(row+3)].value = DUC_name
        self.ws_Results['A'+str(row+4)].value = 'Gain (V/A):'
        self.ws_Results['B'+str(row+4)].value = DUC_gain

        self.ws_Results['C'+str(row+2)].value = 'Condition:'
        self.ws_Results['C'+str(row+3)].value = 'Value:'
        self.ws_Results['C'+str(row+4)].value = 'Exp Uncert.:'
        self.ws_Results['C'+str(row+5)].value = 'Cov. factor:'

        T_k = GTC.rp.k_factor(T.df)
        self.ws_Results['D'+str(row+2)].value = 'T (GMH)'
        self.ws_Results['D'+str(row+3)].value = T.x
        self.ws_Results['D'+str(row+4)].value = T_k*T.u
        self.ws_Results['D'+str(row+5)].value = T_k

        RH_k = GTC.rp.k_factor(RH.df)
        self.ws_Results['E'+str(row+2)].value = 'RH (%)'
        self.ws_Results['E'+str(row+3)].value = RH.x
        self.ws_Results['E'+str(row+4)].value = RH_k*RH.u
        self.ws_Results['E'+str(row+5)].value = RH_k

        P_k = GTC.rp.k_factor(P.df)
        self.ws_Results['F'+str(row+2)].value = 'P (mBar)'
        self.ws_Results['F'+str(row+3)].value = P.x
        self.ws_Results['F'+str(row+4)].value = P_k * P.u
        self.ws_Results['F'+str(row+5)].value = P_k
        # Add blank line below summary
        self.ws_Results['H'+str(row+6)].value = 'Uncertainty Budget:'
        self.ws_Results['A'+str(row+7)].value = 'Nom. ' + DELTA + 'V'
        self.ws_Results['B'+str(row+7)].value = DELTA + 'I in'
        self.ws_Results['C'+str(row+7)].value = 'Std. u'
        self.ws_Results['D'+str(row+7)].value = 'dof'
        self.ws_Results['E'+str(row+7)].value = 'Exp. U'
        self.ws_Results['F'+str(row+7)].value = 'k'
        self.ws_Results['H'+str(row+7)].value = 'Quantity (label)'
        self.ws_Results['I'+str(row+7)].value = 'Value'
        self.ws_Results['J'+str(row+7)].value = 'Std. u'
        self.ws_Results['K'+str(row+7)].value = 'dof'
        self.ws_Results['L'+str(row+7)].value = 'Sens. Co.'
        self.ws_Results['M'+str(row+7)].value = 'Uncert. Cont.'

        return row+8

    def GetInstrAssignments(self):
        N_ROLES = 7  # 10 roles in total
        self.role_descr = {}
        for row in range(self.Data_start_row, self.Data_start_row + N_ROLES):
            # Read {role:description}
            temp_dict = {self.ws_Data['S' + str(row)].value: self.ws_Data['T' + str(row)].value}
            assert temp_dict.keys()[-1] is not None, 'Instrument assignment: Missing role!'
            assert temp_dict.values()[-1] is not None, 'Instrument assignment: Missing description!'
            self.role_descr.update(temp_dict)

    def Uncertainize(self, items):
        '''
        Convert a list of data to a ureal, where possible.
        Expects items to be a list: [value, uncert, dof, label].
        If uncert is missing or value is non-numeric return value.
        Otherwise, return a ureal (with or without default dof)
        '''
        v = items[0]
        if len(items) < 4:
            return v
        u = items[1]
        d = items[2]
        l = items[3]
        if (u is not None) and isinstance(v, Number):
            if d == u'inf':
                un_num = GTC.ureal(v, u, label=l)  # default dof = inf
            else:
                un_num = GTC.ureal(v, u, d, l)
            return un_num
        else:  # non-numeric value or not enough info to make a ureal
            return v

    def GetParams(self):
        '''
        Extract resistor and instrument parameters
        '''
        print '\nReading parameters...'
        # log.write('Reading parameters...')
        headings = (u'Resistor Info:', u'Instrument Info:',
                    u'description', u'parameter', u'value',
                    u'uncert', u'dof', u'label', u'Comment / Reference')

        # Determine colummn indices from column letters:
        col_A = cell.cell.column_index_from_string('A') - 1
        col_B = cell.cell.column_index_from_string('B') - 1
        col_C = cell.cell.column_index_from_string('C') - 1
        col_D = cell.cell.column_index_from_string('D') - 1
        col_E = cell.cell.column_index_from_string('E') - 1
        col_F = cell.cell.column_index_from_string('F') - 1
        col_G = cell.cell.column_index_from_string('G') - 1

        col_I = cell.cell.column_index_from_string('I') - 1
        col_J = cell.cell.column_index_from_string('J') - 1
        col_K = cell.cell.column_index_from_string('K') - 1
        col_L = cell.cell.column_index_from_string('L') - 1
        col_M = cell.cell.column_index_from_string('M') - 1
        col_N = cell.cell.column_index_from_string('N') - 1
        col_O = cell.cell.column_index_from_string('O') - 1

        R_params = []
        R_row_items = []
        I_params = []
        I_row_items = []
        R_values = []
        I_values = []
        R_DESCR = []
        I_DESCR = []
        R_sublist = []
        I_sublist = []

        for r in self.ws_Params.rows:  # a tuple of row objects
            R_end = 0

            # description, parameter, value, uncert, dof, label:
            R_row_items = [r[col_A].value, r[col_B].value, r[col_C].value,
                           r[col_D].value, r[col_E].value, r[col_F].value,
                           r[col_G].value]

            I_row_items = [r[col_I].value, r[col_J].value, r[col_K].value,
                           r[col_L].value, r[col_M].value, r[col_N].value,
                           r[col_O].value]

            if R_row_items[0] == None:  # end of R_list
                R_end = 1

            # check this row for heading text
            if any(i in I_row_items for i in headings):
                continue  # Skip headings

            else:  # not header - main data
                # Get instrument parameters first...
                '''
                Need to know last row if we write more data, post-analysis:
                '''
                last_I_row = r[col_I].row
                I_params.append(I_row_items[1])
                I_values.append(self.Uncertainize(I_row_items[2:6]))
                '''
                'test' is always the last parameter for each
                instrument description:
                '''
                if I_row_items[1] == u'test':
                    I_DESCR.append(I_row_items[0])  # build description list
                    # Add parameter dictionary to sublist:
                    I_sublist.append(dict(zip(I_params, I_values)))
                    del I_params[:]
                    del I_values[:]

                # Now attend to resistor parameters...
                if R_end == 0:  # If not at end of resistor data-block
                    '''
                    Need to know last row if we write more data, post-analysis:
                    '''
                    last_R_row = r[col_A].row
                    R_params.append(R_row_items[1])
                    R_values.append(self.Uncertainize(R_row_items[2:6]))
                    '''
                    'T_sensor' is always the last parameter for each
                    resistor description:
                    '''
                    if R_row_items[1] == u'T_sensor':
                        R_DESCR.append(R_row_items[0])  # build descr list
                        # Add parameter dictionary to sublist:
                        R_sublist.append(dict(zip(R_params, R_values)))
                        del R_params[:]
                        del R_values[:]

        """
        Compile into dictionaries:
        There are two dictionaries; one for instruments (I_INFO) and one for
        resistors (R_INFO). Each dictionary item is keyed by the description
        (name) of the instrument (resistor). Each dictionary value is itself
        a dictionary, keyed by parameter, such as 'address' (for an instrument)
        or 'R_LV' (for a resistor value, measured at 'low voltage').
        """
        self.I_INFO = dict(zip(I_DESCR, I_sublist))
        print len(self.I_INFO), 'instruments (%d rows)' % last_I_row

        self.R_INFO = dict(zip(R_DESCR, R_sublist))
        print len(self.R_INFO), 'resistors.(%d rows)\n' % last_R_row

    def get_gain_err_param(self, V):
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
        # Convert a resistive T-sensor reading from resistance to temperature
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
        Write results and uncert. budget for nom.Vout (BOTH polarities)
        '''
        r = self.result_row
        print'WriteThisResult(): Starting result_row =', r
        sh = self.ws_Results

        # Positive results 1st..
        sh['A'+str(r)].value = result['Vout']
        sh['B'+str(r)].value = result['I_pos'].x
        sh['C'+str(r)].value = result['I_pos'].u
        sh['D'+str(r)].value = result['I_pos'].df
        k = GTC.rp.k_factor(result['I_pos'].df)
        sh['E'+str(r)].value = k*result['I_pos'].u
        sh['F'+str(r)].value = k

        for line in self.budget_table_pos_sorted:
            sh['H'+str(r)] = line[0]  # Quantity (label)
            sh['I'+str(r)] = line[1]  # Value
            sh['J'+str(r)] = line[2]  # Uncert.
            if math.isinf(line[3]):
                sh['K'+str(r)] = str(line[3])  # dof
            else:
                sh['K'+str(r)] = round(line[3])  # dof
            sh['L'+str(r)] = line[4]  # Sens. coef.
            sh['M'+str(r)] = line[5]  # Uncert. contrib.
            r += 1

        r += 1  # Blank line between polarities

        # ...then negative results...
        sh['A'+str(r)].value = -1*result['Vout']
        sh['B'+str(r)].value = result['I_neg'].x
        sh['C'+str(r)].value = result['I_neg'].u
        sh['D'+str(r)].value = result['I_neg'].df
        k = GTC.rp.k_factor(result['I_neg'].df)
        sh['E'+str(r)].value = k*result['I_neg'].u
        sh['F'+str(r)].value = k

        for line in self.budget_table_neg_sorted:
            sh['H'+str(r)] = line[0]  # Quantity (label)
            sh['I'+str(r)] = line[1]  # Value
            sh['J'+str(r)] = line[2]  # Uncert.
            if math.isinf(line[3]):
                sh['K'+str(r)] = str(line[3])  # dof
            else:
                sh['K'+str(r)] = round(line[3])  # dof
            sh['L'+str(r)] = line[4]  # Sens. coef.
            sh['M'+str(r)] = line[5]  # Uncert. contrib.
            r += 1

        print'WriteThisResult(): Final result_row =', r
        self.ws_Results['B1'].value = r+1
        return r+1  # Blank line between results
