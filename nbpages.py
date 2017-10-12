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
matplotlib.use('WXAgg') # Agg renderer for drawing on a wx canvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
#from matplotlib.backends.backend_wx import NavigationToolbar2Wx
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick
from openpyxl import load_workbook, cell

import IVY_events as evts
import acquisition as acq
#import RLink as rl
import devices

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
        self.GMH_COMBO_CHOICE = ['none'] # devices.GMH_DESCR # ('GMH s/n628', 'GMH s/n627')
#        self.SB_COMBO_CHOICE =  devices.SWITCH_CONFIGS.keys()
        self.IVBOX_COMBO_CHOICE = ['IV_box','none']
        self.T_SENSOR_CHOICE = devices.T_Sensors
        self.cbox_addr_COM = []
        self.cbox_addr_GPIB = []
        self.cbox_instr_SRC = []
        self.cbox_instr_DVM = []
        self.cbox_instr_GMH = []

        self.BuildComboChoices()

        self.GMH1Addr = self.GMH2Addr = 0 # invalid initial address as default

        self.ResourceList = []
        self.ComList = []
        self.GPIBList = []
        self.GPIBAddressList = ['addresses','GPIB0::0'] # dummy values for starters...
        self.COMAddressList = ['addresses','COM0'] # dummy values for starters...

        self.test_btns = [] # list of test buttons

        # Instruments
        SrcLbl = wx.StaticText(self, label='V1 source (SRC):', id = wx.ID_ANY)
        self.Sources = wx.ComboBox(self,wx.ID_ANY, choices = self.SRC_COMBO_CHOICE, size = (150,10), style = wx.CB_DROPDOWN)
        self.Sources.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_SRC.append(self.Sources)
        
        IP_DVM_Lbl = wx.StaticText(self, label='Input DVM (DVM12):', id = wx.ID_ANY)
        self.IP_Dvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.IP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.IP_Dvms)
        OP_DVM_Lbl = wx.StaticText(self, label='Output DVM (DVM3):', id = wx.ID_ANY)
        self.OP_Dvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.OP_Dvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.OP_Dvms)
        TDvmLbl = wx.StaticText(self, label='T-probe DVM (DVMT):', id = wx.ID_ANY)
        self.TDvms = wx.ComboBox(self,wx.ID_ANY, choices = self.DVM_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.TDvms.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_DVM.append(self.TDvms)
        
        GMHLbl = wx.StaticText(self, label='GMH probe (GMH):', id = wx.ID_ANY)
        self.GMHProbes = wx.ComboBox(self,wx.ID_ANY, choices = self.GMH_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.GMHProbes.Bind(wx.EVT_COMBOBOX, self.BuildCommStr)
        self.cbox_instr_GMH.append(self.GMHProbes)
        
        GMHroomLbl = wx.StaticText(self, label='Room conds. GMH probe (GMHroom):', id = wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self,wx.ID_ANY, choices = self.GMH_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)
        
        IVboxLbl = wx.StaticText(self, label='IVbox (IVbox):', id = wx.ID_ANY)
        self.IVbox = wx.ComboBox(self,wx.ID_ANY, choices = self.IVBOX_COMBO_CHOICE, style=wx.CB_DROPDOWN)
        self.IVbox.Bind(wx.EVT_COMBOBOX, self.UpdateInstr)

        # Addresses
        self.SrcAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, size = (150,10), style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.SrcAddr)
        self.SrcAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.IP_DvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.IP_DvmAddr)
        self.IP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.OP_DvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.OP_DvmAddr)
        self.OP_DvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.TDvmAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.GPIBAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_GPIB.append(self.TDvmAddr)
        self.TDvmAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.GMHPorts = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHPorts)
        self.GMHPorts.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.GMHroomPorts = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.GMHroomPorts)
        self.GMHroomPorts.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)
        
        self.IVboxAddr = wx.ComboBox(self,wx.ID_ANY, choices = self.COMAddressList, style=wx.CB_DROPDOWN)
        self.cbox_addr_COM.append(self.IVboxAddr)
        self.IVboxAddr.Bind(wx.EVT_COMBOBOX, self.UpdateAddr)

        # Filename
        FileLbl = wx.StaticText(self, label='Excel file full path:', id = wx.ID_ANY)
        self.XLFile = wx.TextCtrl(self, id = wx.ID_ANY, value=self.GetTopLevelParent().ExcelPath)
        
        # DUC
        self.DUCName = wx.TextCtrl(self, id = wx.ID_ANY, value= 'DUC Name')
        self.DUCName.Bind(wx.EVT_TEXT, self.BuildCommStr)

        # Autopopulate btn
        self.AutoPop = wx.Button(self,id = wx.ID_ANY, label='AutoPopulate')
        self.AutoPop.Bind(wx.EVT_BUTTON, self.OnAutoPop)
        
        # Test buttons
        self.VisaList = wx.Button(self,id = wx.ID_ANY, label='List Visa res')
        self.VisaList.Bind(wx.EVT_BUTTON, self.OnVisaList)
        self.ResList = wx.TextCtrl(self, id = wx.ID_ANY, value = 'Available Visa resources',
                                   style = wx.TE_READONLY|wx.TE_MULTILINE)
                                   
        self.STest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.STest.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.D12Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.D12Test.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.D3Test = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.D3Test.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.DTTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.DTTest.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.GMHTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.GMHTest.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.GMHroomTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.GMHroomTest.Bind(wx.EVT_BUTTON, self.OnTest)
        
        self.IVboxTest = wx.Button(self,id = wx.ID_ANY, label='Test')
        self.IVboxTest.Bind(wx.EVT_BUTTON, self.OnIVBoxTest)
        
        ResponseLbl = wx.StaticText(self, label='Instrument Test Response:', id = wx.ID_ANY)
        self.Response = wx.TextCtrl(self, id = wx.ID_ANY, value= '', style = wx.TE_READONLY)
        
        gbSizer = wx.GridBagSizer()

        # Instruments
        gbSizer.Add(SrcLbl, pos=(0,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Sources, pos=(0,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(IP_DVM_Lbl, pos=(1,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.IP_Dvms, pos=(1,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)      
        gbSizer.Add(OP_DVM_Lbl, pos=(2,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.OP_Dvms, pos=(2,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(TDvmLbl, pos=(3,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.TDvms, pos=(3,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(GMHLbl, pos=(4,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHProbes, pos=(4,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(GMHroomLbl, pos=(5,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomProbes, pos=(5,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(IVboxLbl, pos=(6,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.IVbox, pos=(6,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        # Addresses
        gbSizer.Add(self.SrcAddr, pos=(0,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.IP_DvmAddr, pos=(1,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.OP_DvmAddr, pos=(2,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.TDvmAddr, pos=(3,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHPorts, pos=(4,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomPorts, pos=(5,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)    
        gbSizer.Add(self.IVboxAddr, pos=(6,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        # DUC Name
        gbSizer.Add(self.DUCName, pos=(6,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        # Filename
        gbSizer.Add(FileLbl, pos=(8,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.XLFile, pos=(8,1), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        
        # Test buttons
        gbSizer.Add(self.STest, pos=(0,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.D12Test, pos=(1,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.D3Test, pos=(2,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.DTTest, pos=(3,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHTest, pos=(4,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.GMHroomTest, pos=(5,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.IVboxTest, pos=(6,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        gbSizer.Add(ResponseLbl, pos=(3,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Response, pos=(4,4), span=(1,3), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.VisaList, pos=(0,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.ResList, pos=(0,4), span=(3,1), flag=wx.ALL|wx.EXPAND, border=5)

        # Autopopulate btn
        gbSizer.Add(self.AutoPop, pos=(2,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)

        self.SetSizerAndFit(gbSizer)

        # Roles and corresponding comboboxes/test btns are associated here:
        devices.ROLES_WIDGETS = {'SRC':{'icb':self.Sources,'acb':self.SrcAddr,'tbtn':self.STest}}
        devices.ROLES_WIDGETS.update({'DVM12':{'icb':self.IP_Dvms,'acb':self.IP_DvmAddr,'tbtn':self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVM3':{'icb':self.OP_Dvms,'acb':self.OP_DvmAddr,'tbtn':self.D3Test}})
        devices.ROLES_WIDGETS.update({'DVMT':{'icb':self.TDvms,'acb':self.TDvmAddr,'tbtn':self.DTTest}})
        devices.ROLES_WIDGETS.update({'GMH':{'icb':self.GMHProbes,'acb':self.GMHPorts,'tbtn':self.GMHTest}})
        devices.ROLES_WIDGETS.update({'GMHroom':{'icb':self.GMHroomProbes,'acb':self.GMHroomPorts,'tbtn':self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'IVbox':{'icb':self.IVbox,'acb':self.IVboxAddr,'tbtn':self.IVboxTest}})


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
        self.log = open(logfile,'a')
        
        # Read parameters sheet - gather instrument info:
        self.wb = load_workbook(self.XLFile.GetValue(),data_only = True) # Need cell VALUE, not FORMULA, so set data_only = True
        self.ws_params = self.wb.get_sheet_by_name('Parameters')
        
        headings = (None, u'description',u'Instrument Info:',u'parameter',u'value',u'uncert',u'dof',u'label')
        
        # Determine colummn indices from column letters:
        col_I = cell.cell.column_index_from_string('I') - 1
        col_J = cell.cell.column_index_from_string('J') - 1
        col_K = cell.cell.column_index_from_string('K') - 1
        col_L = cell.cell.column_index_from_string('L') - 1
        col_M = cell.cell.column_index_from_string('M') - 1
        col_N = cell.cell.column_index_from_string('N') - 1

        params = []
        values = []
        
        for r in self.ws_params.rows: # a tuple of row objects
            descr = r[col_I].value # cell.value
            param = r[col_J].value # cell.value
            v_u_d_l = [r[col_K].value, r[col_L].value, r[col_M].value, r[col_N].value] # value,uncert,dof,label
        
            if descr in headings and param in headings:
                continue # Skip this row
            else: # not header
                params.append(param)
                if v_u_d_l[1] is None: # single-valued (no uncert)
                    values.append(v_u_d_l[0]) # append value as next item
                    print descr,' : ',param,' = ',v_u_d_l[0]
                    print >>self.log, descr,' : ',param,' = ',v_u_d_l[0]
                else: # multi-valued
                    while v_u_d_l[-1] is None: # remove empty cells
                        del v_u_d_l[-1] # v_u_d_l.pop()
                    values.append(v_u_d_l) # append value-list as next item 
                    print descr,' : ',param,' = ',v_u_d_l
                    print >>self.log, descr,' : ',param,' = ',v_u_d_l
                
                if param == u'test': # last parameter for this description
                    devices.DESCR.append(descr) # build description list
                    devices.sublist.append(dict(zip(params,values))) # adds parameter dictionary to sublist
                    del params[:]
                    del values[:] 

        print '----END OF PARAMETER LIST----' 
        print >>self.log, '----END OF PARAMETER LIST----'
        
        # Compile into a dictionary that lives in devices.py...  
        devices.INSTR_DATA = dict(zip(devices.DESCR,devices.sublist))
        self.BuildComboChoices()
        self.OnAutoPop(wx.EVT_BUTTON) # Populate combo boxes immediately


    def OnAutoPop(self, e):
        '''
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        '''
        self.instrument_choice = {'SRC':'SRC: D4808',
                                  'DVM12':'DVM: HP3458A, s/n452',
                                  'DVM3':'DVM: HP3458A, s/n518',
                                  'DVMT':'DVM: HP34401A, s/n976',
                                  'GMH':'GMH: s/n627',
                                  'GMHroom':'GMH: s/n367',
                                  'IVbox':'IV_box'}
        for r in self.instrument_choice.keys():
            d = self.instrument_choice[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d) # Update i_cb
            self.CreateInstr(d,r)
        if self.DUCName.GetValue() == u'DUC Name':
            self.DUCName.SetValue('CHANGE_THIS!')
        


    def UpdateInstr(self, e):
        '''
        An instrument was selected for a role.
        Find description d and role r, then pass to CreatInstr()
        '''
        d = e.GetString()
        for r in devices.ROLES_WIDGETS.keys(): # Cycle through roles
            if devices.ROLES_WIDGETS[r]['icb'] == e.GetEventObject():
                break # stop looking when we've found the right instrument & role
        self.CreateInstr(d,r)


    def CreateInstr(self,d,r):
        # Called by both OnAutoPop() and UpdateInstr()
        # Create each instrument in software & open visa session (for GPIB instruments)
        # For GMH instruments, use GMH dll not visa

        if 'GMH' in r: # Changed from d to r
            # create and open a GMH instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): Creating GMH device (%s -> %s).'%(d,r)
            devices.ROLES_INSTR.update({r:devices.GMH_Sensor(d)})
        else:
            # create a visa instrument instance
            print'\nnbpages.SetupPage.CreateInstr(): Creating VISA device (%s -> %s).'%(d,r)
            devices.ROLES_INSTR.update({r:devices.instrument(d)})
            devices.ROLES_INSTR[r].Open()
        self.SetInstr(d,r)


    def SetInstr(self,d,r):
        """
        Called by CreateInstr().
        Updates internal info (INSTR_DATA) and Enables/disables testbuttons as necessary.
        """
#        print 'nbpages.SetupPage.SetInstr():',d,'assigned to role',r,'demo mode:',devices.ROLES_INSTR[r].demo
        assert devices.INSTR_DATA.has_key(d),'Unknown instrument: %s - check Excel file is loaded.'%d
        assert devices.INSTR_DATA[d].has_key('role'),'Unknown instrument parameter - check Excel Parameters sheet is populated.'
        devices.INSTR_DATA[d]['role'] = r # update default role
        
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
        acb = e.GetEventObject() # 'a'ddress 'c'ombo 'b'ox
        for r in devices.ROLES_WIDGETS.keys():
            if devices.ROLES_WIDGETS[r]['acb'] == acb:
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break # stop looking when we've found the right instrument description
        
        # ...Now change INSTR_DATA...
        a = e.GetString() # address string, eg 'COM5' or 'GPIB0::23'
        if (a not in self.GPIBAddressList) or (a not in self.COMAddressList): # Ignore dummy values, like 'NO_ADDRESS'
            devices.INSTR_DATA[d]['str_addr'] = a
            devices.ROLES_INSTR[r].str_addr = a
            addr = a.lstrip('COMGPIB0:') # leave only numeric part of address string
            devices.INSTR_DATA[d]['addr'] = int(addr)
            devices.ROLES_INSTR[r].addr = int(addr)
        print'UpdateAddr():',r,'using',d,'set to addr',addr,'(',a,')'
            

    def OnTest(self, e):
        # Called when a 'test' button is clicked
        d = 'none'
        for r in devices.ROLES_WIDGETS.keys(): # check every role
            if devices.ROLES_WIDGETS[r]['tbtn'] == e.GetEventObject():
                d = devices.ROLES_WIDGETS[r]['icb'].GetValue()
                break # stop looking when we've found the right instrument description
        print'\nnbpages.SetupPage.OnTest():',d
        assert devices.INSTR_DATA[d].has_key('test'), 'No test exists for this device.'
        test = devices.INSTR_DATA[d]['test'] # test string
        print '\tTest string:',test
        self.Response.SetValue(str(devices.ROLES_INSTR[r].Test(test)))
        self.status.SetStatusText('Testing %s with cmd %s' % (d,test),0)


    def OnIVBoxTest(self, e):
        resource = self.IVboxAddr.GetValue()
        config = str(devices.IVBOX_CONFIGS['V1'])
        try:
            instr = devices.RM.open_resource(resource)
            instr.write(config)
        except devices.visa.VisaIOError:
            self.Response.SetValue('Couldn\'t open visa resource for IV_box!')


    def BuildCommStr(self,e):
        # Called by a change in GMH probe selection, or DUC name
        d = e.GetString()
        if 'GMH' in d: # A GMH probe selection changed
            # Find the role associated with the selected instrument description
            for r in devices.ROLES_WIDGETS.keys():
                if devices.ROLES_WIDGETS[r]['icb'].GetValue() == d:
                    break
            # Update our knowledge of role <-> instr. descr. association
            self.CreateInstr(d,r)
        RunPage = self.GetParent().GetPage(1)
        params={'DUC':self.DUCName.GetValue(),'GMH':self.GMHProbes.GetValue()}
        joinstr = ' monitored by '
        commstr = 'IVY v.' + self.version + '. DUC: ' + params['DUC'] + joinstr + params['GMH']
        evt = evts.UpdateCommentEvent(str = commstr)
        wx.PostEvent(RunPage,evt)


    def OnVisaList(self, e):
        res_list = devices.RM.list_resources()
        del self.ResourceList[:] # list of COM ports ('COM X') & GPIB addresses
        del self.ComList[:] # list of COM ports (numbers only)
        del self.GPIBList[:] # list of GPIB addresses (numbers only)
        for item in res_list:
            self.ResourceList.append(item.replace('ASRL','COM'))
        for item in self.ResourceList:
            addr = item.replace('::INSTR','')
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
        
        self.GAINS_CHOICE = ['1e3','1e4','1e5','1e6','1e7','1e8','1e9','1e10']
        self.Rs_CHOICE = ['1k','10k','100k','1M','10M','100M','1G']
        self.Rs_SWITCHABLE = self.Rs_CHOICE[:4]
        self.Rs_VALUES = [1e3,1e4,1e5,1e6,1e7,1e8,1e9]
        self.Rs_choice_to_val = dict(zip(self.Rs_CHOICE,self.Rs_VALUES))
        self.VNODE_CHOICE = ['V1','V2','V3']
        
        # Event bindings
        self.Bind(evts.EVT_UPDATE_COM_STR, self.UpdateComment)
        self.Bind(evts.EVT_DATA, self.UpdateData)
        self.Bind(evts.EVT_START_ROW, self.UpdateStartRow)

        self.RunThread = None

        # Comment widgets
        CommentLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Comment:')
        self.Comment = wx.TextCtrl(self, id = wx.ID_ANY, size=(600,20))
        self.Comment.Bind(wx.EVT_TEXT,self.OnComment)
        comtip = 'This string is auto-generated from data on the Setup page. Other notes may be added manually at the end.'
        self.Comment.SetToolTipString(comtip)
        
        self.NewRunIDBtn = wx.Button(self, id = wx.ID_ANY, label='Create new run id')
        idcomtip = 'Create new id to uniquely identify this set of measurement data.'
        self.NewRunIDBtn.SetToolTipString(idcomtip)
        self.NewRunIDBtn.Bind(wx.EVT_BUTTON, self.OnNewRunID)
        self.RunID = wx.TextCtrl(self, id = wx.ID_ANY) # size=(500,20)

        # Run Setup widgets
        DUCgainLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_LEFT, label = 'DUC gain (V/A):')
        self.DUCgain = wx.ComboBox(self,wx.ID_ANY, choices = self.GAINS_CHOICE, style=wx.CB_DROPDOWN)
        RsLbl = wx.StaticText(self,id = wx.ID_ANY, style = wx.ALIGN_LEFT, label = 'I/P Rs:')
        self.Rs = wx.ComboBox(self,wx.ID_ANY, choices = self.Rs_CHOICE, style=wx.CB_DROPDOWN)
        self.Rs.Bind(wx.EVT_COMBOBOX, self.OnRs)
        SettleDelLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Settle delay:')
        self.SettleDel = wx.SpinCtrl(self,id = wx.ID_ANY,value ='0', min = 0, max=600)
        SrcLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_LEFT, label = 'V1 Setting:')
        self.V1Setting = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=8, groupDigits=True)
        self.V1Setting.Bind(wx.lib.masked.EVT_NUM, self.OnV1Set)
        ZeroVoltsBtn = wx.Button(self, id = wx.ID_ANY, label='Set zero volts', size = (200,20))
        ZeroVoltsBtn.Bind(wx.EVT_BUTTON, self.OnZeroVolts)
        StartRowLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Start row:')
        self.StartRow = wx.TextCtrl(self, id = wx.ID_ANY, style=wx.TE_READONLY) #, size = (20,20)
        
        self.h_sep1 = wx.StaticLine(self,id = wx.ID_ANY,style = wx.LI_HORIZONTAL)
        
        #  Run control and progress widgets
        self.StartBtn = wx.Button(self, id = wx.ID_ANY, label='Start run')
        self.StartBtn.Bind(wx.EVT_BUTTON, self.OnStart)
        self.StopBtn = wx.Button(self, id = wx.ID_ANY, label='Abort run')
        self.StopBtn.Bind(wx.EVT_BUTTON, self.OnAbort)
        self.StopBtn.Enable(False)
        NodeLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Node:')
        self.Node = wx.ComboBox(self,wx.ID_ANY, choices = self.VNODE_CHOICE, style=wx.CB_DROPDOWN)
        VavLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Mean V:')
        self.Vav = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=9, groupDigits=True)
        VsdLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Stdev V:')
        self.Vsd = NumCtrl(self, id = wx.ID_ANY, integerWidth=3, fractionWidth=9, groupDigits=True)
        TimeLbl = wx.StaticText(self,id = wx.ID_ANY, label = 'Timestamp:')
        self.Time = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY, size = (200,20))
        RowLbl =  wx.StaticText(self,id = wx.ID_ANY, label = 'Current row:')
        self.Row = wx.TextCtrl(self, id = wx.ID_ANY, style = wx.TE_READONLY) #, size = (20,20)
        ProgressLbl = wx.StaticText(self,id = wx.ID_ANY, style=wx.ALIGN_RIGHT, label = 'Run progress:')
        self.Progress = wx.Gauge(self,id = wx.ID_ANY,range=100, name='Progress')

        gbSizer = wx.GridBagSizer()

        # Comment widgets
        gbSizer.Add(CommentLbl,pos=(0,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Comment, pos=(0,1), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.NewRunIDBtn, pos=(1,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.RunID, pos=(1,1), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)

        # Run setup widgets
        gbSizer.Add(DUCgainLbl, pos=(2,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.DUCgain, pos=(3,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(RsLbl, pos=(2,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Rs, pos=(3,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(SettleDelLbl, pos=(2,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.SettleDel, pos=(3,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(SrcLbl,pos=(2,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.V1Setting,pos=(3,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(ZeroVoltsBtn, pos=(3,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(StartRowLbl, pos=(2,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StartRow, pos=(3,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        
        gbSizer.Add(self.h_sep1, pos=(4,0), span=(1,6), flag=wx.ALL|wx.EXPAND, border=5)
        
        #  Run control and progress widgets
        gbSizer.Add(self.StartBtn, pos=(5,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.StopBtn, pos=(6,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(NodeLbl, pos=(5,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Node, pos=(6,1), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(VavLbl, pos=(5,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Vav, pos=(6,2), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(VsdLbl, pos=(5,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Vsd, pos=(6,3), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(TimeLbl, pos=(5,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Time, pos=(6,4), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(RowLbl, pos=(5,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Row, pos=(6,5), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(ProgressLbl, pos=(7,0), span=(1,1), flag=wx.ALL|wx.EXPAND, border=5)
        gbSizer.Add(self.Progress, pos=(7,1), span=(1,5), flag=wx.ALL|wx.EXPAND, border=5)
        
        self.SetSizerAndFit(gbSizer)

        self.autocomstr = ''
        self.manstr = ''
        
        
    def OnNewRunID(self,e):
        start = self.fullstr.find('DUC: ')
        end = self.fullstr.find(' monitored',start)
        DUCname = self.fullstr[start+4:end]
        self.run_id = str('IVY.v' + self.version + ' ' + DUCname +  ' ' + dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.status.SetStatusText('Id for subsequent runs:',0)
        self.status.SetStatusText(str(self.run_id),1)
        self.RunID.SetValue(str(self.run_id))


    def UpdateComment(self,e):
        # writes combined auto-comment and manual comment when
        # auto-generated comment is re-built
        self.autocomstr = e.str # store a copy of automtically-generated comment
        self.Comment.SetValue(e.str+self.manstr)


    def OnComment(self,e):
        # Called when comment emits EVT_TEXT (i.e. whenever it's changed)
        # Make sure comment field (with extra manually-entered notes) isn't overwritten
        self.fullstr = self.Comment.GetValue() # store a copy of full comment
        # Extract last part of comment (the manually-inserted bit)
        # - assume we manually added extra notes to END
        self.manstr = self.fullstr[len(self.autocomstr):]


    def UpdateData(self,e):
        # Triggered by an 'update data' event
        # event parameter is a dictionary: ud{'node:,'Vm':,'Vsd':,'time':,'row':,'Prog':,'end_flag':[0,1]}
        if e.ud.has_key('node'):
            self.Node.SetValue(str(e.ud['node']))
        if e.ud.has_key('Vm'):
            self.Vav.SetValue(str(e.ud['Vm']))   
        if e.ud.has_key('Vsd'):
            self.Vsd.SetValue(str(e.ud['Vsd']))
        if e.ud.has_key('time'):
            self.Time.SetValue(str(e.ud['time']))
        if e.ud.has_key('row'):
            self.Row.SetValue(str(e.ud['row']))
        if e.ud.has_key('progress'):
            self.Progress.SetValue(e.ud['progress'])
        if e.ud.has_key('end_flag'): # Aborted or Finished
            self.RunThread = None
            self.StartBtn.Enable(True)

#        if e.flag in 'AF':# Aborted or Finished
#            self.RunThread = None
#            self.StartBtn.Enable(True)
#        else:
#            self.Node.SetValue(e.flag)
#            self.Time.SetValue(str(e.t))
#            self.Vav.SetValue(str(e.Vm))
#            self.Vsd.SetValue(str(e.Vsd))
#            self.Row.SetValue(str(e.r))
#            self.Progress.SetValue(e.P)


    def UpdateStartRow(self,e):
        # Triggered by an 'update startrow' event
        self.StartRow.SetValue(str(e.row))


    def OnRs(self,e):
        self.Rs_val = self.Rs_choice_to_val[e.GetString()] # an INT
        print 'RunPage.OnRs(): Rs =',self.Rs_val
        if e.GetString() in self.Rs_SWITCHABLE: # a STRING
            s = str(math.log10(round(self.Rs_val))) # '3','4','5' or '6'
            print 'Switching Rs - Sending %s to IVbox'%s
            devices.ROLES_INSTR['IVbox'].SendCmd(s)
        

    def OnV1Set(self,e):
        # Called by change in value (manually OR by software!)
        V1 = e.GetValue()
        src = devices.ROLES_INSTR['SRC']
        src.SetV(V1)
        time.sleep(0.5)
        if V1 == 0:
            src.Stby()
        else:
            src.Oper()
        time.sleep(0.5)


    def OnZeroVolts(self,e):
        # V1:
        src = devices.ROLES_INSTR['SRC']
        if self.V1Setting.GetValue() == 0:
            print'RunPage.OnZeroVolts(): Zero/Stby directly (not via V1 display)'
            src.SetV(0)
            src.Stby()
        else:
            self.V1Setting.SetValue('0') # Calls OnV1Set() ONLY IF VALUE CHANGES
            print'RunPage.OnZeroVolts():  Zero/Stby via V1 display'


    def OnStart(self,e):
        self.Progress.SetValue(0)
        self.RunThread = None
        self.status.SetStatusText('',1)
        self.status.SetStatusText('Starting run',0)
        if self.RunThread is None:
            self.StopBtn.Enable(True) # Enable Stop button
            self.StartBtn.Enable(False) # Disable Start button
            # start acquisition thread here
            self.RunThread = acq.AqnThread(self)


    def OnAbort(self,e):
        self.StartBtn.Enable(True)
        self.StopBtn.Enable(False) # Disable Stop button
        self.RunThread._want_abort = 1 #.abort


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

        self.Bind(evts.EVT_PLOT,self.UpdatePlot)
        self.Bind(evts.EVT_CLEARPLOT,self.ClearPlot)

        self.figure = Figure()

        self.figure.subplots_adjust(hspace = 0.3) # 0.3" height space between subplots
        
        self.V3ax = self.figure.add_subplot(3,1,3) # 3high x 1wide, 3rd plot down 
        self.V3ax.ticklabel_format(style='sci', useOffset=False, axis='y', scilimits=(2,-2)) # Auto offset to centre on data
        self.V3ax.yaxis.set_major_formatter(mtick.ScalarFormatter(useMathText=True, useOffset=False)) # Scientific notation .
        self.V3ax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        self.V3ax.set_xlabel('time')
        self.V3ax.set_ylabel('V3')

        self.V1ax = self.figure.add_subplot(3,1,1, sharex=self.V3ax) # 3high x 1wide, 1st plot down 
        self.V1ax.ticklabel_format(useOffset=False, axis='y') # Auto offset to centre on data
        self.V1ax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        plt.setp(self.V1ax.get_xticklabels(), visible=False) # Hide x-axis labels
        self.V1ax.set_ylabel('V1')
        self.V1ax.set_ylim(auto=True)
        V1_y_ost = self.V1ax.get_xaxis().get_offset_text()
        V1_y_ost.set_visible(False)

        self.V2ax = self.figure.add_subplot(3,1,2, sharex=self.V3ax) # 3high x 1wide, 2nd plot down 
        self.V2ax.ticklabel_format(useOffset=False, axis='y') # Auto offset to centre on data
        self.V2ax.autoscale(enable=True, axis='y', tight=False) # Autoscale with 'buffer' around data extents
        plt.setp(self.V2ax.get_xticklabels(), visible=False) # Hide x-axis labels
        self.V2ax.set_ylabel('V2')
        self.V2ax.set_ylim(auto=True)
        V2_y_ost = self.V2ax.get_xaxis().get_offset_text()
        V2_y_ost.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)


    def UpdatePlot(self, e):
        print'PlotPage.UpdatePlot(): len(t)=',len(e.t)
        print e.node,'len(V1)=',len(e.V12),'len(V3)=',len(e.V3)
        if e.node == 'V1':
            self.V1ax.plot_date(e.t, e.V12, 'bo')
        else: # V2 data
            self.V2ax.plot_date(e.t, e.V12, 'go')
        self.V3ax.plot_date(e.t, e.V3, 'ro')
        self.figure.autofmt_xdate() # default settings
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
