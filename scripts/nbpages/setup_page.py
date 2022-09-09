# -*- coding: utf-8 -*-
# setup_page.py
""" Defines individual setup page as a panel-like object,
for inclusion in a wx.Notebook object

PYTHON3 DEVELOPMENT VERSION

Created on Fri Mar 5 13:38:16 2021

@author: t.lawson
"""

import wx
from wx.lib.masked import NumCtrl
import logging
from scripts import devices, IVY_events as Evts

logger = logging.getLogger(__name__)


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
                                  'DVM12': 'DVM_3458A:s/n382',
                                  'DVM3': 'DVM_3458A:s/n452',
                                  'DVMT': 'DVM_34401A:s/n976',
                                  'GMH': 'GMH:s/n530',
                                  'GMHroom': 'GMH:s/n367'}
        self.T_SENSOR_CHOICE = devices.T_Sensors  # 'none', 'Pt', 'SR104t', 'thermistor'
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
        SrcLbl = wx.StaticText(self, label='V1 source (SRC):',
                               id=wx.ID_ANY)
        self.Sources = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.SRC_COMBO_CHOICE,
                                   size=(150, 10), style=wx.CB_DROPDOWN)
        self.Sources.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_SRC.append(self.Sources)

        IP_DVM_Lbl = wx.StaticText(self, label='Input DVM (DVM12):',
                                   id=wx.ID_ANY)
        self.IP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.IP_Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.IP_Dvms)
        OP_DVM_Lbl = wx.StaticText(self, label='Output DVM (DVM3):',
                                   id=wx.ID_ANY)
        self.OP_Dvms = wx.ComboBox(self, wx.ID_ANY,
                                   choices=self.DVM_COMBO_CHOICE,
                                   style=wx.CB_DROPDOWN)
        self.OP_Dvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.OP_Dvms)
        TDvmLbl = wx.StaticText(self, label='T-probe DVM (DVMT):',
                                id=wx.ID_ANY)
        self.TDvms = wx.ComboBox(self, wx.ID_ANY,
                                 choices=self.DVM_COMBO_CHOICE,
                                 style=wx.CB_DROPDOWN)
        self.TDvms.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_DVM.append(self.TDvms)

        GMHLbl = wx.StaticText(self, label='GMH probe (GMH):',
                               id=wx.ID_ANY)
        self.GMHProbes = wx.ComboBox(self, wx.ID_ANY,
                                     choices=self.GMH_COMBO_CHOICE,
                                     style=wx.CB_DROPDOWN)
        self.GMHProbes.Bind(wx.EVT_COMBOBOX, self.build_comm_str)
        self.cbox_instr_GMH.append(self.GMHProbes)

        GMHroomLbl = wx.StaticText(self,
                                   label='Room conds. GMH probe (GMHroom):',
                                   id=wx.ID_ANY)
        self.GMHroomProbes = wx.ComboBox(self, wx.ID_ANY,
                                         choices=self.GMH_COMBO_CHOICE,
                                         style=wx.CB_DROPDOWN)
        self.GMHroomProbes.Bind(wx.EVT_COMBOBOX, self.update_instr)
        self.cbox_instr_GMH.append(self.GMHroomProbes)

        IVboxLbl = wx.StaticText(self, label='IV_box (IVbox) setting:',
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
                                   style=wx.TE_READONLY | wx.TE_MULTILINE)  # | wx.TE_MULTILINE

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
        gb_sizer.Add(SrcLbl, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Sources, pos=(0, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(IP_DVM_Lbl, pos=(1, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.IP_Dvms, pos=(1, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(OP_DVM_Lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.OP_Dvms, pos=(2, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(TDvmLbl, pos=(3, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.TDvms, pos=(3, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(GMHLbl, pos=(4, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHProbes, pos=(4, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(GMHroomLbl, pos=(5, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.GMHroomProbes, pos=(5, 1), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(IVboxLbl, pos=(6, 0), span=(1, 1),
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

        gb_sizer.Add(response_lbl, pos=(4, 4), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.Response, pos=(5, 4), span=(1, 3),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.VisaList, pos=(0, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        gb_sizer.Add(self.ResList, pos=(0, 4), span=(4, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Autopopulate btn
        gb_sizer.Add(self.AutoPop, pos=(1, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.SetSizerAndFit(gb_sizer)

        # Associate roles and corresponding comboboxes/test btns here:
        devices.ROLES_WIDGETS = {'SRC': {'lbl': SrcLbl,
                                         'icb': self.Sources,
                                         'acb': self.SrcAddr,
                                         'tbtn': self.STest}}
        devices.ROLES_WIDGETS.update({'DVM12': {'lbl': IP_DVM_Lbl,
                                                'icb': self.IP_Dvms,
                                                'acb': self.IP_DvmAddr,
                                                'tbtn': self.D12Test}})
        devices.ROLES_WIDGETS.update({'DVM3': {'lbl': OP_DVM_Lbl,
                                               'icb': self.OP_Dvms,
                                               'acb': self.OP_DvmAddr,
                                               'tbtn': self.D3Test}})
        devices.ROLES_WIDGETS.update({'DVMT': {'lbl': TDvmLbl,
                                               'icb': self.TDvms,
                                               'acb': self.TDvmAddr,
                                               'tbtn': self.DTTest}})
        devices.ROLES_WIDGETS.update({'GMH': {'lbl': GMHLbl,
                                              'icb': self.GMHProbes,
                                              'acb': self.GMHPorts,
                                              'tbtn': self.GMHTest}})
        devices.ROLES_WIDGETS.update({'GMHroom': {'lbl': GMHroomLbl,
                                                  'icb': self.GMHroomProbes,
                                                  'acb': self.GMHroomPorts,
                                                  'tbtn': self.GMHroomTest}})
        devices.ROLES_WIDGETS.update({'IVbox': {'lbl': IVboxLbl,
                                                'icb': self.IVbox,
                                                'acb': self.IVboxAddr,
                                                'tbtn': self.IVboxTest}})

        # # Create IV-box instrument once:
        # d = 'IV_box'  # Description
        # r = 'IVbox'  # Role
        # self.create_instr(d, r)
        #
        # self.build_combo_choices()

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

        # (No choices for IV box - there's only one)

    def update_dir(self, e):  # Triggered by Evts.EVT_FILEPATH
        """
        Display working directory once selected.
        """
        self.WorkingDir.SetValue(e.Dir)

    def on_auto_pop(self, e):  # Triggered by wx.EVT_BUTTON from Autopop btn
        """
        Pre-select instrument and address comboboxes -
        Choose from instrument descriptions listed in devices.DESCR
        (Uses address assignments in devices.INSTR_DATA)
        """
        for r in self.INSTRUMENT_CHOICE.keys():
            d = self.INSTRUMENT_CHOICE[r]
            devices.ROLES_WIDGETS[r]['icb'].SetValue(d)  # Update instr. cbox
            a = devices.INSTR_DATA[d]['str_addr']
            devices.ROLES_WIDGETS[r]['acb'].SetValue(a)  # Update addr. cbox
            self.create_instr(d, r)
        if self.DUCName.GetValue() == u'DUC Name':
            self.DUCName.SetForegroundColour((255, 0, 0))  # red
            self.DUCName.SetValue('CHANGE_THIS!')

    def update_instr(self, e):  # Triggered by wx.EVT_COMBOBOX (instruments)
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
        logger.info('CreateInstr({0:s},{1:s})...'.format(d, r))
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

    @staticmethod  # THIS MAY BELONG ELSEWHERE!!!
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

    def on_IV_box_test(self, e):
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
