# -*- coding: utf-8 -*-
"""
devices.py

Gathers together all info required to use external instruments.

All device data is collected in a 'dictionary of dictionaries' - INSTR_DATA
Each piece of information is accessed as:
INSTR_DATA[<instrument description>][<parameter>].
E.g. the initiation string for the HP3458(s/n452) is:
INSTR_DATA['DVM_3458A:s/n452']['init_str'],
which retrieves the string 'DCV AUTO; NPLC 20'.
Note that some parameters are LISTS of strings (i.e. multiple commands)

Created on Fri Mar 17 13:52:15 2017

@author: t.lawson
"""

import os
import pyvisa as visa  # Deprecation warning - replace 'import visa' with: 'import pyvisa as visa'
# import logging
import json
from scripts.GMHstuff import GMHstuff as GMH

# logger = logging.getLogger(__name__)

ROLES_WIDGETS = {}
"""Dictionary of GUI widgets keyed by role."""

ROLES_INSTR = {}
"""Dictionary of GMH_sensor or Instrument objects, keyed by role."""

RES_DATA = {}
"""Dictionary of known resistance standards."""

INSTR_DATA = {}
"""Dictionary of instrument parameter dictionaries."""

RED = (255, 0, 0)
GREEN = (0, 127, 0)

"""
The following are generally-useful utility functions and code to
ensure information about resistors and instruments are available:
"""


def strip_chars(oldstr, charlist=''):
    """
    Strip characters from oldstr and return newstr that does not
    contain any of the characters in charlist.

    Note that the built-in str.strip() only removes characters from
    the start and end (not throughout).
    """
    newstr = ''
    for ch in charlist:
        newstr = ''.join(oldstr.split(ch))
        oldstr = newstr
    return newstr


def refresh_params(directory):
    """
    Refreshes state-of-knowledge of resistors and instruments.

    Returns:
    res_data: Dictionary of known resistance standards.
    instr_data: Dictionary of instrument parameter dictionaries.

    Both are keyed by description and used to update global dictionaries
    RES_DATA and INSTR_DATA.
    """
    resistor_file = r'data\IVY_Resistors.json'
    with open(os.path.join(directory, resistor_file), 'r') as new_resistor_fp:
        resistor_str = strip_chars(new_resistor_fp.read(), '\t\n')  # Remove tabs & newlines
    res_data = json.loads(resistor_str)

    instrument_file = r'data\IVY_Instruments.json'
    with open(os.path.join(directory, instrument_file), 'r') as new_instr_fp:
        instr_str = strip_chars(new_instr_fp.read(), '\t\n')
    instr_data = json.loads(instr_str)

    return res_data, instr_data


"""
VISA-specific stuff:
Only ONE VISA resource manager is required at any time -
Communications for all GPIB and RS232 instruments (except GMH)
are handled by RM.
"""
RM = visa.ResourceManager()  # 'C:\\WINDOWS\\system32\\visa32.dll'
print('Available visa resources:\n\t', RM.list_resources())

# Switchbox
IVBOX_CONFIGS = {'V1': '1', 'V2': '2', '1k': '3', '10k': '4', '100k': '5',
                 '1M': '6', '10M': '7', '100M': '8', '1G': '9'}

T_Sensors = ('none', 'Pt', 'SR104t', 'thermistor')


class GMHSensor(GMH.GMHSensor):
    """
    A derived class of GMHstuff.GMHSensor with additional functionality.
    On creation, an instance needs a description string, descr.
    """
    def __init__(self, descr, role):
        self.descr = descr
        self.role = role
        self.port = int(INSTR_DATA[self.descr]['addr'])
        super().__init__(self.port)
        self.demo = True

    def test(self, meas):
        """
        Test that the device is functioning.
        :argument meas (str) - an alias for the measurement type:
            'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.
        :returns measurement tuple: (<value>, <unit string>)
        """
        # print('\ndevices.GMH_Sensor.Test()...')
        self.open_port()
        reply = self.measure(meas)
        self.set_power_off_time(120)  # Ensure sensor stays on during whole run.
        self.close()
        return reply

    def init(self):
        pass


'''
###############################################################################
'''


class Device(object):
    """
    A generic external device or instrument
    """
    def __init__(self, demo=True):
        self.demo = demo

    def open(self):
        pass

    def close(self):
        pass


class Instrument(Device):
    """
    A class for associating instrument data with a VISA instance of
    that instrument
    """
    def __init__(self, descr, role, demo=True):  # Default to demo mode
        self.instr = None
        self.descr = descr
        self.demo = demo
        self.is_open = 0
        self.is_operational = 0

        assert_msg = 'Unknown instrument ({0:s})'.format(self.descr)
        # check instrument data is loaded from Excel Parameters sheet.'
        assert self.descr in INSTR_DATA, assert_msg

        self.addr = INSTR_DATA[self.descr]['addr']
        self.str_addr = INSTR_DATA[self.descr]['str_addr']
        self.role = role

        if 'init_str' in INSTR_DATA[self.descr]:
            self.InitStr = INSTR_DATA[self.descr]['init_str']  # a str
            # print(f'{self.descr} - init_str: "{self.InitStr}"')
        else:
            self.InitStr = ''  # empty string
        if 'setfn_str' in INSTR_DATA[self.descr]:
            self.SetFnStr = INSTR_DATA[self.descr]['setfn_str']
        else:
            self.SetFnStr = ''  # an empty string
        if 'oper_str' in INSTR_DATA[self.descr]:
            self.OperStr = INSTR_DATA[self.descr]['oper_str']
        else:
            self.OperStr = ''  # an empty string
        if 'stby_str' in INSTR_DATA[self.descr]:
            self.StbyStr = INSTR_DATA[self.descr]['stby_str']
        else:
            self.StbyStr = ''
        if 'chk_err_str' in INSTR_DATA[self.descr]:
            self.ChkErrStr = INSTR_DATA[self.descr]['chk_err_str']
        else:
            self.ChkErrStr = ('',)
        if 'setV_str' in INSTR_DATA[self.descr]:
            self.VStr = INSTR_DATA[self.descr]['setV_str']  # a tuple or list of str
        else:
            self.VStr = ('',)  # a tuple of empty strings

    def open(self):
        msg_head = 'devices.instrument.Open(): {}'
        try:
            self.instr = RM.open_resource(self.str_addr)
            if '3458A' in self.descr:
                self.instr.read_termination = '\r\n'
                self.instr.write_termination = '\r\n'
            self.instr.timeout = 2000  # default 2 s timeout
            INSTR_DATA[self.descr]['demo'] = False  # A real working instrument
            self.demo = False  # A real instrument ONLY on Open() success
            green = (0, 255, 0)
            ROLES_WIDGETS[self.role]['lbl'].SetBackgroundColour(green)
            # print(msg_head.format(f'{self.descr} session handle={self.instr.session}.'))
            # logger.info(msg_head.format(f'{self.descr} session handle={self.instr.session}.'))
            self.is_open = 1
        except visa.VisaIOError:
            self.demo = True  # default to demo mode if can't open
            self.instr = None
            red = (255, 0, 0)
            ROLES_WIDGETS[self.role]['lbl'].SetForegroundColour(red)
            ROLES_WIDGETS[self.role]['lbl'].Refresh()
            INSTR_DATA[self.descr]['demo'] = True
            print(msg_head.format(f'failed: {self.descr} opened in demo mode'))
            # logger.warning(msg_head.format(f'failed: {self.descr} opened in demo mode'))
        return self.instr

    def close(self):
        # Close comms with instrument
        msg_head = 'devices.instrument.Close(): {}'
        if self.demo is True:
            print(msg_head.format(f'{self.descr} in demo mode - nothing to close.'))
            # logger.info(msg_head.format(f'{self.descr} in demo mode - nothing to close.'))
        elif self.instr is not None:
            # print(msg_head.format(f'Closing {self.descr} (session handle={self.instr.session})'))
            # logger.info(msg_head.format(f'Closing {self.descr} (session handle={self.instr.session})'))
            self.instr.close()
        else:
            print(msg_head.format(f'{self.descr} is "None" or already closed.'))
            # logger.info(msg_head.format(f'{self.descr} is "None" or already closed.'))
        self.is_open = 0

    def init(self):
        # Send initiation string
        msg_head = 'devices.instrument.init(): {}'
        s = self.InitStr
        if self.demo is True:
            # print(msg_head.format(f'{self.descr} in demo mode - no initiation necessary.'))
            # logger.info(msg_head.format(f'{self.descr} in demo mode - no initiation necessary.'))
            return 0
        else:
            if s != '':  # instrument has an initiation string
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    # print(msg_head.format(f'Failed to write {s} to {self.descr}'))
                    # logger.warning(msg_head.format(f'Failed to write {s} to {self.descr}'))
                    return -1
            else:
                print(msg_head.format(f'{self.descr} has no initiation string.'))
                # logger.warning(msg_head.format(f'{self.descr} has no initiation string.'))
                return 1
            print(msg_head.format(f'{self.descr} initiated with cmd:"{s}".'))
            # logger.info(msg_head.format(f'{self.descr} initiated with cmd:"{s}".'))
        return 1

    def set_v(self, v):
        """
        Set output voltage (SRC) or input range (DVM)
        """
        msg_head = 'devices.instrument.SetV() '
        if self.demo is True:
            return 0
        elif 'SRC_' in self.descr:
            # Set voltage-source to V
            s = str(v).join(self.VStr)
            # print(msg_head.format(f'{self.descr}: V = {v}; s = "{s}"'))
            # logging.info(msg_head.format(f'{self.descr}: V = {v}; s = "{s}"'))
            try:
                self.instr.write(s)
            except visa.VisaIOError:
                print(msg_head.format(f'Failed to write "{s}" to {self.descr},'
                                               f'via handle {self.instr.session}'))
                # logger.warning(msg_head.format(f'Failed to write "{s}" to {self.descr},'
                #                                f'via handle {self.instr.session}'))
                return -1
            return 1
        elif 'DVM_' in self.descr:
            # Set DVM range to V
            s = str(v).join(self.VStr)
            self.instr.write(s)
            # print(s)
            return 1
        else:  # 'none' in self.Descr, (or something odd has happened)
            print(msg_head.format(f'Invalid function for instrument {self.descr}'))
            # logger.warning(msg_head.format(f'Invalid function for instrument {self.descr}'))
            return -1

    def set_fn(self):
        # Set DVM function
        msg_head = 'devices.instrument.SetFn(): '
        if self.demo is True:
            return 0
        if 'DVM' in self.descr:
            s = self.SetFnStr
            if s != '':
                self.instr.write(s)
            # print(msg_head.format(f'{self.descr} - OK.'))
            # logger.info(msg_head.format(f'{self.descr} - OK.'))
            return 1
        else:
            print(msg_head.format(f'Invalid function for {self.descr}'))
            # logger.warning(msg_head.format(f'Invalid function for {self.descr}'))
            return -1

    def oper(self):
        # Enable O/P terminals
        # For V-source instruments only
        msg_head = 'devices.instrument.oper(): '
        if self.demo is True:
            return 0
        if 'SRC' in self.descr:
            s = self.OperStr
            if s != '':
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    print(msg_head.format(f'Failed to write "{s}" to {self.descr}'))
                    # logger.warning(msg_head.format(f'Failed to write "{s}" to {self.descr}'))
                    return -1
            # print(msg_head.format(f'{self.descr} output ENABLED.'))
            # logger.info(msg_head.format(f'{self.descr} output ENABLED.'))
            return 1
        else:
            print(msg_head.format(f'Invalid function for {self.descr}'))
            # logger.warning(msg_head.format(f'Invalid function for {self.descr}'))
            return -1

    def stby(self):
        # Disable O/P terminals
        # For V-source instruments only
        msg_head = 'devices.instrument.stby(): '
        if self.demo is True:
            return 0
        if 'SRC' in self.descr:
            s = self.StbyStr
            if s != '':
                self.instr.write(s)  # was: query(s)
            # print(msg_head.format(f'{self.descr} output DISABLED.'))
            # logger.debug(msg_head.format(f'{self.descr} output DISABLED.'))
            return 1
        else:
            print(msg_head.format(f'Invalid function for {self.descr}.'))
            # logger.warning(msg_head.format(f'Invalid function for {self.descr}.'))
            return -1

    def check_err(self):
        # Get last error string and clear error queue
        # For V-source instruments only (F5520A)
        msg_head = 'devices.instrument.check_err(): '
        reply = '-1'
        if self.demo is True:
            return '0'
        if 'F5520A' in self.descr:
            s = self.ChkErrStr
            if s != ('',):
                reply = self.instr.query(s[0])  # read error message
                self.instr.write(s[1])  # clear registers
            return reply
        else:
            print(msg_head.format(f'Invalid function for {self.descr}'))
            # logger.warning(msg_head.format(f'Invalid function for {self.descr}'))
            return '-1'

    def send_cmd(self, s):
        demo_reply = f'{self.descr} - DEMO resp. to {s}.'
        reply = ''
        if self.role == 'IVbox':  # update icb
            pass  # may need an event here...
        if self.demo is True:
            return demo_reply
        '''
        Check if s contains '?' or 'X' or is an empty string,
        in which case a response is expected:
        '''
        if any(x in s for x in '?X'):
            reply = self.instr.query(s)
            return reply
        elif s == '':
            reply = self.instr.read()
            return reply
        else:
            self.instr.write(s)
            # print(f'sending "{s}" to {self.descr}')
            # logger.info(f'sending "{s}" to {self.descr}')
            return reply

    def read(self):
        msg_head = 'devices.instrument.read(): '
        demo_reply = f'{self.descr} - DEMO resp.'
        reply = ''
        if self.demo is True:
            return demo_reply
        if 'DVM' in self.descr:
            # print(msg_head, f'from {self.descr}')
            # logger.debug(msg_head.format(f'from {self.descr}'))
            if '3458A' in self.descr:
                reply = self.instr.read()
                # print(reply)
                # logger.debug(msg_head.format(f'Reply = {reply}'))
                return reply
            else:
                reply = self.instr.query('READ?')
                return reply
        else:
            print(msg_head, f'Invalid function for {self.descr}.')
            # logger.warning(msg_head, f'Invalid function for {self.descr}.')
            return reply

    def test(self, s):
        """ Used to test that the instrument is functioning. """
        return self.send_cmd(s)
# __________________________________________
