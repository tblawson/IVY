# -*- coding: utf-8 -*-
"""
devices.py

Gathers together all info required to use external instruments.

All device data is collected in a 'dictionary of dictionaries' - INSTR_DATA
Each piece of information is accessed as:
INSTR_DATA[<instrument description>][<parameter>].
E.g. the 'set function' string for the HP3458(s/n518) is:
INSTR_DATA['DVM_3458A:s/n518']['init_str'],
which retrieves the string 'FUNC OHMF;OCOMP ON'.
Note that some of the strings are LISTS of strings (i.e. multiple commands)

Created on Fri Mar 17 13:52:15 2017

@author: t.lawson
"""

import numpy as np
import os
import ctypes as ct
import visa
import logging


logger = logging.getLogger(__name__)
'''
INSTR_DATA:Dictionary of instrument parameter dictionaries,
keyed by description.
ROLES_WIDGETS: Dictionary of GUI widgets keyed by role.
ROLES_INSTR: Dictionary of GMH_sensor or Instrument objects,
keyed by role.
'''
INSTR_DATA = {
    'none': {'addr': 0, 'str_addr': 'NOADDRESS',
             'test': None,
             'demo': True,
             'role': None},
    'GMH:s/n628': {'addr': 5, 'str_addr': 'COM5',
                   'test': 'T',
                   'demo': True,
                   'role': None},
    'GMH:s/n367': {'addr': 8, 'str_addr': 'COM8',
                   'test': 'RH',
                   'demo': True,
                   'role': None},
    'GMH:s/n627': {'addr': 9, 'str_addr': 'COM9',
                   'test': 'T',
                   'demo': True,
                   'role': None},
    'DVM_34420A:s/n130': {'addr': 7, 'str_addr': 'GPIB0::7::INSTR',
                          'test': '*IDN?',
                          'init_str': 'FUNC OHMF;OCOMP ON',
                          'demo': True,
                          'role': None},
    'DVM_34401A:s/n976': {'addr': 17, 'str_addr': 'GPIB0::17::INSTR',
                          'test': '*IDN?',
                          'init_str': 'FUNC OHMF;OCOMP ON',
                          'demo': True,
                          'role': None},
    'DVM_3458A:s/n066': {'addr': 20, 'str_addr': 'GPIB0::20::INSTR',  # Was 0
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n129': {'addr': 25, 'str_addr': 'GPIB0::25::INSTR',
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n230': {'addr': 21, 'str_addr': 'GPIB0::21::INSTR',  # Was 22
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n382': {'addr': 22, 'str_addr': 'GPIB0::22::INSTR',
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n452': {'addr': 23, 'str_addr': 'GPIB0::23::INSTR',
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n518': {'addr': 24, 'str_addr': 'GPIB0::24::INSTR',
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'DVM_3458A:s/n452': {'addr': 23, 'str_addr': 'GPIB0::23::INSTR',
                         'test': 'ID?',
                         'init_str': 'DCV',
                         'demo': True,
                         'role': None},
    'SRC_D4808': {'addr': 2, 'str_addr': 'GPIB0::2::INSTR',
                  'test': 'X8=',
                  'init_str': ['F0G0D0S0=', 'M+0O1=', 'R0='],
                  'setV_str': ['M', '='],
                  'oper_str': 'O1=',
                  'stby_str': 'O0=',
                  'demo': True,
                  'role': None},
    'SRC_F5520A': {'addr': 4, 'str_addr': 'GPIB0::4::INSTR',
                   'test': '*IDN?',
                   'init_str': None,
                   'SetV_str': ['OUT', 'V,0Hz'],
                   'oper_str': 'OPER',
                   'stby_str': 'STBY',
                   'chk_err_str': ['ERR?', '*CLS'],
                   'demo': True,
                   'role': None},
    'IV_box': {'addr': 4, 'str_addr': 'COM4',
               'test': None,  # Depends on icb setting
               'demo': True,
               'role': 'IVbox'}
}

RED = (255, 0, 0)
GREEN = (0, 127, 0)

ROLES_WIDGETS = {}
ROLES_INSTR = {}

"""
VISA-specific stuff:
Only ONE VISA resource manager is required at any time -
All comunications for all GPIB and RS232 instruments (except GMH)
are handled by RM.
"""
RM = visa.ResourceManager()

# Switchbox
IVBOX_CONFIGS = {'V1': '1', 'V2': '2', '1k': '3', '10k': '4', '100k': '5',
                 '1M': '6', '10M': '7', '100M': '8', '1G': '9'}

T_Sensors = ('none', 'Pt', 'SR104t', 'thermistor')

"""
---------------------------------------------------------------
GMH-specific stuff:
GMH probe communications are handled by low-level routines in GMHdll.dll.
"""
os.environ['GMHPATH'] = 'I:\MSL\Private\Electricity\Staff\TBL\Python\High_Res_Bridge\GMHdll'
gmhpath = os.environ['GMHPATH']
GMHLIB = ct.windll.LoadLibrary(os.path.join(gmhpath, 'GMH3x32E'))
GMH_DESCR = ('GMH, s/n627',
             'GMH, s/n628')
LANG_OFFSET = 4096
'''--------------------------------------------------------------'''


class device():
    """
    A generic external device or instrument
    """
    def __init__(self, demo=True):
        self.demo = demo

    def open(self):
        pass

    def close(self):
        pass


class GMH_Sensor(device):
    """
    A class to wrap around the low-level functions of GMH3x32E.dll.
    For use with most Greisinger GMH devices.
    """
    def __init__(self, descr, role, demo=True):
        self.Descr = descr
        self.demo = demo

        '''
        self.addr: COM port-number assigned to USB 3100N adapter cable.
        '''
        self.addr = int(INSTR_DATA[self.Descr]['addr'])
        self.str_addr = INSTR_DATA[self.Descr]['str_addr']
        self.role = role

        '''
        self.flData: pointer to output data -
        Don't change this type!! It's the exactly right one!
        self.lang_offset: English language-offset
        self.MeasFn: access to GetMeasCode() fn
        self.UnitFn: access to GetUnitCode() fn
        self.ValFn: access to GetValue() fn
        '''
        self.Prio = ct.c_short()
        self.flData = ct.c_double()
        self.intData = ct.c_long()
        self.meas_str = ct.create_string_buffer(30)
        self.unit_str = ct.create_string_buffer(10)
        self.lang_offset = ct.c_int16(LANG_OFFSET)
        self.MeasFn = ct.c_short(180)
        self.UnitFn = ct.c_int16(178)
        self.ValFn = ct.c_short(0)
        self.SetPowOffFn = ct.c_short(223)
        self.error_msg = ct.create_string_buffer(70)
        self.meas_alias = {'T': 'Temperature',
                           'P': 'Absolute Pressure',
                           'RH': 'Rel. Air Humidity',
                           'T_dew': 'Dewpoint Temperature',
                           'T_wb': 'Wet Bulb Temperature',
                           'H_atm': 'Atmospheric Humidity',
                           'H_abs': 'Absolute Humidity'}
        self.info = {}

    def Open(self):
        """
        Use COM port number to open device
        Returns 1 if successful, 0 if not
        """
        self.error_code = ct.c_int16(GMHLIB.GMH_OpenCom(self.addr))
        self.GetErrMsg()  # Get self.error_msg

        if self.error_code.value in range(0, 4) or self.error_code.value == -2:
            print 'devices.GMH_Sensor.Open(): ', self.str_addr, 'is open.'
            logger.info('%s is open', self.str_addr)

            # We're not there yet - test device responsiveness
            self.GetErrMsg()
            if self.error_code.value in range(0, 4):  # Sensor responds...
                # Ensure max poweroff time
                self.intData.value = 120  # 120 mins B4 power-off
                self.Transmit(1, self.SetPowOffFn)

                self.Transmit(1, self.ValFn)
                if len(self.info) == 0:  # No device info yet
                    print 'devices.GMH_Sensor.Open(): Getting sensor info...'
                    logger.info('Getting sensor info...')
                    self.GetSensorInfo()
                    self.demo = False  # If we've got this far, probably OK
                    ROLES_WIDGETS[self.role]['lbl'].SetForegroundColour(GREEN)
                    ROLES_WIDGETS[self.role]['lbl'].Refresh()
                    return True
                else:  # Already have device measurement info
                    print'devices.GMH_Sensor.Open(): Instr ready. demo=False.'
                    logger.info('Instr ready. demo=False.')
                    self.demo = False  # If we've got this far, probably OK
                    return True
            else:  # No response
                print 'devices.GMH_Sensor.Open():', self.error_msg.value
                logger.info('%s', self.error_msg.value)
                self.Close()
                self.demo = True
                ROLES_WIDGETS[self.role]['lbl'].SetForegroundColour(RED)
                ROLES_WIDGETS[self.role]['lbl'].Refresh()
                return False

        else:  # Com open failed
            print'devices.GMH_Sensor.Open() FAILED:', self.Descr
            logger.warning('FAILED: %s', self.Descr)
            ROLES_WIDGETS[self.role]['lbl'].SetForegroundColour(RED)
            ROLES_WIDGETS[self.role]['lbl'].Refresh()
            self.Close()
            self.demo = True
            return False

    def Init(self):
        print'devices.GMH_Sensor.Init():', self.Descr,
        'initiated (nothing happens here).'
        logger.info('%s initiated (nothing happens here).', self.Descr)
        pass

    def Close(self):
        """
        Closes all / any GMH devices that are currently open.
        """
        self.demo = True
        self.error_code = ct.c_int16(GMHLIB.GMH_CloseCom())
        return 1

    def Transmit(self, Addr, Func):
        """
        A wrapper for the general-purpose interrogation function
        GMH_Transmit().
        """
        self.error_code = ct.c_int16(GMHLIB.GMH_Transmit(Addr, Func,
                                                         ct.byref(self.Prio),
                                                         ct.byref(self.flData),
                                                         ct.byref(self.intData)))
        self.GetErrMsg()
        if self.error_code.value < 0:
            print'\ndevices.GMH_Sensor.Transmit():FAIL'
            logger.warning('FAIL')
            return False
        else:
            print'\ndevices.GMH_Sensor.Transmit():PASS'
            logger.info('PASS')
            return True

    def GetErrMsg(self):
        """
        Translate return code into error message and store in self.error_msg.
        """
        error_code_ENG = ct.c_int16(self.error_code.value +
                                    self.lang_offset.value)
        GMHLIB.GMH_GetErrorMessageRet(error_code_ENG, ct.byref(self.error_msg))
        if self.error_code.value in range(0, 4):  # Correct message_0
            self.error_msg.value = 'Success'
        return 1

    def GetSensorInfo(self):
        """
        Interrogates GMH sensor.
        Returns a dictionary keyed by measurement string.
        Values are tuples: (<address>, <measurement unit>),
        where <address> is an int and <measurement unit> is a string.

        The address corresponds with a unique measurement function within
        the device. It's assumed the measurement functions are at
        consecutive addresses starting at 1.

        measurements list contains measurement-type strings e.g.
        'Temperature', 'Absolute Pressure', 'Rel. Air Humidity',etc
        """
        addresses = []  # Between 1 and 99
        measurements = []
        units = []  # E.g. 'deg C', 'hPascal', '%RH',...
        self.info.clear()

        for Address in range(1, 100):
            Addr = ct.c_short(Address)
            if self.Transmit(Addr, self.MeasFn):  # Result -> self.intData
                # Transmit() was successful
                addresses.append(Address)

                meas_code = ct.c_int16(self.intData.value +
                                       self.lang_offset.value)
                GMHLIB.GMH_GetMeasurement(meas_code,
                                          ct.byref(self.meas_str))
                measurements.append(self.meas_str.value)

                self.Transmit(Addr, self.UnitFn)  # Result -> self.intData

                unit_code = ct.c_int16(self.intData.value +
                                       self.lang_offset.value)
                GMHLIB.GMH_GetUnit(unit_code,
                                   ct.byref(self.unit_str))
                units.append(self.unit_str.value)

                print'Found', self.meas_str.value, '(', self.unit_str.value, ')', 'at address', Address
                logger.info('Found %s (%s) at address %d',
                            self.meas_str.value, self.unit_str.value, Address)
            else:
                print'devices.GMH_Sensor.GetSensorInfo(): Exhausted addresses at', Address
                logger.info('Exhausted addresses at %d', Address)
                if Address > 1:  # Don't let last address tried screw it up.
                    self.error_code.value = 0
                    self.demo = False
                else:
                    self.demo = True
                break  # Assumes all functions are in a contiguous address range from 1

        self.info = dict(zip(measurements, zip(addresses, units)))
        print 'devices.GMH_Sensor.GetSensorInfo():\n', self.info,
        'demo =', self.demo
        logger.info('%s demo = %s', self.info, self.demo)
        return len(self.info)

    def Measure(self, meas):
        """
        Measure either temperature, pressure or humidity, based on parameter
        meas
        Returns a float.
        meas is one of: 'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.

        NOTE that because GMH_CloseCom() acts on ALL open GMH devices it makes
        sense to only have a device open when communicating with it and to
        immediately close it afterwards. This way the default state is closed
        and the open state is treated as a special case. Hence an
        Open()-Close() 'bracket' surrounds the Measure() function.
        """

        self.flData.value = 0
        if self.Open():  # port and device open success
            assert self.demo is False, 'Illegal access to demo device!'
            Address = self.info[self.meas_alias[meas]][0]
            Addr = ct.c_short(Address)
            self.Transmit(Addr, self.ValFn)
            self.Close()

            print'devices.Measure():', self.meas_alias[meas],
            '=', self.flData.value
            logger.info('%s = %f', self.meas_alias[meas], self.flData.value)
            return self.flData.value
        else:
            assert self.demo is True, 'Illegal denial to demo device!'
            print'Generating demo data...'
            logger.info('Generating demo data...')
            demo_rtn = {'T': (-20.5, 0.2), 'P': (-1013, 5), 'RH': (-50, 10)}
            return np.random.normal(*demo_rtn[meas])

    def Test(self, meas):
        """ Used to test that the device is functioning. """
        print'\ndevices.GMH_Sensor.Test()...'
        logger.info('Testing %s with cmd %s...', self.Descr, meas)
        result = self.Measure(meas)
        return result


'''
###############################################################################
'''


class instrument(device):
    '''
    A class for associating instrument data with a VISA instance of
    that instrument
    '''
    def __init__(self, descr, role, demo=True):  # Default to demo mode
        self.Descr = descr
        self.demo = demo
        self.is_open = 0
        self.is_operational = 0

        assert_msg = 'Unknown instrument ({0:s})'.format(self.Descr)
        # check instrument data is loaded from Excel Parameters sheet.'
        assert self.Descr in INSTR_DATA, assert_msg

        self.addr = INSTR_DATA[self.Descr]['addr']
        self.str_addr = INSTR_DATA[self.Descr]['str_addr']
        self.role = role

        if 'init_str' in INSTR_DATA[self.Descr]:
            self.InitStr = INSTR_DATA[self.Descr]['init_str']  # tuple of str
        else:
            self.InitStr = ('',)  # a tuple of empty strings
        if 'setfn_str' in INSTR_DATA[self.Descr]:
            self.SetFnStr = INSTR_DATA[self.Descr]['setfn_str']
        else:
            self.SetFnStr = ''  # an empty string
        if 'oper_str' in INSTR_DATA[self.Descr]:
            self.OperStr = INSTR_DATA[self.Descr]['oper_str']
        else:
            self.OperStr = ''  # an empty string
        if 'stby_str' in INSTR_DATA[self.Descr]:
            self.StbyStr = INSTR_DATA[self.Descr]['stby_str']
        else:
            self.StbyStr = ''
        if 'chk_err_str' in INSTR_DATA[self.Descr]:
            self.ChkErrStr = INSTR_DATA[self.Descr]['chk_err_str']
        else:
            self.ChkErrStr = ('',)
        if 'setV_str' in INSTR_DATA[self.Descr]:
            self.VStr = INSTR_DATA[self.Descr]['setV_str']  # a tuple of str
        else:
            self.VStr = ''

    def Open(self):
        try:
            self.instr = RM.open_resource(self.str_addr)
            if '3458A' in self.Descr:
                self.instr.read_termination = '\r\n'
                self.instr.write_termination = '\r\n'
            self.instr.timeout = 2000  # default 2 s timeout
            INSTR_DATA[self.Descr]['demo'] = False  # A real working instrument
            self.demo = False  # A real instrument ONLY on Open() success
            green = (0, 255, 0)
            ROLES_WIDGETS[self.role]['lbl'].SetBackgroundColour(green)
            print 'devices.instrument.Open():', self.Descr,
            'session handle=', self.instr.session
            logger.info('%s: session handle=%d', self.Descr,
                        self.instr.session)
            self.is_open = 1
        except visa.VisaIOError:
            self.instr = None
            self.demo = True  # default to demo mode if can't open
            red = (255, 0, 0)
            ROLES_WIDGETS[self.role]['lbl'].SetForegroundColour(red)
            ROLES_WIDGETS[self.role]['lbl'].Refresh()
            INSTR_DATA[self.Descr]['demo'] = True
            print 'devices.instrument.Open() failed:', self.Descr, 'opened in demo mode' 
            logger.warning('Failed: %s opened in demo mode', self.Descr)
        return self.instr

    def Close(self):
        # Close comms with instrument
        if self.demo is True:
            print 'devices.instrument.Close():', self.Descr,
            'in demo mode - nothing to close'
            logger.info('%s in demo mode - nothing to close.', self.Descr)
        elif self.instr is not None:
            print 'devices.instrument.Close(): Closing', self.Descr,
            '(session handle=', self.instr.session, ')'
            logger.info('Closing %s (session handle=%d)',
                        self.Descr, self.instr.session)
            self.instr.close()
        else:
            print 'devices.instrument.Close():', self.Descr,
            'is "None" or already closed'
            logger.info('%s is "None" or already closed', self.Descr)
        self.is_open = 0

    def Init(self):
        # Send initiation string
        if self.demo is True:
            print 'devices.instrument.Init():', self.Descr,
            'in demo mode - no initiation necessary'
            logger.info('%s in demo mode - no initiation necessary',
                        self.Descr)
            return 1
        else:
            reply = 1
            for s in self.InitStr:
                if s != '':  # instrument has an initiation string
                    try:
                        self.instr.write(s)
                    except visa.VisaIOError:
                        print'Failed to write "%s" to %s' % (s, self.Descr)
                        logger.warning('Failed to write "%s" to %s',
                                       s, self.Descr)
                        reply = -1
                        return reply
            print 'devices.instrument.Init():', self.Descr,
            'initiated with cmd:', s
            logger.info('%s initiated with cmd: %s', self.Descr, s)
        return reply

    def SetV(self, V):
        '''
        Set output voltage (SRC) or input range (DVM)
        '''
        if self.demo is True:
            return 1
        elif 'SRC:' in self.Descr:
            # Set voltage-source to V
            s = str(V).join(self.VStr)
            print'devices.instrument.SetV(): V =', V
            print'devices.instrument.SetV():', self.Descr, 's=', s
            logging.info('%s: V = %f; s = "%s"', self.Descr, V, s)
            try:
                self.instr.write(s)
            except visa.VisaIOError:
                print'Failed to write "%s" to %s,\
                via handle %s' % (s, self.Descr, self.instr.session)
                logger.warning('Failed to write "%s" to %s via handle %s',
                               s, self.Descr, self.instr.session)
                return -1
            return 1
        elif 'DVM:' in self.Descr:
            # Set DVM range to V
            s = str(V).join(self.VStr)
            self.instr.write(s)
            return 1
        else:  # 'none' in self.Descr, (or something odd has happened)
            print 'Invalid function for instrument', self.Descr
            logger.warning('Invalid function for instrument %s', self.Descr)
            return -1

    def SetFn(self):
        # Set DVM function
        if self.demo is True:
            return 1
        if 'DVM' in self.Descr:
            s = self.SetFnStr
            if s != '':
                self.instr.write(s)
            print'devices.instrument.SetFn():', self.Descr, '- OK.'
            logger.info('%s OK', self.Descr)
            return 1
        else:
            print'devices.instrument.SetFn(): Invalid function for', self.Descr
            logger.warning('Invalid function for %s', self.Descr)
            return -1

    def Oper(self):
        # Enable O/P terminals
        # For V-source instruments only
        if self.demo is True:
            return 1
        if 'SRC' in self.Descr:
            s = self.OperStr
            if s != '':
                try:
                    self.instr.write(s)
                except visa.VisaIOError:
                    print'Failed to write "%s" to %s' % (s, self.Descr)
                    logger.warning('Failed to write "%s" to %s', s, self.Descr)
                    return -1
            print'devices.instrument.Oper():', self.Descr, 'output ENABLED.'
            logger.info('%s output ENABLED.', self.Descr)
            return 1
        else:
            print'devices.instrument.Oper(): Invalid function for', self.Descr
            logger.warning('Invalid function for %s', self.Descr)
            return -1

    def Stby(self):
        # Disable O/P terminals
        # For V-source instruments only
        if self.demo is True:
            return 1
        if 'SRC' in self.Descr:
            s = self.StbyStr
            if s != '':
                self.instr.write(s)  # was: query(s)
            print'devices.instrument.Stby():', self.Descr, 'output DISABLED.'
            logger.info('%s output DISABLED.', self.Descr)
            return 1
        else:
            print'devices.instrument.Stby(): Invalid function for', self.Descr
            logger.warning('Invalid function for %s', self.Descr)
            return -1

    def CheckErr(self):
        # Get last error string and clear error queue
        # For V-source instruments only (F5520A)
        if self.demo is True:
            return 1
        if 'F5520A' in self.Descr:
            s = self.ChkErrStr
            if s != ('',):
                reply = self.instr.query(s[0])  # read error message
                self.instr.write(s[1])  # clear registers
            return reply
        else:
            print'devices.instrument.CheckErr(): Invalid function for',
            self.Descr
            logger.warning('Invalid function for %s', self.Descr)
            return -1

    def SendCmd(self, s):
        demo_reply = self.Descr + ' - DEMO resp. to ' + s
        reply = 1
        if self.role == 'IVbox':  # update icb
            pass  # may need an event here...
        if self.demo is True:
            return demo_reply
        '''
        Check if s contains '?' or 'X' or is an empty string,
        in which case a response is expected:
        '''
        if any(x in s for x in'?X'):
            reply = self.instr.query(s)
            return reply
        elif s == '':
            reply = self.instr.read()
            return reply
        else:
            self.instr.write(s)
            return reply

    def Read(self):
        reply = 0
        if self.demo is True:
            return reply
        if 'DVM' in self.Descr:
            print'devices.instrument.Read(): from', self.Descr
            logger.info('Reading from %s...', self.Descr)
            if '3458A' in self.Descr:
                reply = self.instr.read()
                print reply
                logger.info('Reply = %s', reply)
                return reply
            else:
                reply = self.instr.query('READ?')
                return reply
        else:
            print 'devices.instrument.Read(): Invalid function for', self.Descr
            logger.warning('Invalid function for %s', self.Descr)
            return reply

    def Test(self, s):
        """ Used to test that the instrument is functioning. """
        return self.SendCmd(s)
# __________________________________________
