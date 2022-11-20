# -*- coding: utf-8 -*-
"""
GMHstuff.py - required to access dll functions for GMH probes
Created on Wed Jul 29 13:21:22 2015

@author: t.lawson
"""

import os
import ctypes as ct


# os.path.join('C:', 'Users', 't.lawson', 'PycharmProjects', 'GMHstuff2')
# Change PATH to wherever you keep GMH3x32E.dll:
gmhlibpath = '..\\scripts\\GMHstuff\\GMHdll'  # G:\\Shared drives\\MSL - Electricity\\Ongoing\\OHM\\Temperature_PRTs\\GMHdll'
os.environ['GMH_PATH'] = gmhlibpath
GMHpath = os.getenv('GMH_PATH')
print(f'GMHpath = {GMHpath}')
GMHLIB = ct.windll.LoadLibrary(os.path.join(GMHpath, 'GMH3x32E'))  # (os.path.join(GMHpath, 'GMH3x32E'))

# A (useful) subset of Transmit() function calls:
TRANSMIT_CALLS = {'GetValue': 0, 'GetStatus': 3, 'GetTypeCode': 12, 'GetMinRange': 176, 'GetMaxRange': 177,
                  'GetUnitCode': 178, 'GetMeasCode': 180, 'GetDispMinRange': 200, 'GetDispMaxRange': 201,
                  'GetDispUnitCode': 202, 'GetDispDecPoint': 204, 'GetChannelCount': 208, 'GetPowerOffTime': 222,
                  'SetPowerOffTime': 223, 'GetSoftwareInfo': 254}

MEAS_ALIAS = {'T': 'Temperature', 'P': 'Absolute Pressure', 'RH': 'Rel. Air Humidity',
              'T_dew': 'Dewpoint Temperature', 'T_wb': 'Wet Bulb Temperature',
              'H_atm': 'Atmospheric Humidity', 'H_abs': 'Absolute Humidity'}

C_LANG_OFFSET = ct.c_int16(4096)  # English language-offset


class GMHSensor:
    """
    A class to wrap around the low-level functions of GMH3x32E.dll.

    For use with Greisinger GMH devices (GFTB200, etc.).
    """
    def __init__(self, port, demo=True):
        """
        Initiate a GMH sensor object.

        :param port: COM port number to which GMH device is attached (via 3100N cable)

        :param demo: Describes if this object is NOT capable of operating as a real GMH sensor (default is True)

        False - the COM port connecting to a real device is open AND that device is turned on and present,

        True - the device is a combination of one or more of:

        * not turned on OR

        * not connected OR

        * not associated with an open COM port.
        """
        self.demo = demo
        self.port = port
        self.com_open = False
        self.error_msg = '-'
        self.status_msg = '-'
        self.error_code = 0
        self.type_str = '-'
        self.chan_count = 0

        # All ctypes objects have a c_ suffix:
        self.c_Prio = ct.c_short()
        self.c_flData = ct.c_double()
        self.c_intData = ct.c_long()
        self.c_meas_str = ct.create_string_buffer(30)
        self.c_error_msg = ct.create_string_buffer(70)

        self._info = {}  # Don't access directly - use get functions

    @staticmethod
    def rtncode_to_errmsg(rtn_code, stat=False):
        """
        Translate a function's return code to a message-string.

        Shift return code by language offset,
        Obtain message string (as byte-stream) corresponding to translated code,
        Decode byte-stream to message string,
        print message.

        :arguments:
            int
        :returns
            string (Unicode)
        """
        c_msg = ct.create_string_buffer(70)
        c_status = ct.create_string_buffer(70)
        c_translated_code = ct.c_int16(rtn_code + C_LANG_OFFSET.value)
        GMHLIB.GMH_GetErrorMessageRet(c_translated_code, ct.byref(c_msg))
        # GMHLIB.GMH_GetStatusMessage(c_translated_code, ct.byref(c_status))
        msg = c_msg.value.decode('ISO-8859-1')
        status = c_status.value.decode('ISO-8859-1')
        if 'EASYBus' in msg:
            rtn = 'OK.'
        else:
            rtn = msg
        if stat:
            return rtn, status
        else:
            return rtn

    def open_port(self):
        """
        Open a single COM port for a 3100N GMH adapter cable.

        Only one COM port can be open at a time. Up to 5 GMH devices can be serviced through
        one COM port (but requires special hardware fan-out).

        :returns:
            1 for success,
            -1 for failure.
        """
        if self.com_open is True:
            return 1
        try:
            c_rtn_code = ct.c_int16(GMHLIB.GMH_OpenCom(self.port))
            self.error_code = c_rtn_code.value
            self.error_msg = self.rtncode_to_errmsg(c_rtn_code.value)
            assert self.error_code >= 0, 'GMHLIB.GMH_OpenCom() failed'
        except (AssertionError, ct.ArgumentError, TypeError) as msg:
            print('open_port()_except:', msg, '{} "{}"'.format(self.error_code, self.error_msg))
            pass
            print('port={}, type(port)={}'.format(self.port, type(self.port)))
            self.com_open = False
            return -1
        else:
            self.error_code = c_rtn_code.value
            self.com_open = True
            # print('open_port(): calling get_sensor_info()...')
            self.get_sensor_info()
            return 1

    def close(self):
        """
        Close open COM port (only one).

        Note that COM port could still be open even if object is in demo mode.

        :returns
            1 if successful, 0 if no action taken (port already closed), -1 otherwise.
        """
        if self.com_open is False:
            return 0
        else:
            try:
                rtn = GMHLIB.GMH_CloseCom()
                assert rtn >= 0, 'GMH_CloseCom() error.'
            except AssertionError as msg:
                print(msg)
                self.com_open = True
                return -1
            else:
                self.com_open = False
                return 1

    def transmit(self, chan, func_name):
        """
        A wrapper for the GMH general-purpose interrogation function GMH_Transmit().

        Runs func_name on instrument channel chan and updates self.error_code.
        self.error_code is then used to update self.error_msg and self.status_msg.

        :argument
            chan: Measurement channel (int) of GMH device (0-99)
            func_name: Function-call name (unicode string) - see transmit_calls dict

        :returns
            1 for success (non-demo mode),
            -1 for GMHLIB.GMH_Transmit() failure
        """
        c_chan = ct.c_int16(chan)
        c_func = ct.c_int16(TRANSMIT_CALLS[func_name])
        try:
            self.error_code = GMHLIB.GMH_Transmit(c_chan, c_func, ct.byref(self.c_Prio),
                                                  ct.byref(self.c_flData),
                                                  ct.byref(self.c_intData))
            self.error_msg = self.rtncode_to_errmsg(self.error_code)
            assert self.error_code >= 0, 'GMH_Transmit().'
        except AssertionError as msg:
            print('transmit(ch{})_except:'.format(c_chan.value), msg, '{} "{}"'.format(self.error_code, self.error_msg))
            return -1
        else:
            # print('transmit({})_else: GMH_Transmit() return: {} "{}"'.format(c_func, self.error_code,
            #                                                                  self.error_msg))
            return 1

    def get_type(self):
        """
        Get instrument type

        :return: Instrument type (as unicode string)
        """
        # Get instrument type code -> self.c_intData:
        self.transmit(1, 'GetTypeCode')
        c_translated_type_code = ct.c_int16(self.c_intData.value + C_LANG_OFFSET.value)
        c_type_str = ct.create_string_buffer(30)
        # Interpret type code to type string:
        try:
            c_rtn_len = ct.c_byte(GMHLIB.GMH_GetType(c_translated_type_code, ct.byref(c_type_str)))
            self.error_code = c_rtn_len.value
            type_str = c_type_str.value.decode('ISO-8859-1')
            assert self.error_code >= 1, 'GMHLIB.GMH_GetType() failed'
        except AssertionError as msg:
            type_str = 'UNKNOWN instrument type.'
            print('get_type():', msg, '{} "{}"'.format(self.error_code, type_str))
        else:
            return type_str

    def get_num_chans(self):
        """
        Update self.chan_count.

        :return: self.chan_count.
        """
        # Get number of measurement channels for this instrument -> self.c_intData:
        self.transmit(1, 'GetChannelCount')
        if self.error_code < 0:
            self.chan_count = 0
            print('No channels found!: {}'.format(self.rtncode_to_errmsg(self.error_code)))
        else:
            self.chan_count = self.c_intData.value
            print('{} channels found: {}'.format(self.chan_count, self.rtncode_to_errmsg(self.error_code)))
        return self.chan_count

    def get_status(self, chan):
        """
        Get instrument's status string.

        :param chan: instrument measurement channel (int)
        :return: status message (unicode)
        """
        # Get instrument status code -> self.c_intData:
        self.error_code = self.transmit(chan, 'GetStatus')
        c_status_msg = ct.create_string_buffer(70)
        if self.error_code < 0:
            status_msg = 'status: Not available!'
        else:
            c_translated_status_code = ct.c_int16(self.c_intData.value + C_LANG_OFFSET.value)
            # Interpret status code -> status string:
            GMHLIB.GMH_GetStatusMessage(c_translated_status_code, ct.byref(c_status_msg))
            status_msg = c_status_msg.value.decode('ISO-8859-1')
        return status_msg

    def get_unit(self, chan):
        """
        Get measurement unit for this channel

        :param chan: Instrument measurement channel (int)
        :return: Measurement unit (unicode string)
        """
        self.error_code = self.transmit(chan, 'GetUnitCode')
        c_unit_str = ct.create_string_buffer(10)
        if self.error_code < 0:
            unit = 'No unit'
            print('No measurement unit found!')
        else:
            c_translated_unit_code = ct.c_int16(self.c_intData.value + C_LANG_OFFSET.value)
            # Write result to self.c_unit_str:
            GMHLIB.GMH_GetUnit(c_translated_unit_code, ct.byref(c_unit_str))
            unit = c_unit_str.value.decode('ISO-8859-1')
        return unit

    def get_disp_min_range(self, chan):
        """
        Get Min range of channel display.

        :param chan: Instrument measurement channel (int)
        :return: Min range of display (int)
        """
        self.error_code = self.transmit(chan, 'GetDispMinRange')
        if self.error_code < 0:
            disp_min_range = 0
            print('No display min range found!')
        else:
            disp_min_range = self.c_intData.value
        return disp_min_range

    def get_disp_max_range(self, chan):
        """
        Get max range of channel display.

        :param chan: Instrument measurement channel (int)
        :return: Max range of display (int)
                """
        self.error_code = self.transmit(chan, 'GetDispMaxRange')
        if self.error_code < 0:
            disp_max_range = 0
            print('No display max range found!')
        else:
            disp_max_range = self.c_intData.value
        return disp_max_range

    def get_min_range(self, chan):
        """
        Get min range of this channel.

        :param chan: Instrument measurement channel (int)
        :return: Min range (int)
        """
        self.error_code = self.transmit(chan, 'GetMinRange')
        if self.error_code < 0:
            min_range = 0
            print('No min range found!')
        else:
            min_range = self.c_intData.value
        return min_range

    def get_max_range(self, chan):
        """
        Get max range of this channel.

        :param chan: Instrument measurement channel (int)
        :return: Max range (int)
        """
        self.error_code = self.transmit(chan, 'GetMaxRange')
        if self.error_code < 0:
            max_range = 0
            print('No max range found!')
        else:
            max_range = self.c_intData.value
        return max_range

    def get_disp_unit(self, chan):
        """
        Get measurement unit for this channel

        :param chan: Instrument measurement channel (int)
        :return: Display unit (unicode string)
        """
        self.error_code = self.transmit(chan, 'GetDispUnitCode')
        c_unit_str = ct.create_string_buffer(10)
        if self.error_code < 0:
            unit = 'No unit'
            print('No display unit found!')
        else:
            c_translated_unit_code = ct.c_int16(self.c_intData.value + C_LANG_OFFSET.value)
            # Write result to self.c_unit_str:
            GMHLIB.GMH_GetUnit(c_translated_unit_code, ct.byref(c_unit_str))
            unit = c_unit_str.value.decode('ISO-8859-1')
        return unit

    def get_power_off_time(self):
        """
        Get power-off time.

        :return: power-off time in minutes (int) - negative if error.
        """
        self.error_code = self.transmit(1, 'GetPowerOffTime')
        if self.error_code < 0:
            pow_off_t = -1
            print('No power-off time found!')
        else:
            pow_off_t = self.c_intData.value
        return pow_off_t

    def set_power_off_time(self, mins):
        """
        Set power-off time.

        :return: requested power-off time in minutes (int) or -1 if error.
        """
        self.c_intData = ct.c_int16(mins)
        self.error_code = self.transmit(1, 'SetPowerOffTime')
        if self.error_code < 0:
            pow_off_t = -1
            print('Power-off time not changed.')
        else:
            # Fn returns requested power-off time (mins)
            pow_off_t = self.c_intData.value
        return pow_off_t

    def get_sw_info(self):
        """
        Get software version and identifier.

        :return: (version, identifier) tuple (both ints).
        """
        self.error_code = self.transmit(1, 'GetSoftwareInfo')
        if self.error_code < 0:
            version = 0.0
            ident = 0
            print('Software info not found!')
        else:
            version = self.c_flData.value
            ident = self.c_intData.value
        return version, ident

    def get_sensor_info(self):
        """
        Interrogates GMH sensor for measurement capabilities.

        Checks if self.info is already populated. If so, returns self.info.
        Otherwise, proceeds to gather required info...

        :returns
        self.info - a dictionary keyed by measurement string (eg: 'Temperature').
        Values are tuples: (<address>, <measurement unit>),
        where <address> is an int and <measurement unit> is a unicode string.
        """
        channels = []  # Between 1 and 99
        measurements = []  # E.g. 'Temperature', 'Absolute Pressure', ...
        units = []  # E.g. 'deg C', 'hPascal', ...

        if len(self._info) > 0:  # Device info already determined.
            # print('\nget_sensor_info(): device info already determined.')
            return self._info
        else:
            # Find all channel-independent parameters
            self.get_num_chans()
            if self.chan_count == 0:
                return {'NO SENSOR': (0, 'NO UNIT')}
            else:
                # Visit all the channels and note their capabilities:
                channel = 0
                while channel <= self.chan_count:
                    # print('get_sensor_info(): Testing channel {}...'.format(channel))
                    # Try reading a value, Write result to self.c_intData:
                    self.error_code = self.transmit(channel, 'GetValue')
                    if self.error_code < 0:
                        # print('get_sensor_info(): No measurement function at channel {}'.format(channel))
                        channel += 1
                        continue  # Skip to next channel if this one has no value to read
                    else:  # Successfully got a dummy value
                        self.error_code = self.transmit(channel, 'GetMeasCode')
                        if self.c_intData.value < 0:
                            print('get_sensor_info(): transmit() failure to get meas code.')
                            channel += 1
                            continue  # Bail-out if not a valid measurement code

                        # Now we have a valid measurement code...
                        c_translated_meas_code = ct.c_int16(self.c_intData.value + C_LANG_OFFSET.value)
                        # Write result to self.c_meas_str:
                        GMHLIB.GMH_GetMeasurement(c_translated_meas_code, ct.byref(self.c_meas_str))
                        measurements.append(self.c_meas_str.value.decode('ISO-8859-1'))

                        units.append(self.get_unit(channel))
                        channels.append(channel)
                        channel += 1

                        self.demo = False  # If we've got this far we must have a fully-functioning instrument.
                        # print('get_sensor_info(): demo mode = {}'.format(self.demo))
                self._info = dict(zip(measurements, zip(channels, units)))
                return self._info

    def get_meas_attributes(self, meas):
        """
        Look up attributes associated with a type of measurement.

        :param meas: Measurement-type alias (any key in MEAS_ALIAS dict) - a unicode string,
        :return: Tuple consisting of the device channel (int) and measurement unit (unicode).
        """
        fail_rtn = (0, 'NO_UNIT')
        try:
            measurement = MEAS_ALIAS[meas]
        except KeyError as msg:
            print('\n{} - No known alias for {}.'.format(msg, meas))
            return fail_rtn
        else:
            try:
                chan_unit_str = self._info[measurement]
            except KeyError as msg:
                print("\n{} - measurement {} doesn't exist for this device.".format(msg, meas))
                return fail_rtn
            else:
                return chan_unit_str

    def measure(self, meas_type):
        """
        Make a measurement (temperature, pressure, humidity, etc).

        :argument meas - an alias (as defined in MEAS_ALIAS) for the measurement type.
        meas is one of: 'T', 'P', 'RH', 'T_dew', 't_wb', 'H_atm' or 'H_abs'.

        :returns a tuple: (<measurement as float>, <unit as unicode string>)
        """
        self.open_port()
        default_reading = (0.0, 'NO UNIT')
        if len(self._info) == 0:
            reading = default_reading
            print('Measure(): No measurement info available!')
        elif meas_type not in MEAS_ALIAS.keys():
            reading = default_reading
            print('Unknown function "{}"'.format(meas_type))
        elif MEAS_ALIAS[meas_type] not in self._info.keys():
            reading = default_reading
            print('Function', meas_type, 'not available on this instrument!')
        else:
            channel = self._info[MEAS_ALIAS[meas_type]][0]
            self.error_code = self.transmit(channel, 'GetValue')
            if self.error_code < 0:
                reading = default_reading
                print('Measurement value not found!- Check sensor is connected and ON.')
            else:
                # chan = self._info[MEAS_ALIAS[meas]][0]
                unit_str = self._info[MEAS_ALIAS[meas_type]][1]
                reading = (self.c_flData.value, unit_str)
                # print('Measured {} from port {}, chan: {}.'.format(reading, self.port, channel))
                # print(self._info)
        self.close()
        return reading
