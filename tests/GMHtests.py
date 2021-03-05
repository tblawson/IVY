# GMHtests.py

from IVY import GMHstuff as Gmh
import pytest


# Sanity-checking:
def test_always_passes():
    assert True


# def test_always_fails():
#     assert False


# Now the REAL testing starts...

# Register marks:
def pytest_configure(config):
    config.addinivalue_line(
        "markers", "sensor_info: Tests get_sensor_info() fn."
    )
    config.addinivalue_line(
        "markers", "rtncode: Tests rtncode_to_errmsg() fn."
    )
    config.addinivalue_line(
        "markers", "get_disp: Tests get_disp_unit() fn."
    )
    config.addinivalue_line(
        "markers", "get_range: Tests get_min_range() and get_max_range() fns."
    )
    config.addinivalue_line(
        "markers", "sw_info: Tests get_sw_info() fn."
    )
    config.addinivalue_line(
        "markers", "meas_attribs: Tests get_meas_attributes() fn."
    )
    config.addinivalue_line(
        "markers", "measure: Tests measure() fn."
    )
    config.addinivalue_line(
        "markers", "open_close: Tests open() and close() fns."
    )
    config.addinivalue_line(
        "markers", "transmit: Tests transmit() fn."
    )
    config.addinivalue_line(
        "markers", "get_xxx: Tests various get_...() fns."
    )

# Set up fixtures first...
@pytest.fixture
def sensor():  # Test behaviour of GMH sensor in demo mode (COM0):
    return Gmh.GMHSensor(0, demo=True)


@pytest.fixture
def probe_info(sensor):
    return sensor.get_sensor_info()


@pytest.fixture
def first_key(probe_info):
    return next(probe_info.keys().__iter__())


# get_sensor_info tests:
@pytest.mark.sensor_info
def test_get_sensor_info_rtn_dict(probe_info):
    assert isinstance(probe_info, dict)


@pytest.mark.sensor_info
def test_get_sensor_info_rtnkey0_str(first_key):
    assert isinstance(first_key, str)


@pytest.mark.sensor_info
def test_get_sensor_info_rtnval0_tuple(first_key, probe_info):
    assert isinstance(probe_info[first_key], tuple)


@pytest.mark.sensor_info
def test_get_sensor_info_rtnval00_int(first_key, probe_info):
    assert isinstance(probe_info[first_key][0], int)


@pytest.mark.sensor_info
def test_get_sensor_info_rtnval00_int(first_key, probe_info):
    assert isinstance(probe_info[first_key][1], str)


# rtncode_to_errmsg() tests:
@pytest.mark.rtncode
def test_demo_rtncode_to_errmsg_no_stat_zero_in_rtn_str(sensor):
    assert isinstance(sensor.rtncode_to_errmsg(0), str)


@pytest.mark.rtncode
def test_demo_rtncode_to_errmsg_stat_zero_in_rtn_tuple(sensor):
    assert isinstance(sensor.rtncode_to_errmsg(0, True), tuple)


@pytest.mark.rtncode
def test_demo_rtncode_to_errmsg_stat_zero_in_rtn0_str(sensor):
    assert isinstance(sensor.rtncode_to_errmsg(0, True)[0], str)


@pytest.mark.rtncode
def test_demo_rtncode_to_errmsg_stat_zero_in_rtn1_str(sensor):
    assert isinstance(sensor.rtncode_to_errmsg(0, True)[1], str)


# open_port() tests:
@pytest.mark.open_close
def test_demo_open_port_rtn_int(sensor):
    assert isinstance(sensor.open_port(), int)


# close() tests:
@pytest.mark.open_close
def test_demo_close_rtn_int(sensor):
    assert isinstance(sensor.close(), int)


# transmit() tests:
@pytest.mark.transmit
def test_demo_transmit_rtn_int(sensor):
    assert isinstance(sensor.transmit(1, 'GetStatus'), int)


# get_type() tests:
@pytest.mark.get_xxx
def test_demo_get_type_rtn_str(sensor):
    assert isinstance(sensor.get_type(), str)


# get_num_chans() tests:
@pytest.mark.get_xxx
def test_demo_get_num_chans_rtn_int(sensor):
    assert isinstance(sensor.get_num_chans(), int)


# get_status() tests:
@pytest.mark.get_xxx
def test_demo_get_status_rtn_str(sensor):
    assert isinstance(sensor.get_status(1), str)


# get_unit() tests:
@pytest.mark.get_xxx
def test_demo_get_unit_rtn_str(sensor):
    assert isinstance(sensor.get_unit(1), str)


# get_disp_unit() tests:
@pytest.mark.get_disp
def test_demo_get_disp_unit_rtn_str(sensor):
    assert isinstance(sensor.get_disp_unit(1), str)


# get_disp_min_range tests:
@pytest.mark.get_disp
def test_demo_get_disp_min_range_rtn_number(sensor):
    assert isinstance(sensor.get_disp_min_range(1), (int, float))


# get_disp_max_range tests:
@pytest.mark.get_disp
def test_demo_get_disp_max_range_rtn_number(sensor):
    assert isinstance(sensor.get_disp_max_range(1), (int, float))


# get_min_range tests:
@pytest.mark.get_range
def test_demo_get_min_range_rtn_number(sensor):
    assert isinstance(sensor.get_min_range(1), (int, float))


# get_max_range tests:
@pytest.mark.get_range
def test_demo_get_max_range_rtn_number(sensor):
    assert isinstance(sensor.get_max_range(1), (int, float))


# get_power_off_time tests:
def test_get_power_off_time_rtn_number(sensor):
    assert isinstance(sensor.get_power_off_time(), (int, float))


# get_sw_info tests:
@pytest.mark.sw_info
def test_demo_get_sw_info_rtn_tuple(sensor):
    assert isinstance(sensor.get_sw_info(), tuple)


@pytest.mark.sw_info
def test_demo_get_sw_info_rtn0_float(sensor):
    assert isinstance(sensor.get_sw_info()[0], float)


@pytest.mark.sw_info
def test_demo_get_sw_info_rtn1_int(sensor):
    assert isinstance(sensor.get_sw_info()[1], int)


# set_power_off_time (mins) tests:
@pytest.mark.power_off
# @pytest.mark.parametrize('off_time,rtn', [
#     (-1, -1),
#     (0, -1),
#     (1, 1),
#     (60, 60)
# ])
# def test_set_power_off_time_output_is_input(sensor, off_time, rtn):
#     assert sensor.set_power_off_time(off_time) == rtn
def test_demo_set_power_off_time_output_is_err(sensor):
    assert sensor.set_power_off_time(60) == -1


# get_meas_attributes tests:
@pytest.mark.meas_attribs
def test_demo_get_meas_attributes_rtn_tuple(sensor):
    assert isinstance(sensor.get_meas_attributes('T'), tuple)


@pytest.mark.meas_attribs
def test_demo_get_meas_attributes_rtn_val0_int(sensor):
    assert isinstance(sensor.get_meas_attributes('T')[0], int)


@pytest.mark.meas_attribs
def test_demo_get_meas_attributes_rtn_val1_str(sensor):
    assert isinstance(sensor.get_meas_attributes('T')[1], str)


# measure tests:
@pytest.mark.measure
def test_demo_measure_rtn_tuple(sensor):
    assert isinstance(sensor.measure('T'), tuple)


@pytest.mark.measure
def test_demo_measure_rtn_val0_float(sensor):
    assert isinstance(sensor.measure('T')[0], float)


@pytest.mark.measure
def test_demo_measure_rtn_val1_str(sensor):
    assert isinstance(sensor.measure('T')[1], str)

