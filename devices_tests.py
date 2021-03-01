# devices_tests.py

import pytest
import devices as dev


# Register marks:
# def pytest_configure(config):
#     config.addinivalue_line(
#         "markers", "sensor_info: Tests get_sensor_info() fn."
#     )


# strip_chars() tests:
# fixtures
@pytest.fixture()
def in_str():
    return '1234##&this part remains909uy+=-@'


@pytest.fixture()
def strip_str():
    return '1234#&90uy+=-@'


@pytest.fixture()
def remain_str():
    return 'this part remains'


@pytest.fixture()
def multi_in_str():
    return '###this #part #rem#ains####'


@pytest.fixture()
def multi_strip_str():
    return '#'


def test_strip_chars(in_str, strip_str, remain_str):
    assert dev.strip_chars(in_str, strip_str) == remain_str


def test_strip_chars_multi_chrs_to_strip(multi_in_str, multi_strip_str, remain_str):
    assert dev.strip_chars(multi_in_str, multi_strip_str) == remain_str


# ____Instrument class____ tests:
@pytest.fixture()
def demo_dvm():
    instr = dev.Instrument('DVM_34420A:s/n130', 'DVMd')
    dev.ROLES_WIDGETS.update({'DVMd': None})
    return instr


@pytest.fixture()
def visa_resources():
    return dev.RM.list_resources()


def test_visa_resources_is_not_empty(visa_resources):
    assert len(visa_resources) > 0


# Test sample INSTR_DATA entry:
def test_demo_instrument_init(demo_dvm):
    assert demo_dvm.str_addr == 'GPIB0::7::INSTR'
    assert demo_dvm.descr == 'DVM_34420A:s/n130'


def test_demo_open_is_none(demo_dvm):
    assert demo_dvm.open() is None


# Test that at least ONE resource (non-demo) can be found:
def test_visa_res_list(visa_resources):
    test_device = dev.RM.open_resource(visa_resources[0])
    assert test_device.session is not None


# Test that <demo instr>.init() returns 0
def test_demo_init_rtn0(demo_dvm):
    assert demo_dvm.init() == 0


# Test that <demo instr>.set_fn(<some value>) returns 0
def test_demo_set_v_rtn0(demo_dvm):
    assert demo_dvm.set_v(1.3) == 0


# Test that <demo instr>.set_fn() returns 0
def test_demo_set_fn_rtn0(demo_dvm):
    assert demo_dvm.set_fn() == 0


# Test that <demo instr>.oper() returns 0
def test_demo_oper_rtn0(demo_dvm):
    assert demo_dvm.oper() == 0


# Test that <demo instr>.stby() returns 0
def test_demo_stby_rtn0(demo_dvm):
    assert demo_dvm.stby() == 0


# Test that <demo instr>.check_err() returns '0'
def test_demo_check_err_rtn0str(demo_dvm):
    assert demo_dvm.check_err() == '0'


# Test that <demo instr>.send_cmd(<test_str>) returns demo reply
def test_demo_send_cmd_rtndemoresp(demo_dvm):
    d = demo_dvm.descr
    cmd = 'test_cmd'
    demo_resp = f'{d} - DEMO resp. to {cmd}.'
    assert demo_dvm.send_cmd(cmd) == demo_resp


# Test that <demo instr>.read() returns demo reply
def test_demo_read_rtndemoresp(demo_dvm):
    d = demo_dvm.descr
    demo_resp = f'{d} - DEMO resp.'
    assert demo_dvm.read() == demo_resp


