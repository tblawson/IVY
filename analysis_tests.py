# analysis_tests.py

import pytest
import nbpages as pages
import GTC


# Fixtures
@pytest.fixture()
def neg_val():
    return -1234.567


@pytest.fixture()
def pos_val():
    return 1234.567


@pytest.fixture()
def unc():
    return 0.012


def test_set_precision_neg_val(neg_val, unc):
    assert pages.CalcPage.set_precision(neg_val, unc) == 7


def test_set_precision_pos_val(pos_val, unc):
    assert pages.CalcPage.set_precision(pos_val, unc) == 7


@pytest.fixture()
def test_id():
    return 'IVY.v1.0 CHANGE_THIS! (Gain=; Rs=) 26/02/2021 09:02:43'


def test_get_duc_name_from_run_id(test_id):  # FAILS!
    print(type(test_id))
    print(pages.CalcPage.get_duc_name_from_run_id(test_id))
    pages.CalcPage.version = '1.0'
    assert pages.CalcPage.get_duc_name_from_run_id(test_id) == 'CHANGE_THIS!'


@pytest.fixture()
def un_dict():
    return {'value': 123.45,
            'uncert': 0.67,
            'dof': 8,
            'label': 'an_uncert_num'}


def test_build_ureal(un_dict):
    assert isinstance(pages.CalcPage.build_ureal(un_dict), GTC.lib.UncertainReal)

