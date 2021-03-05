# analysis_tests.py

import pytest
from IVY.nbpages import nbpages as pages
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


def test_set_precision_neg_val(analysis_page, neg_val, unc):
    assert analysis_page.set_precision(neg_val, unc) == 7


def test_set_precision_pos_val(pos_val, unc):
    assert pages.CalcPage.set_precision(pos_val, unc) == 7


@pytest.fixture()
def test_id():
    return 'IVY.v1.0 CHANGE_THIS! (Gain=; Rs=) 26/02/2021 09:02:43'


def test_get_duc_name_from_run_id(test_id):  # FAILS!
    print(type(test_id))
    print(pages.CalcPage.get_duc_name_from_run_id(self=None, runid=test_id))
    pages.CalcPage.version = '1.0'
    assert pages.CalcPage.get_duc_name_from_run_id(runid=test_id) == 'CHANGE_THIS!'


@pytest.fixture()
def un_dict():
    return {'value': 123.45,
            'uncert': 0.67,
            'dof': 8,
            'label': 'an_uncert_num'}


def test_build_ureal(un_dict):
    assert isinstance(pages.CalcPage.build_ureal(un_dict), GTC.lib.UncertainReal)


@pytest.fixture()
def alpha():
    return GTC.ureal(1e-3, 1e-6)


@pytest.fixture()
def beta_zero():
    return GTC.ureal(0, 0)


@pytest.fixture()
def beta_nonzero():
    return GTC.ureal(-1e-6, 1e-8)


@pytest.fixture()
def R():
    return GTC.ureal(107, 0.1)


@pytest.fixture()
def R0():
    return GTC.ureal(100, 0.1)


@pytest.fixture()
def T0():
    return GTC.ureal(0, 0.1)


def test_R_to_T_beta_zero(alpha, beta_zero, R, R0, T0):
    T = pages.CalcPage.R_to_T(self=None, alpha=alpha, beta=beta_zero, R=R, R0=R0, T0=T0)
    print(f'T={T.x}')
    assert round(T.x) == 70


def test_R_to_T_beta_nonzero(alpha, beta_nonzero, R, R0, T0):
    T = pages.CalcPage.R_to_T(self=None, alpha=alpha, beta=beta_nonzero, R=R, R0=R0, T0=T0)
    print(f'T={T.x}')
    assert round(T.x) == 76


@pytest.fixture()
def lst():
    return ['a', 'b', 'c']


def test_add_if_unique_item_not_added(lst):
    assert pages.CalcPage.add_if_unique('a', lst) == lst


def test_add_if_unique_item_added(lst):
    assert pages.CalcPage.add_if_unique('d', lst) == ['a', 'b', 'c', 'd']
