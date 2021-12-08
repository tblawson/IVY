# -*- coding: utf-8 -*-
# calc_page.py
""" Defines individual calc page as a panel-like object,
for inclusion in a wx.Notebook object

PYTHON3 DEVELOPMENT VERSION

Created on Fri Mar 5 16:03:30 2021

@author: t.lawson
"""

import logging
import wx
import time
import datetime as dt
import math
import json
import GTC
import os

from scripts import devices

DELTA = u'\N{GREEK CAPITAL LETTER DELTA}'
PT_T_DEF_UNCERT = 0.5
GMH_T_DEF_UNCERT = 0.5

logger = logging.getLogger(__name__)


class CalcPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.version = self.GetTopLevelParent().version
        self.data_dir = self.GetTopLevelParent().directory
        self.data_file = self.GetTopLevelParent().data_file
#        self.results_file = os.path.join(os.path.dirname(self.data_file),
#                                         'IVY_Results.json')
        self.Results = {}  # Accumulate run-analyses here

        self.Rs_VALUES = self.GetTopLevelParent().page2.Rs_VALUES
        self.Rs_NAMES = ['IV1k 1k', 'IV10k 10k',
                         'IV100k 100k', 'IV1M 1M', 'IV10M 10M',
                         'IV100M 100M', 'IV1G 1G']
        self.Rs_VAL_NAME = dict(zip(self.Rs_VALUES, self.Rs_NAMES))

        gb_sizer = wx.GridBagSizer()

        # Analysis set-up:
        self.ListRuns_btn = wx.Button(self, id=wx.ID_ANY, label='List run IDs')
        self.ListRuns_btn.Bind(wx.EVT_BUTTON, self.on_list_runs)
        gb_sizer.Add(self.ListRuns_btn, pos=(0, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.RunID_cb = wx.ComboBox(self, id=wx.ID_ANY,
                                    style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.RunID_cb.Bind(wx.EVT_COMBOBOX, self.on_run_choice)
        self.RunID_cb.Bind(wx.EVT_TEXT, self.on_run_choice)
        gb_sizer.Add(self.RunID_cb, pos=(0, 1), span=(1, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.Analyze_btn = wx.Button(self, id=wx.ID_ANY, label='Analyze')
        self.Analyze_btn.Bind(wx.EVT_BUTTON, self.on_analyze)
        gb_sizer.Add(self.Analyze_btn, pos=(0, 7), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)  #

        # -----------------------------------------------------------------
        h_sep1 = wx.StaticLine(self, id=wx.ID_ANY, size=(720, 1),
                                    style=wx.LI_HORIZONTAL)
        gb_sizer.Add(h_sep1, pos=(1, 0), span=(1, 8),
                     flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------

        # Run summary:
        run_info_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Run Summary:')
        gb_sizer.Add(run_info_lbl, pos=(2, 0), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        self.RunSummary_txtctrl = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_MULTILINE |
                                                                        wx.TE_READONLY | wx.HSCROLL,
                                              size=(250, 1))
        gb_sizer.Add(self.RunSummary_txtctrl, pos=(3, 0), span=(20, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)

        # Analysis results:
        # Headings & environment readings
        quantity_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Quantity')
        gb_sizer.Add(quantity_lbl, pos=(2, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        ureal_summary_lbl = wx.StaticText(self, id=wx.ID_ANY, size=(200, 1),
                                          label='Val, Unc, DoF')
        gb_sizer.Add(ureal_summary_lbl, pos=(2, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        k_lbl = wx.StaticText(self, id=wx.ID_ANY, label='k(95%)')
        gb_sizer.Add(k_lbl, pos=(2, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        exp_u_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Exp Uncert')
        gb_sizer.Add(exp_u_lbl, pos=(2, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        p_lbl = wx.StaticText(self, id=wx.ID_ANY, label='Pressure:')
        gb_sizer.Add(p_lbl, pos=(3, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.PSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.PSummary, pos=(3, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.Pk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.Pk, pos=(3, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.PExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.PExpU, pos=(3, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        rh_lbl = wx.StaticText(self, id=wx.ID_ANY, label='%RH:')
        gb_sizer.Add(rh_lbl, pos=(4, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHSummary, pos=(4, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHk, pos=(4, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.RHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.RHExpU, pos=(4, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        tgmh_lbl = wx.StaticText(self, id=wx.ID_ANY, label='T (GMH):')
        gb_sizer.Add(tgmh_lbl, pos=(5, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHSummary = wx.TextCtrl(self, id=wx.ID_ANY,
                                       style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHSummary, pos=(5, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHk = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHk, pos=(5, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.TGMHExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.TGMHExpU, pos=(5, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND
        # -----------------------------------------------------------------
        self.h_sep3 = wx.StaticLine(self, id=wx.ID_ANY, size=(480, 1),
                                    style=wx.LI_HORIZONTAL)
        gb_sizer.Add(self.h_sep3, pos=(6, 2), span=(1, 5),
                     flag=wx.ALL | wx.EXPAND, border=5)  #
        # -----------------------------------------------------------------
        vout_lbl = wx.StaticText(self, id=wx.ID_ANY,
                                 label='Nom. ' + DELTA + 'Vout:')
        gb_sizer.Add(vout_lbl, pos=(7, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.NomVout = wx.ComboBox(self, id=wx.ID_ANY,
                                   style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.NomVout.Bind(wx.EVT_COMBOBOX, self.on_nom_vout_choice)
        gb_sizer.Add(self.NomVout, pos=(7, 3), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)

        i_in_lbl = wx.StaticText(self, id=wx.ID_ANY, label=DELTA + 'I_in:')
        gb_sizer.Add(i_in_lbl, pos=(8, 2), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.IinSummary = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.IinSummary, pos=(8, 3), span=(1, 2),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.Iink = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.Iink, pos=(8, 5), span=(1, 1),
                     flag=wx.ALL | wx.EXPAND, border=5)
        self.IinExpU = wx.TextCtrl(self, id=wx.ID_ANY, style=wx.TE_READONLY)
        gb_sizer.Add(self.IinExpU, pos=(8, 6), span=(1, 1),
                     flag=wx.ALL, border=5)  # | wx.EXPAND

        self.Budget = wx.TextCtrl(self, style=wx.TE_MULTILINE |
                                  wx.TE_READONLY | wx.HSCROLL, id=wx.ID_ANY,)
        budget_font = wx.Font(8, wx.MODERN, wx.NORMAL, wx.NORMAL)
        self.Budget.SetFont(budget_font)
        gb_sizer.Add(self.Budget, pos=(9, 2), span=(14, 6),
                     flag=wx.ALL | wx.EXPAND, border=5)  #

        self.SetSizerAndFit(gb_sizer)

        self.data_file = None
        self.run_data = None
        self.run_IDs = []
        self.runstr = ''
        self.run_ID = ''
        self.ThisResult = {}
        self.nom_Vout = {}
        self.budget_table_sorted = {'pos': [], 'neg': []}
        self.result_row = []
        self.results_file = ''

    def on_list_runs(self, e):
        """
        Open the .json data file and de-serialize into the dictionary
        self.run_data for later access. Update the choices in the 'Run ID'
        widget to the run ids (used as primary keys in self.run_data).
        """
        self.data_file = self.GetTopLevelParent().data_file
        print(f'Attempting to open "{self.data_file}"')
        with open(self.data_file, 'r') as in_file:
            self.run_data = json.load(in_file)
            self.run_IDs = list(self.run_data.keys())
            print(f'Found the following runs:\n{self.run_IDs}')

        self.RunID_cb.Clear()
        self.RunID_cb.AppendItems(self.run_IDs)
        self.RunID_cb.SetSelection(0)

    def on_run_choice(self, e):
        id = e.GetString()
        self.runstr = json.dumps(self.run_data[id], indent=4)
        self.RunSummary_txtctrl.SetValue(self.runstr)
        self.PSummary.Clear()
        self.Pk.Clear()
        self.PExpU.Clear()
        self.RHSummary.Clear()
        self.RHk.Clear()
        self.RHExpU.Clear()
        self.TGMHSummary.Clear()
        self.TGMHk.Clear()
        self.TGMHExpU.Clear()
        self.NomVout.Clear()  # Added recently.
        self.IinSummary.Clear()
        self.Iink.Clear()
        self.IinExpU.Clear()
        self.Budget.Clear()

    def on_analyze(self, e):
        self.run_ID = self.RunID_cb.GetValue()
        this_run = self.run_data[self.run_ID]

        logger.info('STARTING ANALYSIS...')

        # Correction for Pt-100 sensor DVM:
        dvmt = this_run['Instruments']['DVMT']
        dvmt_cor = self.build_ureal(devices.INSTR_DATA[dvmt]['correction_100r'])

        """
        Pt sensor is a few cm away from input resistors, so assume a
        fairly large type B Tdef of 0.5 deg C:
        """
        pt_t_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](PT_T_DEF_UNCERT),
                             3, label='Pt_T_def')

        pt_alpha = self.build_ureal(devices.RES_DATA['Pt 100r']['alpha'])
        pt_beta = self.build_ureal(devices.RES_DATA['Pt 100r']['beta'])
        pt_r0 = self.build_ureal(devices.RES_DATA['Pt 100r']['R0_LV'])
        pt_t_ref = self.build_ureal(devices.RES_DATA['Pt 100r']['TRef_LV'])

        """
        GMH sensor is a few cm away from DUC which, itself, has a size of
        several cm, so assume a fairly large type B Tdef of 0.5 deg C:
        """
        gmh_t_def = GTC.ureal(0, GTC.type_b.distribution['gaussian'](GMH_T_DEF_UNCERT),
                              3, label='GMH_T_def')

        comment = this_run['Comment']
        duc_name = self.get_duc_name_from_run_id(self.run_ID)
        duc_gain = this_run['DUC_G']
        mean_date = self.get_mean_date()
        processed_date = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        print('Comment:', comment)
        print('Run_Id:', self.run_ID)
        print('gain ={}'.format(duc_gain))
        print('Mean_date:', mean_date)
        logger.info('Comment: %s\nRun_ID: %s\ngain = %s\nMean_date: %s',
                    comment, self.run_ID, duc_gain, mean_date)

        # Determine mean env. conditions
        gmh_temps = []
        gmh_room_rhs = []
        gmh_room_ps = []
        for T in this_run['T_GMH']:
            gmh_temps.append(T[0])
        for RH in this_run['Room_conds']['RH']:
            gmh_room_rhs.append(RH[0])
        for P in this_run['Room_conds']['P']:
            gmh_room_ps.append(P[0])

        d = this_run['Instruments']['GMH']
        t_gmh_cor = self.build_ureal(devices.INSTR_DATA[d]['T_correction'])
        t_gmh_raw = GTC.ta.estimate_digitized(gmh_temps, 0.01)
        t_gmh = t_gmh_raw + t_gmh_cor + gmh_t_def
        t_gmh_k = GTC.rp.k_factor(t_gmh.df)
        t_gmh_eu = t_gmh.u*t_gmh_k

        d = this_run['Instruments']['GMHroom']
        rh_cor = self.build_ureal(devices.INSTR_DATA[d]['RH_correction'])
        rh_raw = GTC.ta.estimate_digitized(gmh_room_rhs, 0.1)
        rh = rh_raw*(1 + rh_cor)
        rh_k = GTC.rp.k_factor(rh.df)
        rh_eu = rh.u*rh_k

        # Re-use d (same instrument description)
        press_cor = self.build_ureal(devices.INSTR_DATA[d]['P_correction'])
        press_raw = GTC.ta.estimate_digitized(gmh_room_ps, 0.1)
        press = press_raw * (1 + press_cor)
        press_k = GTC.rp.k_factor(press.df)
        press_eu = press.u * press_k

        self.Results.update({self.run_ID: {}})
        self.ThisResult = self.Results[self.run_ID]
        self.ThisResult.update({'Comment': comment,
                                'Date': mean_date,
                                'Processed date': processed_date,
                                'DUC name': duc_name,
                                'DUC gain': duc_gain,
                                'T_GMH': {},
                                'RH': {},
                                'P': {},
                                'Nom_dV': {}})

        self.ThisResult['T_GMH'].update({'value': t_gmh.x,
                                         'uncert': t_gmh.u,
                                         'dof': t_gmh.df,
                                         'label': t_gmh.label,
                                         'k': t_gmh_k,
                                         'ExpU': t_gmh_eu})

        self.ThisResult['RH'].update({'value': rh.x,
                                      'uncert': rh.u,
                                      'dof': rh.df,
                                      'label': rh.label,
                                      'k': rh_k,
                                      'ExpU': rh_eu})

        self.ThisResult['P'].update({'value': press.x,
                                     'uncert': press.u,
                                     'dof': press.df,
                                     'label': press.label,
                                     'k': press_k,
                                     'ExpU': press_eu})

        self.PSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(press.x, press.u, press.df))  # str(press.s)
        self.Pk.SetValue('{0:.1f}'.format(press_k))
        self.PExpU.SetValue('{0:.2f}'.format(press_eu))
        self.RHSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(rh.x, rh.u, rh.df))  # str(rh.s)
        self.RHk.SetValue('{0:.1f}'.format(rh_k))
        self.RHExpU.SetValue('{0:.2f}'.format(rh_eu))
        self.TGMHSummary.SetValue('{0:.3f} +/- {1:.3f}. dof={2:.1f}'.format(t_gmh.x, t_gmh.u, t_gmh.df))  # str(t_gmh.s)
        self.TGMHk.SetValue('{0:.1f}'.format(t_gmh_k))
        self.TGMHExpU.SetValue('{0:.2f}'.format(t_gmh_eu))

        influencies = []
        v1s = []
        v2s = []
        v3s = []

        num_rows = len(this_run['Nom_Vout'])
        for row in range(0, num_rows, 8):  # 0, [8, [16]]
            gains = []  # gains = set()
            # 'neg' and 'pos' refer to polarity of OUTPUT VOLTAGE, not
            # input current! nom_Vout = +/-( 0.1,[1,[10]] ):
            self.nom_Vout = {'pos': this_run['Nom_Vout'][row+2],
                             'neg': this_run['Nom_Vout'][row+1]}
            abs_nom_vout = self.nom_Vout['pos']

            # Construct ureals from raw voltage data, including gain correction
            for n in range(4):
                label_suffix_1 = this_run['Node'][row+n]+'_'+str(n)
                label_suffix_2 = this_run['Node'][row+4+n]+'_'+str(n)
                label_suffix_3 = 'V3' + '_' + str(n)

                v1_v = this_run['IP_V']['val'][row+n]
                v1_u = this_run['IP_V']['sd'][row+n]
                v1_dof = this_run['Nreads'] - 1
                v1_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_1

                d1 = this_run['Instruments']['DVM12']
                gain_param = self.get_gain_err_param(v1_v)
                print(devices.INSTR_DATA[d1].keys())
                gain = self.build_ureal(devices.INSTR_DATA[d1][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v1_raw = GTC.ureal(v1_v, v1_u, v1_dof, label=v1_label)
                v1s.append(GTC.result(v1_raw/gain))

                v2_v = this_run['IP_V']['val'][row+4+n]
                v2_u = this_run['IP_V']['sd'][row+4+n]
                v2_dof = this_run['Nreads'] - 1
                v2_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_2

                d2 = d1  # Same DVM
                gain_param = self.get_gain_err_param(v2_v)
                gain = self.build_ureal(devices.INSTR_DATA[d2][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v2_raw = GTC.ureal(v2_v, v2_u, v2_dof, label=v2_label)
                v2s.append(GTC.result(v2_raw/gain))

                v3_v = this_run['OP_V']['val'][row+n]
                v3_u = this_run['OP_V']['sd'][row+n]
                v3_dof = this_run['Nreads'] - 1
                v3_label = 'OP' + str(abs_nom_vout) + '_' + label_suffix_3

                d3 = this_run['Instruments']['DVM3']
                gain_param = self.get_gain_err_param(v3_v)
                gain = self.build_ureal(devices.INSTR_DATA[d3][gain_param])
                gains = self.add_if_unique(gain, gains)  # gains.add(gain)
                v3_raw = GTC.ureal(v3_v, v3_u, v3_dof, label=v3_label)
                v3s.append(GTC.result(v3_raw/gain))

                influencies.extend([v1_raw, v2_raw, v3_raw])

            influencies.extend(gains)  # List of unique gain corrections.
            print('list of gains:')
            for g in gains:
                print('{} +/- {}, dof={}'.format(g.x, g.u, g.df))

            # Offset-adjustment
            v1_pos = GTC.result(v1s[2] - (v1s[0] + v1s[3]) / 2)
            v1_neg = GTC.result(v1s[1] - (v1s[0] + v1s[3]) / 2)
            v2_pos = GTC.result(v2s[2] - (v2s[0] + v2s[3]) / 2)
            v2_neg = GTC.result(v2s[1] - (v2s[0] + v2s[3]) / 2)
            v3_pos = GTC.result(v3s[2] - (v3s[0] + v3s[3]) / 2)
            v3_neg = GTC.result(v3s[1] - (v3s[0] + v3s[3]) / 2)

            # V-drop across Rs
            v_rs_pos = GTC.result(v1_pos - v2_pos)
            v_rs_neg = GTC.result(v1_neg - v2_neg)
            assert v_rs_pos.x * v_rs_neg.x < 0, 'V_Rs polarity error!'

            # Rs Temperature
            t_rs = []
            pt_r_cor = []
            for n, R_raw in enumerate(this_run['Pt_DVM'][row:row+8]):
                pt_r_cor.append(GTC.result(R_raw * (1 + dvmt_cor),
                                           label='Pt_Rcor'+str(n)))
                t_rs.append(GTC.result(self.R_to_T(alpha=pt_alpha, beta=pt_beta,
                                                   R=pt_r_cor[n],
                                                   R0=pt_r0, T0=pt_t_ref)))

            av_t_rs = GTC.result(GTC.fn.mean(t_rs),
                                 label='av_T_Rs'+str(abs_nom_vout))
            influencies.extend(pt_r_cor)
            influencies.extend([pt_alpha, pt_beta, pt_r0, pt_t_ref,
                                dvmt_cor, pt_t_def])  # av_T_Rs
            assert pt_alpha in influencies, 'Influencies missing Pt_alpha!'
            assert pt_beta in influencies, 'Influencies missing Pt_beta!'
            assert pt_r0 in influencies, 'Influencies missing Pt_R0!'
            assert pt_t_ref in influencies, 'Influencies missing Pt_TRef!'
            assert dvmt_cor in influencies, 'Influencies missing DVMT_cor!'
            assert pt_t_def in influencies, 'Influencies missing Pt_T_def!'

            # Value of Rs
            nom_Rs = this_run['Rs']
            msg = 'Nominal Rs value: {0}, '\
                  'Abs. Nom. Vout: {1:.1f}'.format(nom_Rs, abs_nom_vout)
            print( '\n', msg, '\n')
            logger.info(msg)
            rs_name = self.Rs_VAL_NAME[nom_Rs]
            rs_0 = self.build_ureal(devices.RES_DATA[rs_name]['R0_LV'])
            rs_t_ref = self.build_ureal(devices.RES_DATA[rs_name]['TRef_LV'])
            rs_alpha = self.build_ureal(devices.RES_DATA[rs_name]['alpha'])
            rs_beta = self.build_ureal(devices.RES_DATA[rs_name]['beta'])

            # Correct Rs value for temperature
            delta_t = GTC.result(av_t_rs - rs_t_ref + pt_t_def)
            rs = GTC.result(rs_0 * (1 + rs_alpha * delta_t + rs_beta * delta_t ** 2))

            influencies.extend([rs_0, rs_alpha, rs_beta, rs_t_ref])

            '''
            Finally, calculate current-change in,
            for nominal voltage-change out:
            REMEMBER: POSITIVE 'Vout' is caused by 'I_pos'
            (which may be NEGATIVE for inverting device)!
            '''
            i_in_pos = GTC.result(v_rs_pos / rs)
            i_in_neg = GTC.result(v_rs_neg / rs)
            assert i_in_pos.x * i_in_neg.x < 0, 'Iin polarity error!'

            # Calculate I_in that would produce nominal Vout:
            i_pos = GTC.result(i_in_pos * self.nom_Vout['pos'] / v3_pos)
            i_neg = GTC.result(i_in_neg * self.nom_Vout['neg'] / v3_neg)
            assert i_pos.x * i_neg.x < 0, 'I polarity error!'

            this_result = {'pos': i_pos, 'neg': i_neg}

            # build uncertainty budget table
            budget_table = {'pos': [], 'neg': []}
            for i in influencies:
                print('Working through influence variables: {}.'.format(i.label))
                logger.info('Working through influence variables: {}.'.format(i.label))
                if i.u == 0:
                    sensitivity = {'pos': 0, 'neg': 0}
                else:
                    sensitivity = {'pos': GTC.component(i_pos, i)/i.u,
                                   'neg': GTC.component(i_neg, i)/i.u}

                # Only include non-zero influencies:
                if abs(GTC.component(i_pos, i)) > 0:
                    print('Included component of I+: {}'.format(GTC.component(i_pos, i)))
                    logger.info('Included component of I+: {}'.format(GTC.component(i_pos, i)))
                    budget_table['pos'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['pos'],
                                                GTC.component(i_pos, i)])
                else:
                    print('ZERO COMPONENT of I+')
                    logger.info('ZERO COMPONENT of I+')

                if abs(GTC.component(i_neg, i)) > 0:
                    print('Included component of I-: {}'.format(GTC.component(i_neg, i)))
                    logger.info('Included component of I-: {}'.format(GTC.component(i_neg, i)))
                    budget_table['neg'].append([i.label, i.x, i.u, i.df,
                                                sensitivity['neg'],
                                                GTC.component(i_neg, i)])
                else:
                    print('ZERO COMPONENT of I-')
                    logger.info('ZERO COMPONENT of I-')

            # self.budget_table_sorted = {'pos': [], 'neg': []}
            self.budget_table_sorted['pos'] = sorted(budget_table['pos'],
                                                     key=self.by_u_cont,
                                                     reverse=True)
            self.budget_table_sorted['neg'] = sorted(budget_table['neg'],
                                                     key=self.by_u_cont,
                                                     reverse=True)

            # Write results (including budgets)
            self.result_row = self.write_this_result(this_result)
            time.sleep(0.1)
            row += 8
            del influencies[:]
            del v1s[:]
            del v2s[:]
            del v3s[:]
        # <-- End of analysis loop for this run

        # Save analysis result
        self.results_file = self.GetTopLevelParent().results_file
        with open(self.results_file, 'w') as results_fp:
            json.dump(self.Results, results_fp, indent=4)

        # Ensure Vout c-box is cleared before updating choices.
        # Vout choices should be chosen from: (0.1, -0.1, 1, -1, 10, -10):
        self.NomVout.Clear()
        for V in self.ThisResult['Nom_dV'].keys():
            self.NomVout.Append(str(V))
        # ______________END OF OnAnalyze()________________

    def on_nom_vout_choice(self, e):
        """
        Select which analysed data sub-set (by Vout) to display
        - Delta_Iin and
        - Uncert budget
        """
        v_out = float(e.GetString())
        i_in = self.build_ureal(self.ThisResult['Nom_dV'][v_out]['Delta_Iin'])
        i_in_str = '{0:.5e}, u={1:.1e}, df={2:.1f}'.format(i_in.x, i_in.u, i_in.df)
        i_in_k = GTC.rp.k_factor(i_in.df)
        i_in_EU = i_in.u * i_in_k
        self.IinSummary.SetValue(i_in_str)
        self.Iink.SetValue('{0:.1f}'.format(i_in_k))
        self.IinExpU.SetValue('{0:.2e}'.format(i_in_EU))

        # Build uncertainty budget table (as a string for display):
        budget_str = '{:28}{:<14}{:<10}{:<6}{:<10}{:<13}\n'\
            .format('Quantity', 'Value', 'Std u.', 'dof',
                    'Sens. Co.', 'U. Cont.')
        u_budget_dict = self.ThisResult['Nom_dV'][v_out]['U_Budget']
        for n, q in enumerate(u_budget_dict['quantity(label)']):
            v = u_budget_dict['value'][n]
            u = u_budget_dict['std. uncert'][n]
            d = float(u_budget_dict['dof'][n])
            s = u_budget_dict['sens. co.'][n]
            c = u_budget_dict['uncert. cont.'][n]
            line = '{0:<28}{1:<16.{2}g}{3:<10.2g}{4:<6.1f}{5:<10.1e}{6:<13.1e}\n'\
                .format(q, v, self.set_precision(v, u), u, d, s, c)
            budget_str += line
        self.Budget.SetValue(budget_str)

    @staticmethod
    def set_precision(v, u):
        if v == 0:
            v_lg = 0
        else:
            v_lg = math.log10(abs(v))
        if u == 0:
            u_lg = 0
        else:
            u_lg = math.log10(u)
        if abs(v) > u:
            return int(round(v_lg - u_lg)) + 2
        else:
            return 2

    def get_duc_name_from_run_id(self, runid):
        start = 'scripts.v{} '.format(self.version)
        end = ' (Gain='
        return runid[len(start): runid.find(end)]

    def get_mean_date(self):
        """
        Accept a list of times (as strings).
        Return a mean time (as a string).
        """
        time_fmt = '%d/%m/%Y %H:%M:%S'
        t_av = 0.0
        t_lst = self.run_data[self.run_ID]['Date_time']

        for t_str in t_lst:
            t_dt = dt.datetime.strptime(t_str, time_fmt)
            t_tup = dt.datetime.timetuple(t_dt)  # A Python time tuple object
            t_av += time.mktime(t_tup)  # time as float (seconds from epoch)

        t_av /= len(t_lst)
        t_av_dt = dt.datetime.fromtimestamp(t_av)
        return t_av_dt.strftime(time_fmt)  # av. time as string

    @staticmethod
    def build_ureal(d):
        """
        Accept a dictionary containing keys of 'value', 'uncert', 'dof' and
        'label' and use corresponding values as input to GTC.ureal().
        Return resulting ureal. Ignore any other keys.
        """
        try:
            assert isinstance(d, dict), 'Not a dictionary!'
            return GTC.ureal(float(d['value']),
                             float(d['uncert']),
                             float(d['dof']),
                             str(d['label']))
        except (AssertionError, KeyError) as msg:
            # raise TypeError('Non-ureal input')
            print('build_ureal():', msg)
            return 0

    @staticmethod
    def get_gain_err_param(v):
        """
        Return the key (a string) identifying the correct gain parameter
        for V (and appropriate range).
        """
        if abs(v) < 0.001:
            nom_v = '0.0001'
            nom_range = '0.1'
        elif abs(v) < 0.022:
            nom_v = '0.01'
            nom_range = '0.1'
        elif abs(v) < 0.071:
            nom_v = '0.05'
            nom_range = '0.1'
        elif abs(v) < 0.22:
            nom_v = '0.1'
            nom_range = '0.1'
        elif abs(v) < 0.71:
            nom_v = '0.5'
            nom_range = '1'
        elif abs(v) < 2.2:
            nom_v = '1'
            nom_range = '1'
        else:
            nom_v = nom_range = str(int(abs(round(v))))  # '1' or '10'
        gain_param = f'Vgain_{nom_v}r{nom_range}'
        return gain_param

    @staticmethod
    def R_to_T(alpha, beta, R, R0, T0):
        """
        Convert a resistive T-sensor reading from resistance to temperature.
        All arguments and return value are ureals.
        """
        if beta.x == 0 and beta.u == 0:  # No 2nd-order T-Co
            T = GTC.result((R/R0 - 1)/alpha + T0)
        else:
            a = GTC.result(beta)
            b = GTC.result(alpha - 2*beta*T0, True)  # Why is label set to True?
            c = GTC.result(1 - alpha*T0 + beta*T0**2 - (R/R0))
            T = GTC.result((-b + GTC.sqrt(b**2 - 4*a*c))/(2*a))
        return T

    @staticmethod
    def by_u_cont(line):
        """
        A function required for budget-table sorting
        by uncertainty contribution.
        """
        return line[5]

    def write_this_result(self, result):
        """
        Write results and uncert. budget for THIS NOM_VOUT (BOTH polarities)
        """
        result_dict = {'pos': {'Delta_Iin': {}, 'U_Budget': {}},
                       'neg': {'Delta_Iin': {}, 'U_Budget': {}}}

        # Summary for each polarity:
        for polarity in result.keys():  # 'pos' or 'neg'
            value = result[polarity].x
            uncert = result[polarity].u
            dof = result[polarity].df
            lbl = result[polarity].label
            k = GTC.rp.k_factor(dof)
            exp_u = k * uncert

            # Delta_I_in:
            result_dict[polarity]['Delta_Iin'].update({'value': value,
                                                       'uncert': uncert,
                                                       'dof': dof,
                                                       'label': lbl,
                                                       'k': k, 'EU': exp_u})
            # Uncert Budget:
            quantity = []
            v = []
            u = []
            df = []
            sens = []
            cont = []
            for line in self.budget_table_sorted[polarity]:
                for i, heading in enumerate((quantity, v, u, df, sens, cont)):
                    heading.append(line[i])
            result_dict[polarity]['U_Budget'].update({'quantity(label)': quantity,
                                                      'value': v,
                                                      'std. uncert': u,
                                                      'dof': df,
                                                      'sens. co.': sens,
                                                      'uncert. cont.': cont})
            d = {self.nom_Vout[polarity]: result_dict[polarity]}
            self.ThisResult['Nom_dV'].update(d)
        return 1

    @staticmethod
    def add_if_unique(item, lst):
        """
        Append 'item' to 'lst' only if it is not already present.
        """
        if item not in lst:
            lst.append(item)
        return lst
