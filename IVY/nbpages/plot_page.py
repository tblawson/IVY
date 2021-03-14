# -*- coding: utf-8 -*-
# plot_page.py
""" Defines individual plot page as a panel-like object,
for inclusion in a wx.Notebook object

PYTHON3 DEVELOPMENT VERSION

Created on Fri Mar 5 15:44:30 2021

@author: t.lawson
"""

import logging
import wx
from IVY import IVY_events as Evts
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick
import matplotlib
matplotlib.use('WXAgg')  # Agg renderer for drawing on a wx canvas
matplotlib.rc('lines', linewidth=1, color='blue')

logger = logging.getLogger(__name__)


class PlotPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self.Bind(Evts.EVT_PLOT, self.update_plot)
        self.Bind(Evts.EVT_CLEARPLOT, self.clear_plot)

        self.figure = Figure()

        self.figure.subplots_adjust(hspace=0.3)  # 0.3" v-space b/tw subplots

        # 3high x 1wide, 3rd plot down:
        self.V3ax = self.figure.add_subplot(3, 1, 3)
        # Auto offset to centre on data:
        self.V3ax.ticklabel_format(style='sci', useOffset=False, axis='y',
                                   scilimits=(2, -2))

        # Scientific notation:
        fmt = mtick.ScalarFormatter(useMathText=True, useOffset=False)
        self.V3ax.yaxis.set_major_formatter(fmt)
        # Autoscale with 'buffer' around data extents:
        self.V3ax.autoscale(enable=True, axis='y',
                            tight=False)
        self.V3ax.set_xlabel('time')
        self.V3ax.set_ylabel('V3')

        # 3high x 1wide, 1st plot down:
        self.V1ax = self.figure.add_subplot(3, 1, 1, sharex=self.V3ax)
        self.V1ax.ticklabel_format(useOffset=False,
                                   axis='y')  # Auto o/set to centre on data
        self.V1ax.autoscale(enable=True,
                            axis='y',
                            tight=False)  # Autoscale with data 'buffer'
        plt.setp(self.V1ax.get_xticklabels(),
                 visible=False)  # Hide x-axis labels
        self.V1ax.set_ylabel('V1')
        self.V1ax.set_ylim(auto=True)
        v1_y_ost = self.V1ax.get_xaxis().get_offset_text()
        v1_y_ost.set_visible(False)

        # 3high x 1wide, 2nd plot down:
        self.V2ax = self.figure.add_subplot(3, 1, 2, sharex=self.V3ax)
        self.V2ax.ticklabel_format(useOffset=False,
                                   axis='y')  # Auto offset to centre on data

        # Autoscale with 'buffer' around data extents
        self.V2ax.autoscale(enable=True, axis='y', tight=False)
        plt.setp(self.V2ax.get_xticklabels(),
                 visible=False)  # Hide x-axis labels
        self.V2ax.set_ylabel('V2')
        self.V2ax.set_ylim(auto=True)
        v2_y_ost = self.V2ax.get_xaxis().get_offset_text()
        v2_y_ost.set_visible(False)

        self.canvas = FigureCanvas(self, wx.ID_ANY, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizerAndFit(self.sizer)

    def update_plot(self, e):
        msg_head = 'PlotPage.UpdatePlot(): {}'
        print(msg_head.format('len(t)={}'.format(len(e.t))))
        logger.info(msg_head.format('len(t)={}'.format(len(e.t))))
        print(msg_head.format('{} len(V1)={}', 'len(V3)={}'.format(e.node, len(e.V12), len(e.V3))))
        logger.info(msg_head.format('{} len(V1)={}', 'len(V3)={}'.format(e.node, len(e.V12), len(e.V3))))
        if e.node == 'V1':
            self.V1ax.plot_date(e.t, e.V12, 'bo')
        else:  # V2 data
            self.V2ax.plot_date(e.t, e.V12, 'go')
        self.V3ax.plot_date(e.t, e.V3, 'ro')
        self.figure.autofmt_xdate()  # default settings
        self.V3ax.fmt_xdata = mdates.DateFormatter('%d-%m-%Y, %H:%M:%S')
        self.canvas.draw()
        self.canvas.Refresh()

    def clear_plot(self, e):
        self.V1ax.cla()
        self.V2ax.cla()
        self.V3ax.cla()
        self.V1ax.set_ylabel('V1')
        self.V2ax.set_ylabel('V2')
        self.V3ax.set_ylabel('V3')
        self.canvas.draw()
        self.canvas.Refresh()
