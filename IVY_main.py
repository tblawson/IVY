#!python
# -*- coding: utf-8 -*-
"""
DEVELOPMENT VERSION

Created on Mon Jul 31 12:00:00 2017

@author: t.lawson

IVY_main.py - Version 1.0

A Python-3 version of the I-to-V TestPoint application.
This app is intended to offer the same functionality as the original
TestPoint version but avoiding the clutter. It uses a wxPython notebook,
with separate pages (tabs) dedicated to:
* Instrument / file setup,
* Run controls,
* Plotting and
* Analysis

The same data input/output protocol as the original is used, i.e.
initiation parameters are read from the same spreadsheet as the results
are output to.

NOTE: Because the 'Parameters' sheet of the Excel file is interogated twice -
once for obtaining instrument control info (INSTR_DATA) and a second time to
get calibration and uncertainty info (R_INFO and I_INFO), there is redundancy
of information (especially for instruments). Bear in mind that data stored in
INSTR_DATA is just plain numbers or strings, whereas R_INFO and I_INFO can
also contain GTC.ureals.
"""

import os
import wx
from wx.adv import SplashScreen as SplashScreen
import nbpages as page
import IVY_events as evts
import devices
import time
import datetime as dt
import logging

VERSION = "1.0"

print('IVY', VERSION)

# Start logging
logname = 'IVYv'+VERSION+'_'+str(dt.date.today())+'.log'
logfile = os.path.join(os.getcwd(), logname)

fmt = '%(asctime)s %(levelname)s %(name)s:%(funcName)s(L%(lineno)d): '\
      '%(message)s'
datefmt = '%Y-%m-%d %H:%M:%S'
logging.basicConfig(filename=logfile, format=fmt, datefmt=datefmt,
                    level=logging.INFO)
logger = logging.getLogger('IVY.main')
logger.info('\n_____________START____________')


class MainFrame(wx.Frame):
    """
    MainFrame Definition: holds the MainPanel in which the appliction runs
    """
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, size=(900, 600), *args, **kwargs)
        self.data_file = ""
        self.directory = os.getcwd()  # default value
        self.Center()
        self.version = VERSION

        # Event bindings
        self.Bind(evts.EVT_STAT, self.update_status)

        self.sb = self.CreateStatusBar()
        self.sb.SetFieldsCount(2)

        menu_bar = wx.MenuBar()
        file_menu = wx.Menu()

        about = file_menu.Append(wx.ID_ABOUT, '&About', 'About HighResBridgeControl (HRBC)')
        self.Bind(wx.EVT_MENU, self.on_about, about)

        set_dir = file_menu.Append(wx.ID_OPEN, 'Set &Directory',
                                   'Set working directory for raw data and analysis results')
        self.Bind(wx.EVT_MENU, self.on_set_dir, set_dir)  # OnOpen

        file_menu.AppendSeparator()

        exit_app = file_menu.Append(wx.ID_EXIT, '&Quit', 'Exit IVY {}'.format(VERSION))
        self.Bind(wx.EVT_MENU, self.on_quit, exit_app)

        menu_bar.Append(file_menu, "&File")
        self.SetMenuBar(menu_bar)

        # Create a panel to hold the NoteBook...
        self.MainPanel = wx.Panel(self)
        # ... and a Notebook to hold some pages
        self.NoteBook = wx.Notebook(self.MainPanel)

        # Create the page windows as children of the notebook
        self.page1 = page.SetupPage(self.NoteBook)
        self.page2 = page.RunPage(self.NoteBook)
        self.page3 = page.PlotPage(self.NoteBook)
        self.page4 = page.CalcPage(self.NoteBook)

        # Add the pages to the notebook with the label to show on the tab
        self.NoteBook.AddPage(self.page1, "Setup")
        self.NoteBook.AddPage(self.page2, "Run")
        self.NoteBook.AddPage(self.page3, "Plots")
        self.NoteBook.AddPage(self.page4, "Analysis")

        # Finally, put the notebook in a sizer for the panel to manage
        # the layout
        sizer = wx.BoxSizer()
        sizer.Add(self.NoteBook, 1, wx.EXPAND)
        self.MainPanel.SetSizer(sizer)

    def update_status(self, e):
        if e.field == 'b':
            self.sb.SetStatusText(e.msg, 0)
            self.sb.SetStatusText(e.msg, 1)
        else:
            self.sb.SetStatusText(e.msg, e.field)

    def on_about(self, e):
        # A message dialog with 'OK' button. wx.OK is a standard wxWidgets ID.
        dlg_description = "IVY v"+VERSION+": A Python'd version of the TestPoint \
I-to-V converter program for Light Standards."
        dlg_title = "About IVY"
        dlg = wx.MessageDialog(self, dlg_description, dlg_title, wx.OK)
        dlg.ShowModal()  # Show dialog.
        dlg.Destroy()  # Destroy when done.

    def on_set_dir(self, e):
        dlg = wx.DirDialog(self, message='Select working directory',
                           style=wx.DD_DEFAULT_STYLE | wx.DD_CHANGE_DIR)
        dlg.SetPath(os.getcwd())
        if dlg.ShowModal() == wx.ID_OK:
            self.directory = dlg.GetPath()
            self.data_file = os.path.join(self.directory, 'IVY_RunData.json')
            print(self.directory)
            # Get resistor and instrument data:
            devices.RES_DATA, devices.INSTR_DATA = devices.refresh_params(self.directory)
            # Ensure working directory is displayed on SetupPage:
            file_evt = evts.FilePathEvent(Dir=self.directory)
            wx.PostEvent(self.page1, file_evt)
        dlg.Destroy()

    @staticmethod
    def close_instr_sessions():
        for r in devices.ROLES_INSTR.keys():
            devices.ROLES_INSTR[r].close()
            time.sleep(0.1)
        devices.RM.close()
        print('Main.close_instr_sessions(): closed VISA resource manager.')

    def on_quit(self, e):
        self.close_instr_sessions()
        time.sleep(0.1)
        print('Closing IVY...')
        self.Close()


"""__________________End of MainFrame class_______________"""


class IvySplashScreen(SplashScreen):
    def __init__(self, parent=None):
        ivy_bmp = wx.Bitmap(name="ivy-splash.png", type=wx.BITMAP_TYPE_PNG)  # wx.Image(...).ConvertToBitmap()
        splash_style = wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT
        splash_duration = 3000  # milliseconds
        wx.adv.SplashScreen.__init__(self, ivy_bmp, splash_style, splash_duration, parent)


class MainApp(wx.App):
    """Class MainApp."""
    def OnInit(self):
        """Initiate Main App."""
        splash = IvySplashScreen()
        splash.Show()
        frame = MainFrame(None, wx.ID_ANY)  # was self.frame...
        frame.Show(True)  # was self.frame...
        self.SetTopWindow(frame)  # was ...(self.frame)
        frame.SetTitle("IVY v"+VERSION)  # was self.frame...
        return True


if __name__ == '__main__':
    app = MainApp(0)
    app.MainLoop()
