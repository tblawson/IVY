#!python
# -*- coding: utf-8 -*-
"""
# **IVY**
DEVELOPMENT VERSION

Created on Mon Jul 31 12:00:00 2017

@author: t.lawson

IVY_main.py - Version 1.1

## Introduction

A Python-3 version of the I-to-V TestPoint application.
This app is intended to offer the same functionality as the original
TestPoint version but avoiding the clutter. It uses a wxPython notebook,
with separate pages (tabs) dedicated to:
* Instrument / file setup,
* Run controls,
* Plotting and
* Analysis

## Pre-requisites

The following file structure is assumed to be pre-existing before running
the application and the working directory can be set by the File > Set Directory
pull-down menu-item:
```
└─ <working directory> (defaults to the directory where `IVY_main.py` resides)
   ├─ data
   │  ├─ IVY_Instruments.json
   │  └─ IVY_Resistors.json
   └─ log
```
The contents of the 'data' directory should initially consist of at least:

* IVY_Instruments.json (a list of available instruments and external devices) and
* IVY_Resistors.json (a list of resistors).

Example json files can be found in `<project root-directory>\IVY\scripts\data\`.

`data` is also the destination for `IVY_RunData.json` (after initial data acquisition)
and `IVY_Results.json` (after analysis).

log files are written to the `log` directory.

## Setting up a run

Once the working directory has been set, the instrument fields on the **Setup** page can be
populated. List the available visa resources by clicking the 'List Visa res' button.
Each instrument role can be assigned to a physical instrument using the drop-down boxes.
Alternatively a default list can be used by clicking the 'AutoPopulate' button.

Each instrument role has a corresponding 'Test' button to confirm proper operation. The role
description of any GPIB instrument opened in demo mode will (e.g because it is not physically
available) will be displayed in red.

Before proceeding to the Run page, enter the name of the DUC (lower right text-entry field).
The field text will be red until a suitable DUC description has been entered, after which the
text will be green.

## Starting a run

Select the **Run** tab and add a description of the run in the 'comment' field.
Next, click 'Create new run id' to generate a unique run identifier. Enter the device gain and
input resistor *Rs*: this will cause the input resistance to be automatically selected
***if it is no more than 1 MΩ*** - higher-value resistors (10 MΩ to 1 GΩ) must be *manually configured*!

An optional settle delay (in seconds) can also be set to allow the equipment to settle after the user
has left the lab.

A test voltage can be applied at the *V1* position to sanity-check the setup by entering a voltage in the
 'V1 Setting' number control. The 'Node' combo-box allows the user to switch between *V1*, *V2* and *V3*
 (primarily to switch DVM12 between *V1* and *V2* - DVM3 is permanently wired to *V3*).
 'Node' is also used to display which node is currently being measured during a run. Ensure that the voltage
 is reset to zero (by clicking 'Set Zero Volts') before starting the run ('Start run' button).

 Once a run is started the 'Abort run' button becomes enabled and can be used to halt (and abandon) the current run.

When the run is completed the file `IVY_RunData.json` will be written (or appended) to in the working data directory.

## Plots

(Currently not fully implemented)

## Analysis

On the **Analysis** page a list of all run ids contained in `IVY_RunData.json` (in the current data directory)
can be obtained by clicking 'List run IDs'. Any one run_id can then be selected from the list that appears
in the run-list combo-box at the top-centre of the page. Upon selection, a summary of the run appears in the
'Run Summary' scrollable text control on the left.

The selected run can be analysed by clicking 'Analyse' which will fill the environmental conditions text boxes.

Finally, by selecting a nominal output voltage from the `Nom. \u0394Vout` combo box, The measured change in input
current to achieve this output is reported and a full uncertainty budget is also generated. This information is
written to the `IVY_Results.json` file.

**Notepad++** is recommended for editing or inspecting any of the json files.

"""

import os
import wx
from wx.adv import SplashScreen as SplashScreen
import scripts.nbpages.setup_page as setup_p
import scripts.nbpages.run_page as run_p
import scripts.nbpages.plot_page as plot_p
import scripts.nbpages.calc_page as calc_p
from scripts import devices, IVY_events as evts  # import IVY_events as evts
import time
import datetime as dt
import logging

VERSION = "1.1"

print('scripts', VERSION)

# Start logging
logname = 'IVYv'+VERSION+'_'+str(dt.date.today())+'.log'
log_dir = os.path.join(os.path.dirname(os.getcwd()), r'log')
print(f'\nLog dir: {log_dir}\n')
logfile = os.path.join(log_dir, logname)
print(f'Log file: {logfile}\n')

fmt = '%(asctime)s %(levelname)s %(name)s:%(funcName)s(L%(lineno)d): '\
      '%(message)s'
datefmt = '%Y-%m-%d %H:%M:%S'
# logging.basicConfig(filename=logfile, format=fmt, datefmt=datefmt,
#                     level=logging.INFO)
# logger = logging.getLogger(__name__)  # 'scripts/main'
# logger.info('\n_____________START____________')


class MainFrame(wx.Frame):
    """
    MainFrame Definition: holds the MainPanel in which the application runs
    """
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, size=(900, 600), *args, **kwargs)
        self.data_file = ""
        self.log_dir = ""
        self.results_file = ""
        self.directory = os.getcwd()  # default value
        self.Center()
        self.version = VERSION

        # Event bindings
        self.Bind(evts.EVT_STAT, self.update_status)

        self.sb = self.CreateStatusBar()
        self.sb.SetFieldsCount(2)

        menu_bar = wx.MenuBar()
        file_menu = wx.Menu()

        about = file_menu.Append(wx.ID_ABOUT, '&About',
                                 'About IVY v1.1 (I-to-V converter calibration software)')
        self.Bind(wx.EVT_MENU, self.on_about, about)

        set_dir = file_menu.Append(wx.ID_OPEN, 'Set &Directory',
                                   'Set working directory for raw data and analysis results')
        self.Bind(wx.EVT_MENU, self.on_set_dir, set_dir)  # OnOpen

        file_menu.AppendSeparator()

        exit_app = file_menu.Append(wx.ID_EXIT, '&Quit', 'Exit scripts {}'.format(VERSION))
        self.Bind(wx.EVT_MENU, self.on_quit, exit_app)

        menu_bar.Append(file_menu, "&File")
        self.SetMenuBar(menu_bar)

        # Create a panel to hold the NoteBook...
        self.MainPanel = wx.Panel(self)
        # ... and a Notebook to hold some pages
        self.NoteBook = wx.Notebook(self.MainPanel)

        # Create the page windows as children of the notebook
        self.page1 = setup_p.SetupPage(self.NoteBook)
        self.page2 = run_p.RunPage(self.NoteBook)
        self.page3 = plot_p.PlotPage(self.NoteBook)
        self.page4 = calc_p.CalcPage(self.NoteBook)

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
        """An event handler for posting status messages to the status bar."""
        if e.field == 'b':
            self.sb.SetStatusText(e.msg, 0)
            self.sb.SetStatusText(e.msg, 1)
        else:
            self.sb.SetStatusText(e.msg, e.field)

    def on_about(self, e):
        """Self-evident 'about' dialog."""
        # A message dialog with 'OK' button. wx.OK is a standard wxWidgets ID.
        dlg_description = "scripts v"+VERSION+": A Python'd version of the TestPoint \
I-to-V converter program for Light Standards."
        dlg_title = "About scripts"
        dlg = wx.MessageDialog(self, dlg_description, dlg_title, wx.OK)
        dlg.ShowModal()  # Show dialog.
        dlg.Destroy()  # Destroy when done.

    def on_set_dir(self, e):
        """Set the working directory in which the sub-directories 'data' and 'log' can be found."""
        dlg = wx.DirDialog(self, message='Select working directory',
                           style=wx.DD_DEFAULT_STYLE | wx.DD_CHANGE_DIR)
        dlg.SetPath(os.getcwd())
        if dlg.ShowModal() == wx.ID_OK:
            self.directory = dlg.GetPath()
            self.data_file = os.path.join(self.directory, 'data/IVY_RunData.json')
            self.results_file = os.path.join(self.directory, 'data/IVY_Results.json')
            print(f'Working directory: {self.directory}')
            # Get resistor and instrument data:
            devices.RES_DATA, devices.INSTR_DATA = devices.refresh_params(self.directory)
            # Ensure working directory is displayed on SetupPage:
            file_evt = evts.FilePathEvent(Dir=self.directory)
            wx.PostEvent(self.page1, file_evt)

            # Create IV-box instrument once:
            d = 'IV_box'  # Description
            r = 'IVbox'  # Role
            self.page1.create_instr(d, r)
            self.page1.build_combo_choices()

        dlg.Destroy()

    @staticmethod
    def close_instr_sessions():
        """Close all open instruments and external devices."""
        for r in devices.ROLES_INSTR.keys():
            devices.ROLES_INSTR[r].close()
            time.sleep(0.1)
        devices.RM.close()
        print('Main.close_instr_sessions(): closed VISA resource manager.')

    def on_quit(self, e):
        """Exit scripts."""
        self.close_instr_sessions()
        time.sleep(0.1)
        print('Closing scripts...')
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
        frame.SetTitle("scripts v"+VERSION)  # was self.frame...
        return True


if __name__ == '__main__':
    app = MainApp(0)
    app.MainLoop()
