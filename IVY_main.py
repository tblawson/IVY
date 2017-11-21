#!python
# -*- coding: utf-8 -*-
"""
DEVELOPMENT VERSION

Created on Mon Jul 31 12:00:00 2017

@author: t.lawson

IVY_main.py - Version 0.1
A Python version of the I-to-V TestPoint application.
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
import nbpages as page
import IVY_events as evts
import devices
import time

VERSION = "0.1"

print 'IVY', VERSION

""" MainFrame Definition: holds the MainPanel in which the appliction runs"""


class MainFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, size=(900, 500), *args, **kwargs)
        self.version = VERSION
        self.ExcelPath = ""
        self.Center()

        # Event bindings
        self.Bind(evts.EVT_STAT, self.UpdateStatus)

        self.sb = self.CreateStatusBar()
        self.sb.SetFieldsCount(2)

        MenuBar = wx.MenuBar()
        FileMenu = wx.Menu()

        About = FileMenu.Append(wx.ID_ABOUT, text='&About',
                                help='About HighResBridgeControl (HRBC)')
        self.Bind(wx.EVT_MENU, self.OnAbout, About)

        Open = FileMenu.Append(wx.ID_OPEN, text='&Open',
                               help='Open an Excel file')
        self.Bind(wx.EVT_MENU, self.OnOpen, Open)

        Save = FileMenu.Append(wx.ID_SAVE, text='&Save',
                               help='Save data to an Excel file - this \
                               usually happens automatically during a run.')
        self.Bind(wx.EVT_MENU, self.OnSave, Save)

        FileMenu.AppendSeparator()

        Quit = FileMenu.Append(wx.ID_EXIT, text='&Quit',
                               help='Exit HighResBridge')
        self.Bind(wx.EVT_MENU, self.OnQuit, Quit)

        MenuBar.Append(FileMenu, "&File")
        self.SetMenuBar(MenuBar)

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

    def UpdateStatus(self, e):
        if e.field == 'b':
            self.sb.SetStatusText(e.msg, 0)
            self.sb.SetStatusText(e.msg, 1)
        else:
            self.sb.SetStatusText(e.msg, e.field)

    def OnAbout(self, event=None):
        # A message dialog with 'OK' button. wx.OK is a standard wxWidgets ID.
        dlg_description = "IVY v"+VERSION+": A Python'd version of the TestPoint \
        I-to-V converter program for Light Standards."
        dlg_title = "About HighResBridge"
        dlg = wx.MessageDialog(self, dlg_description, dlg_title, wx.OK)
        dlg.ShowModal()  # Show dialog.
        dlg.Destroy()  # Destroy when done.

    def OnSave(self, event=None):
        if self.ExcelPath is not "":
            print 'Main.OnSave(): Saving',
            self.page1.XLFile.GetValue(), '...'
            self.page1.wb.save(self.page1.XLFile.GetValue())
            self.page1.log.close()
        else:
            print 'Main.OnSave(): Nothing to Save.'

    def OnOpen(self, event=None):
        dlg = wx.FileDialog(self, message="Select data file",
                            defaultDir=os.getcwd(),
                            defaultFile="", wildcard="*",
                            style=wx.OPEN | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            self.ExcelPath = dlg.GetPath()
            self.directory = dlg.GetDirectory()
            print self.directory
            print self.ExcelPath
            file_evt = evts.FilePathEvent(XLpath=self.ExcelPath,
                                          d=self.directory, v=VERSION)
            wx.PostEvent(self.page1, file_evt)
        dlg.Destroy()

    def CloseInstrSessions(self, event=None):
        for r in devices.ROLES_INSTR.keys():
            devices.ROLES_INSTR[r].Close()
            time.sleep(0.1)
        devices.RM.close()
        print'Main.CloseInstrSessions(): closed VISA resource manager.'

    def OnQuit(self, event=None):
        self.CloseInstrSessions()
        self.OnSave()
        time.sleep(0.1)
        print 'Closing IVY...'
        self.Close()


"""_______________________________________________"""


class SplashScreen(wx.SplashScreen):
    def __init__(self, parent=None):
        ivy_bmp = wx.Image(name="ivy-splash.png").ConvertToBitmap()
        splashStyle = wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT
        splashDuration = 2000  # milliseconds
        wx.SplashScreen.__init__(self, ivy_bmp, splashStyle,
                                 splashDuration, parent)
        wx.Yield()


class MainApp(wx.App):
    """Class MainApp."""
    def OnInit(self):
        """Initiate Main App."""
        Splash = SplashScreen()
        Splash.Show()
        self.frame = MainFrame(None, wx.ID_ANY)
        self.frame.Show(True)
        self.SetTopWindow(self.frame)
        self.frame.SetTitle("IVY v"+VERSION)
        return True

if __name__ == '__main__':
    app = MainApp(0)
#    wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()
