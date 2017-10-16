# -*- coding: utf-8 -*-
"""
Created on Tue Jun 30 14:31:53 2015

@author: t.lawson
"""
"""
HighRes_events.py
Definitions of event types - since the GUI makes use of events to monitor
the status of widgets (buttons, displays,etc.), use of events is a natural
fit and guarantees a thread-safe means of passing information from the
acquisition thread to the main GUI.
"""

import wx.lib.newevent

# Event used to pass an updated string to the 'comment' TextCtrl on RunPage
UpdateCommentEvent, EVT_UPDATE_COM_STR = wx.lib.newevent.NewEvent()

# Event to pass new data back to the RunPage displays or PlotPage
DataEvent, EVT_DATA = wx.lib.newevent.NewEvent()

# Event to pass new data back to the PlotPage
PlotEvent, EVT_PLOT = wx.lib.newevent.NewEvent()

# Event to clear subplots on the PlotPage
ClearPlotEvent, EVT_CLEARPLOT = wx.lib.newevent.NewEvent()

# Event to pass massages back to MainFrame, to update status bar
StatusEvent, EVT_STAT = wx.lib.newevent.NewEvent()

# Event to update RunPage start_row display 
StartRowEvent, EVT_START_ROW = wx.lib.newevent.NewEvent()

# Event to update RunPage stop_row display
StopRowEvent, EVT_STOP_ROW = wx.lib.newevent.NewEvent()

# Event to update RunPage row display
RowEvent, EVT_ROW = wx.lib.newevent.NewEvent()

# Event to update RunPage delay displays
DelaysEvent, EVT_DELAYS = wx.lib.newevent.NewEvent()

# Event to update Run Id
#RunIdEvent, EVT_RUNID = wx.lib.newevent.NewEvent()

# Event to update file path text_ctrl on SetupPage
FilePathEvent, EVT_FILEPATH = wx.lib.newevent.NewEvent()

# Event to update Switchbox config (description)
SB_ConfEvent, EVT_SBCONF = wx.lib.newevent.NewEvent()

# Event to update log file
LogEvent, EVT_LOG = wx.lib.newevent.NewEvent()