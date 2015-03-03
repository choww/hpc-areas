# !/usr/bin/python
# -*- coding: utf-8 -*-
import wx, wx.lib.scrolledpanel
from panels import *
     
        
class Count(wx.Frame):
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, **kwargs)

        self.Center()

        # APP BODY
        self.pnl_i = IntroPanel(self)
        self.pnl = XLPanel(self)
        self.pnl.Hide()

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.pnl_i, 0, wx.EXPAND)
        sizer.Add(self.pnl, 0, wx.ALL)

        self.b_next = wx.Button(self, label="Next")
        self.b_next.Bind(wx.EVT_BUTTON, self.switch_pnl)
        
        sizer_i = wx.BoxSizer(wx.VERTICAL)
        sizer_i.Add(self.b_next)
        sizer.Add(sizer_i, 0, wx.ALL|wx.CENTER, 10)

        self.SetSizer(sizer)
        sizer.Fit(self)
        
        self.Show()

    def switch_pnl(self, event):
        if self.pnl_i.IsShown():
            self.pnl_i.Hide()
            self.b_next.Hide()
            self.pnl.Show()
        else:
            self.pnl_i.Show()
            self.pnl.Hide()
        self.Fit()
        

def runfile():
    app = wx.App()
    Count(None)
    app.MainLoop()

if __name__ == '__main__':
    runfile()
        
