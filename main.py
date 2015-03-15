# !/usr/bin/python
# -*- coding: utf-8 -*-
import os
import wx, wx.lib.scrolledpanel
import panels
     
        
class Program(wx.Frame):
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, **kwargs)

        self.Center()

        # MENU BAR
        menubar = wx.MenuBar()
        readme = wx.Menu()
        helpme = readme.Append(wx.ID_ANY, '&Help')
        about = readme.Append(wx.ID_ANY, '&About')
        menubar.Append(readme, '&Menu')
        self.SetMenuBar(menubar)

        self.Bind(wx.EVT_MENU, self.about, about)

        # APP BODY
        self.pnl_i = panels.IntroPanel(self)
        self.pnl = panels.XLPanel(self)
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

    def about(self, event):
        msg = "This app will extrapolate count & ImageJ area data obtained from one tenth of the hippocampus to the whole structure.\n" \
              "Calculate counts for the doral/ventral and total dentate gyrus and hilus (multiply values by 10)\n" \
              "Calculate dorsal/ventral and total dentate gyrus and hilar volumes (using Cavalieri's principle)\n" \
              "Calculate cell density based on counts and volume data\n\n" \
              "Direct any questions to Carmen at c_chow568@yahoo.ca"

        dlg = wx.AboutDialogInfo()
        dlg.SetDescription(msg)
    
        wx.AboutBox(dlg)

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
    Program(None)
    app.MainLoop()

if __name__ == '__main__':
    runfile()
        
