# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import wx, wx.lib.scrolledpanel

class AboutDlg(wx.Dialog):
    """
    Info regarding the program 
    """
    def __init__(self, *args, **kwargs):
        wx.Dialog.__init__(self, *args, **kwargs)

        self.Center()
        self.SetSize((450,350))
        pnl = wx.lib.scrolledpanel.ScrolledPanel(self)
        pnl.SetupScrolling(scroll_x=False)

##        txt0 = "This program extrapolates cell count & ImageJ " \
##               "area data obtained from one tenth of  " \
##               "the hippocampus to the whole structure. \n" \
##               "Calculations used: "
##
##        t_about = wx.StaticText(pnl, label=txt0)
##        t_about.Wrap(385)

        def we_are_frozen():
            return hasattr(sys,"frozen")
        def img_dir():
            if we_are_frozen():
                return os.path.dirname(unicode(sys.executable, sys.getfilesystemencoding()))
            return os.path.dirname(unicode(__file__, sys.getfilesystemencoding()))

        imgname = u"%s\\img\\cav.bmp" % (img_dir())
        img = wx.Image(imgname, wx.BITMAP_TYPE_ANY)
        set_img = wx.StaticBitmap(self, -1, wx.BitmapFromImage(img))

##        txt1 = "Please direct any questions to Carmen at c_chow568@yahoo.ca"
##        t_contact = wx.StaticText(pnl, label=txt1)

        sizer = wx.BoxSizer(wx.VERTICAL)
##        sizer.Add(t_about, -1, wx.ALL|wx.LEFT, 10)
##        sizer.AddSpacer(10)
        sizer.Add(set_img, 0, wx.ALL, 10)
##        sizer.AddSpacer(10)
##        sizer.Add(t_contact, 5, wx.ALL|wx.LEFT, 10)

        pnl.SetSizer(sizer)

class HelpDlg(wx.Dialog):
    """
    Troubleshooting tips
    """
    def __init__(self, *args, **kwargs):
        wx.Dialog.__init__(self, *args, **kwargs)

        self.Center()
        self.SetSize((450,350))
        pnl = wx.lib.scrolledpanel.ScrolledPanel(self)
        pnl.SetupScrolling(scroll_x=False)

        txt0 = "Here are some potential problems you might run into " \
               "while using this program and how you can troubleshoot them."
        t_intro = wx.StaticText(pnl, label=txt0)
        t_intro.Wrap(420)

        txt1 = "To analyze one file after another, simply repeat the steps " \
               "listed on the program buttons. "

        txt2 = "If you get an error message regarding your column ID input, " \
               "check to make sure: \n" \
               "\t1. Your selected column contains the correct data \n" \
               "\t2. You have selected the correct excel worksheet \n" \
               "\t3. Your column IDs contain only letters (e.g. A, B, MM, etc)"

        t_rerun = wx.StaticText(pnl, label=txt1)
        t_rerun.Wrap(420)

        t_colerror = wx.StaticText(pnl, label=txt2)
        t_colerror.Wrap(420)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(t_intro, 0, wx.ALL, 10)
        sizer.AddSpacer(10)
        sizer.Add(t_rerun, 0, wx.ALL, 10)
        sizer.Add(t_colerror, 0, wx.ALL, 10)

        pnl.SetSizer(sizer)
