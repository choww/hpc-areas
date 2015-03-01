# !/usr/bin/python
# -*- coding: utf-8 -*-
import wx, wx.lib.scrolledpanel

class ColQueryDialog(wx.Dialog):
    """
    Let users choose which columns their data are located. 
    """
    def __init__(self, *args, **kwargs):
        wx.Dialog.__init__(self, *args, **kwargs)

        self.SetSize((300,500))
        self.Center()
        pnl = wx.lib.scrolledpanel.ScrolledPanel(self)

        pnl.SetupScrolling(scroll_x=False)

        instruction = "Select a column [e.g. A, B, etc] for each parameter.\n" \
                      "Do not leave anything blank!" 
        t_instruct = wx.StaticText(pnl, label=instruction)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(t_instruct, 0, wx.ALL|wx.CENTER, 10)

        t_id = wx.StaticText(pnl, label="Animal ID")
        self.id = wx.TextCtrl(pnl)
        t_group = wx.StaticText(pnl, label="Treatment group")
        self.group = wx.TextCtrl(pnl)
        
        t_count = wx.StaticText(pnl, label="COUNTS")
        t_ddg = wx.StaticText(pnl, label="Dorsal dentate")
        self.ddg = wx.TextCtrl(pnl)
        t_vdg = wx.StaticText(pnl, label="Ventral dentate")
        self.vdg = wx.TextCtrl(pnl)
        t_dhil = wx.StaticText(pnl, label="Dorsal hilus")
        self.dhil = wx.TextCtrl(pnl)
        t_vhil = wx.StaticText(pnl, label="Ventral hilus")
        self.vhil = wx.TextCtrl(pnl)

        t_area = wx.StaticText(pnl, label="AREAS")
        t_ddgarea = wx.StaticText(pnl, label="Dorsal dentate areas")
        self.ddgarea = wx.TextCtrl(pnl)
        t_vdgarea = wx.StaticText(pnl, label="Ventral dentate areas")
        self.vdgarea = wx.TextCtrl(pnl)
        t_dhilarea = wx.StaticText(pnl, label="Dorsal hilus areas")
        self.dhilarea = wx.TextCtrl(pnl)
        t_vhilarea = wx.StaticText(pnl, label="Ventral hilus areas")
        self.vhilarea = wx.TextCtrl(pnl)

        sizer.AddMany([(t_id,0, wx.LEFT|wx.TOP, 5),
                       (self.id, 0, wx.LEFT, 5),
                       (t_group, 0, wx.LEFT, 5),
                       (self.group, 0, wx.LEFT, 5)])

        sizer.AddSpacer(20)
        sizer.AddMany([(t_count, 0, wx.ALL, 10),
                       (t_ddg, 0, wx.LEFT, 5),
                       (self.ddg, 0, wx.LEFT, 5),
                       (t_vdg, 0, wx.LEFT, 5),
                       (self.vdg, 0, wx.LEFT, 5),
                       (t_dhil, 0, wx.LEFT, 5),
                       (self.dhil, 0, wx.LEFT, 5),
                       (t_vhil, 0, wx.LEFT, 5),
                       (self.vhil, 0, wx.LEFT, 5)])

        sizer.AddSpacer(20)
        sizer.AddMany([(t_area, 0, wx.ALL, 10),
                       (t_ddgarea, 0, wx.LEFT, 5),
                       (self.ddgarea, 0, wx.LEFT, 5),
                       (t_vdgarea, 0, wx.LEFT, 5),
                       (self.vdgarea, 0, wx.LEFT, 5),
                       (t_dhilarea, 0, wx.LEFT, 5),
                       (self.dhilarea, 0, wx.LEFT, 5),
                       (t_vhilarea, 0, wx.LEFT, 5),
                       (self.vhilarea, 0, wx.LEFT, 5)])

        b_submit = wx.Button(pnl, wx.ID_OK, label="Submit")
        b_submit.SetDefault()
        b_cancel = wx.Button(pnl, wx.ID_CANCEL, label="Cancel")

        btnsizer = wx.StdDialogButtonSizer()
        btnsizer.AddButton(b_submit)
        btnsizer.AddButton(b_cancel)
        btnsizer.Realize()

        sizer.Add(btnsizer, 0, wx.ALL|wx.CENTER, 5)
        pnl.SetSizer(sizer)
        pnl.Layout()


    def GetValues(self):
        wanted_cols = {}

        wanted_cols['id'] = []
        wanted_cols['id'].append(self.id.GetValue())
        wanted_cols['id'].append(self.group.GetValue())
        
        wanted_cols['counts'] = []
        wanted_cols['counts'].append(self.ddg.GetValue())
        wanted_cols['counts'].append(self.vdg.GetValue())
        wanted_cols['counts'].append(self.dhil.GetValue())
        wanted_cols['counts'].append(self.vhil.GetValue())

        wanted_cols['areas'] = []
        wanted_cols['areas'].append(self.ddgarea.GetValue())
        wanted_cols['areas'].append(self.vdgarea.GetValue())
        wanted_cols['areas'].append(self.dhilarea.GetValue())
        wanted_cols['areas'].append(self.vhilarea.GetValue())

        for k in wanted_cols:
            for i in range(len(wanted_cols[k])):
                wanted_cols[k][i] = wanted_cols[k][i].encode('utf-8').upper()
        return wanted_cols

    def GetIndices(self):
        wanted_cols = self.GetValues()
        
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for key in wanted_cols:
            for c in range(len(wanted_cols[key])):
                wanted_cols[key][c] = wanted_cols[key][c].encode('utf-8').upper()
                if wanted_cols[key][c] == "":
                    del wanted_cols[key][c]
                elif len(wanted_cols[key][c]) == 1:
                    wanted_cols[key][c] = letters.find(wanted_cols[key][c])
                else:
                    a = 26
                    b = 0
                    for l in range(1,len(wanted_cols[key][c])):
                        b += letters.find(wanted_cols[key][c][1])
                    wanted_cols[key][c] = (letters.find(wanted_cols[key][c][0]) + 1) * a + b
        return wanted_cols
