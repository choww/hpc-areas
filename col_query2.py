# !/usr/bin/python
# -*- coding: utf-8 -*-
import wx, wx.lib.scrolledpanel

class ColQueryDialog(wx.Dialog):
    def __init__(self, *args, **kwargs):
        wx.Dialog.__init__(self, *args, **kwargs)

        self.ror = ""

    def GetValues(self, ror):
        wanted_cols = {}

        wanted_cols['id'] = []
        wanted_cols['id'].append(self.id.GetValue())
        wanted_cols['id'].append(self.group.GetValue())

        wanted_cols['counts'] = []
        wanted_cols['areas'] = []

        if ror == 'dv':
            wanted_cols['counts'].append(self.ddg.GetValue())
            wanted_cols['counts'].append(self.vdg.GetValue())
            wanted_cols['counts'].append(self.dhil.GetValue())
            wanted_cols['counts'].append(self.vhil.GetValue())

            wanted_cols['areas'].append(self.ddgarea.GetValue())
            wanted_cols['areas'].append(self.vdgarea.GetValue())
            wanted_cols['areas'].append(self.dhilarea.GetValue())
            wanted_cols['areas'].append(self.vhilarea.GetValue())
        elif ror == 'tot':
            wanted_cols['counts'].append(self.dg.GetValue())
            wanted_cols['counts'].append(self.hil.GetValue())

            wanted_cols['areas'].append(self.dgarea.GetValue())
            wanted_cols['areas'].append(self.hilarea.GetValue())
        else:
            wanted_cols['counts'].append(self.ddg.GetValue())
            wanted_cols['counts'].append(self.vdg.GetValue())
            wanted_cols['counts'].append(self.dhil.GetValue())
            wanted_cols['counts'].append(self.vhil.GetValue())
            wanted_cols['counts'].append(self.dg.GetValue())
            wanted_cols['counts'].append(self.hil.GetValue())

            wanted_cols['areas'].append(self.ddgarea.GetValue())
            wanted_cols['areas'].append(self.vdgarea.GetValue())
            wanted_cols['areas'].append(self.dhilarea.GetValue())
            wanted_cols['areas'].append(self.vhilarea.GetValue())
            wanted_cols['areas'].append(self.dgarea.GetValue())
            wanted_cols['areas'].append(self.hilarea.GetValue())

        self.ror = ror

        for k in wanted_cols:
            for i in range(len(wanted_cols[k])):
                wanted_cols[k][i] = wanted_cols[k][i].encode('utf-8').upper()
        print wanted_cols
        return wanted_cols, self.ror

    def GetIndices(self):
        wanted_cols = self.GetValues(self.ror)[0]
        
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for key in wanted_cols:
            for c in range(len(wanted_cols[key])):
                wanted_cols[key][c] = wanted_cols[key][c].encode('utf-8').upper()
                if len(wanted_cols[key][c]) == 1:
                    wanted_cols[key][c] = letters.find(wanted_cols[key][c])
                else:
                    a = 26
                    b = 0
                    for l in range(1,len(wanted_cols[key][c])):
                        b += letters.find(wanted_cols[key][c][1])
                    wanted_cols[key][c] = (letters.find(wanted_cols[key][c][0]) + 1) * a + b
        print wanted_cols
        return wanted_cols
