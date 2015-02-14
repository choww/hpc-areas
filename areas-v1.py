# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import wx
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

filetypes = "All Excel File (*.xls, *.xlsx)|*.xls;*.xlsx|" \
            "All files (*.*)|*.*"

class Count(wx.Frame):
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, **kwargs)

        self.SetSize((200, 300))
        self.Center()

        self.curr_dir = os.getcwd()

        # EXCEL VARS
        self.orig_xls= ""
        self.wkbook = ""
        self.wksheet = ""
        self.new_book = ""
        self.new_sheet = 0
        self.wanted_rows = 0
        self.wanted_cols = 0

        # CALCULATION VARS
        self.pixels = 0     # input pixels per micron

        # APP BODY

        pnl = wx.Panel(self)
        b_open = wx.Button(pnl, label="Step 1. Select Excel file")
        b_sheet = wx.Button(pnl, label="Step 2. Select Excel sheet")
        t_pixels = wx.StaticText(pnl, label="Step 4. Pixels pers micron:")
        self.pixels = wx.TextCtrl(pnl)
        b_run = wx.Button(pnl, label="Step 5. Start Program")

        b_open.Bind(wx.EVT_BUTTON, self.open_xls)
        b_sheet.Bind(wx.EVT_BUTTON, self.get_sheet)
        b_run.Bind(wx.EVT_BUTTON, self.start)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(b_open, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(b_sheet, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(t_pixels, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(self.pixels, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(b_run, 0, wx.ALL|wx.CENTER, 15)

        pnl.SetSizer(sizer)
        self.Show()
        
    def open_xls(self, event):
        """
        Opens the selected excel file 
        """
        open_xls = wx.FileDialog(self, "Open File",
                                 defaultDir=self.curr_dir,
                                 defaultFile="",
                                 wildcard=filetypes,
                                 style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST|wx.CHANGE_DIR
                                 )
        open_xls.ShowModal()
        self.orig_xls = open_xls.GetPath()
        open_xls.Destroy()

    def get_sheet(self, event):
        """
        Choose excel sheet to work with
        """
        self.wkbook = open_workbook(self.orig_xls, on_demand=True)

        wksheets = []
        for s in self.wkbook.sheet_names():
            wksheets.append(s)

        sheets = wx.SingleChoiceDialog(self, "Choose your excel worksheet",
                                       "Choose Sheet", wksheets)
        sheets.ShowModal()
        self.wksheet = sheets.GetSelection()
        sheets.Destroy()

        return self.wkbook, self.wksheet

    def copy_xls(self):
        """
        Create a copy of the original excel file. 
        """
        self.new_book = copy(self.wkbook)
        new_sheet = self.new_book.get_sheet(self.wksheet)
        new_sheet.name = 'CALCULATED DATA'
        self.new_sheet = self.wkbook.sheet_by_index(self.wksheet)
        self.wanted_rows = self.new_sheet.nrows
        return self.new_book, self.new_sheet, self.wanted_rows
        
    def start(self, event):
        self.copy_xls()
        self.col_prompt()
        self.multiply_by_10()
        self.dg_volume()
        
        self.wkbook.release_resources()
        self.new_book.save('areas-02932075348902.xls')  # TEMPORARY--will allow user to save their own sheet. 

    ##### PROGRAM START #####

    def col_prompt(self):
        """
        Ask user what columns they want. 
        """

        choose = ColQueryDialog(None, title="Choose columns")
        choose.ShowModal()
        self.wanted_cols = choose.GetValues()
        choose.Destroy()

 
    def multiply_by_10(self):
        """
        Cell counts - multiply all values by 10.
        """
        
        for row in range(1, self.wanted_rows):
            for col in range(len(self.wanted_cols['counts'])):
                self.new_sheet.write(row, self.wanted_cols['counts'][col], sheet.cell(row, self.wanted_cols['counts'][col]).value * 10)               

    def dg_volume(self):
        """
        Calculate dentate gyrus & hilar volume for each brain. 
        """
        px = (self.pixels.GetValue()).encode('utf-8')
        calc = (1/int(px))**2 * 2 * 0.4
        for row in range(1, self.wanted_rows):
            for col in range(len(self.wanted_cols['areas'])):
                self.new_sheet.write(row, self.wanted_cols['areas'][col], sheet.cell(row, self.wanted_cols['areas'][col]).value * calc)

    def cell_density(self):
        #calculate cell density for each brain. counts/mm3
        pass

class ColQueryDialog(wx.Dialog):
    """
    Let users choose which columns their data are located. 
    """
    def __init__(self, *args, **kwargs):
        wx.Dialog.__init__(self, *args, **kwargs)

        self.SetSize((300,600))
        self.Center()
        pnl = wx.Panel(self)

        instruction = "Select a column [e.g. A, B, etc] for each parameter."
        t_instruct = wx.StaticText(pnl, label=instruction)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(t_instruct, 0, wx.ALL|wx.CENTER, 5)

        t_ddg = wx.StaticText(pnl, label="Dorsal dentate")
        self.ddg = wx.TextCtrl(pnl)
        t_vdg = wx.StaticText(pnl, label="Ventral dentate")
        self.vdg = wx.TextCtrl(pnl)
        t_dhil = wx.StaticText(pnl, label="Dorsal hilus")
        self.dhil = wx.TextCtrl(pnl)
        t_vhil = wx.StaticText(pnl, label="Ventral hilus")
        self.vhil = wx.TextCtrl(pnl)

        t_ddgarea = wx.StaticText(pnl, label="Dorsal DG areas")
        self.ddgarea = wx.TextCtrl(pnl)
        t_vdgarea = wx.StaticText(pnl, label="Ventral DG areas")
        self.vdgarea = wx.TextCtrl(pnl)
        t_dhilarea = wx.StaticText(pnl, label="Dorsal hilus areas")
        self.dhilarea = wx.TextCtrl(pnl)
        t_vhilarea = wx.StaticText(pnl, label="Ventral hilus areas")
        self.vhilarea = wx.TextCtrl(pnl)

        sizer.Add(t_ddg)
        sizer.Add(self.ddg)
        sizer.Add(t_vdg)
        sizer.Add(self.vdg)
        sizer.Add(t_dhil)
        sizer.Add(self.dhil)
        sizer.Add(t_vhil)
        sizer.Add(self.vhil)

        sizer.Add(t_ddgarea)
        sizer.Add(self.ddgarea)
        sizer.Add(t_vdgarea)
        sizer.Add(self.vdgarea)
        sizer.Add(t_dhilarea)
        sizer.Add(self.dhilarea)
        sizer.Add(t_vhilarea)
        sizer.Add(self.vhilarea)

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

        wanted_cols['counts'] = []
        wanted_cols['counts'].append(self.ddg.GetValue())
        wanted_cols['counts'].append(self.vdg.GetValue())
        wanted_cols['counts'].append(self.dhil.GetValue())
        wanted_cols['counts'].append(self.vhil.GetValue())

        wanted_cols['areas'] = []
        wanted_cols['areas'].append(self.ddgarea.GetValue())
        wanted_cols['areas'].append(self.vdgarea.GetValue())
        wanted_cols['areas'].append(self.vhilarea.GetValue())
        wanted_cols['areas'].append(self.vhilarea.GetValue())
        
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for key in wanted_cols: 
            for c in range(len(wanted_cols[key])):
                wanted_cols[key][c] = wanted_cols[key][c].encode('utf-8').upper()
                if len(wanted_cols[key][c]) == 1 or not wanted_cols[key][c] == None:
                    wanted_cols[key][c] = letters.find(wanted_cols[key][c])
                else:
                    a = 26
                    b = 0
                    for l in range(1,len(wanted_cols[key][c])):
                        b += letters.find(wanted_cols[key][c][1])
                    wanted_cols[key][c] = (letters.find(wanted_cols[key][c][0]) + 1) * a + b
        print wanted_cols
        return wanted_cols

def runfile():
    app = wx.App()
    Count(None)
    app.MainLoop()

if __name__ == '__main__':
    runfile()
        
