# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys, psutil
import wx, wx.lib.scrolledpanel
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

filetypes = "All Excel File (*.xls, *.xlsx)|*.xls;*.xlsx" 

class IntroPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        txt = "Make sure the data in your excel file is formatted " \
              "similarly to this: "
        txt_i = wx.StaticText(self, label=txt)
     
        pwd = os.getcwd()
        wxImg = wx.EmptyBitmap(1,1)
        imgname = u"%s\\format.bmp" % (pwd)
        img = wx.Image(imgname, wx.BITMAP_TYPE_ANY)
        w = img.GetWidth()
        h = img.GetHeight()
        set_img = wx.StaticBitmap(self, -1, wx.BitmapFromImage(img))

        sizer_i = wx.BoxSizer(wx.VERTICAL)
        sizer_i.Add(txt_i, 0, wx.TOP|wx.CENTER, 10)
        sizer_i.AddSpacer(15)
        sizer_i.Add(set_img, 0, wx.ALL|wx.CENTER, 10)
        self.SetSizer(sizer_i)
        

class XLPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        self.SetSize((200, 300))
        self.Center()

        self.curr_dir = os.getcwd()

        # EXCEL VARS
        self.orig_xls= ""
        self.wkbook = ""
        self.wksheet = 0
        self.old_sheet = ""
        self.new_book = ""
        self.new_sheet = 0
        self.wanted_rows = 0
        self.wanted_cols = 0

        # CALCULATION VARS
        self.pixels = 0     # input pixels per micron
        self.volumes = []

        b_open = wx.Button(self, label="Step 1. Select Excel file")
        b_sheet = wx.Button(self, label="Step 2. Select Excel sheet")
        t_pixels = wx.StaticText(self, label="Step 3. Pixels pers micron:")
        self.pixels = wx.TextCtrl(self)
        b_run = wx.Button(self, label="Step 4. Start Program")

        b_open.Bind(wx.EVT_BUTTON, self.open_xls)
        b_sheet.Bind(wx.EVT_BUTTON, self.get_sheet)
        b_run.Bind(wx.EVT_BUTTON, self.start)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(b_open, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(b_sheet, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(t_pixels, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(self.pixels, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(b_run, 0, wx.ALL|wx.CENTER, 15)

        self.SetSizer(sizer)

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
        print type(self.orig_xls)
        open_xls.Destroy()

    def get_sheet(self, event):
        """
        Choose excel sheet to work with
        """
        try:
            self.wkbook = open_workbook(self.orig_xls, on_demand=True)

            wksheets = []
            for s in self.wkbook.sheet_names():
                wksheets.append(s)
        except IOError:
            msg = "Please select an excel file first!" 
            io = wx.MessageDialog(None, msg, "ERROR",
                                wx.OK|wx.ICON_EXCLAMATION)
            io.ShowModal()
            io.Destroy()
        else:
            sheets = wx.SingleChoiceDialog(self, "Choose your excel worksheet",
                                           "Choose Sheet", wksheets)
            sheets.ShowModal()
            self.wksheet = sheets.GetSelection()
            sheets.Destroy()

        self.old_sheet = self.wkbook.sheet_by_index(self.wksheet)
        self.wanted_rows = self.old_sheet.nrows

        return self.wkbook, self.wksheet, self.old_sheet, self.wanted_rows

    def start(self, event):
        try:
            self.copy_xls(self.wkbook)
        except AttributeError:
            msg = "Please select an excel file, excel sheet, and enter in the pixel value!" 
            io = wx.MessageDialog(None, msg, "ERROR",
                                wx.OK|wx.ICON_EXCLAMATION)
            io.ShowModal()
            io.Destroy()
        else: 
            self.col_prompt()
            self.headings()
            self.multiply_by_10()
            self.dg_volume()
            self.cell_density()
            
            self.wkbook.release_resources()
     
    ##### PROGRAM START #####

    def copy_xls(self, wkbook):
        """
        Create a copy of the original excel file and adds a new sheet where the
        calculated data will go.
        """
        self.new_book = copy(wkbook)
        sheet = self.new_book.get_sheet(-1)
        if not sheet.name == 'CALCULATED DATA':
            self.new_sheet = self.new_book.add_sheet('CALCULATED DATA', cell_overwrite_ok = True)
        else:
            self.new_sheet = sheet
        return self.new_book, self.new_sheet

    def col_prompt(self):
        """
        Ask user what columns they want. 
        """
        choose = ColQueryDialog(None, title="Choose columns")
        choose.EnableLayoutAdaptation(True)
        choose.ShowModal()
        self.wanted_cols = choose.GetValues()
        choose.Destroy()
        
        print self.wanted_cols
        return self.wanted_cols

    def headings(self):
        """
        Sets up column headings for the new worksheet. 
        """
        headings = {'id': ["ID", "Group"],
                    'counts': ["Dorsal DG counts", "Ventral DG counts", "Dorsal hil counts", "Ventral hil counts"],
                    'areas': ["Dorsal DG vol", "Ventral DG vol", "Dorsal hil vol", "Ventral hil vol", ''],
                    'density': ["Dorsal DG density", "Ventral DG density", "Dorsal hil density", "Ventral hil density"]
                    }
        h_order = ['id', 'counts', 'areas', 'density']
        for key in h_order:
            if key == 'id':
                h = min(self.wanted_cols[key])
            elif key == 'counts':
                h = max(self.wanted_cols['id']) + 1
            elif key == 'areas':
                h = max(self.wanted_cols['counts']) + 1
            else:
                pass

            for cell in range(len(headings[key])):
                self.new_sheet.write(0, h, headings[key][cell])
                h += 1
                
        for row in range(1, self.wanted_rows):
            for i in range(2):
                self.new_sheet.write(row, i, self.old_sheet.cell(row, self.wanted_cols['id'][i]).value)
 
    def multiply_by_10(self):
        """
        Cell counts - multiply all values by 10.
        """
        for row in range(1, self.wanted_rows):
            c = max(self.wanted_cols['id']) + 1
            for col in range(len(self.wanted_cols['counts'])):
                calc = self.old_sheet.cell(row, self.wanted_cols['counts'][col]).value * 10
                self.new_sheet.write(row, c, calc)
                c += 1

    def dg_volume(self):
        """
        Calculate DG & hilar volume for each brain. 
        """
        px = (self.pixels.GetValue()).encode('utf-8')
        calc = (1/float(px))**2 * 2 * 0.4

        for row in range(1, self.wanted_rows):
            c = max(self.wanted_cols['counts']) + 1
            for col in range(len(self.wanted_cols['areas'])):
                vol = self.old_sheet.cell(row, self.wanted_cols['areas'][col]).value * calc
                self.new_sheet.write(row, c, vol)
                c += 1

        print self.wanted_cols['areas']
        self.new_book.save('temp-areas-file.xls')
        print "temp file created"

    def cell_density(self):
        """
        Calculate cell density for each brain. counts/mm3
        """
        temp_book = open_workbook('temp-areas-file.xls')
        new_book = self.copy_xls(temp_book)
        temp_sheet = temp_book.sheet_by_name('CALCULATED DATA')

        for row in range(1, self.wanted_rows):
            c =  max(self.wanted_cols['id']) + 1
            for col in range(4):
                count = temp_sheet.cell(row, c).value
                area = temp_sheet.cell(row, c + 5).value
                if not (count == "" or area == ""):
                    calc = count/area
                else:
                    calc = ""
                self.new_sheet.write(row, c + 10, calc)
                c += 1

        temp_book.release_resources()
        os.remove('temp-areas-file.xls')

        try: 
            self.new_book.save('areas-02932075348902.xls')
        except IOError:                   
            msg = "Please close the file areas-02932075348902.xls before proceeding\n" \
                  "Click OK when the file is closed"
            io = wx.MessageDialog(None, msg, "ERROR",
                                wx.OK|wx.ICON_EXCLAMATION)
            
            if io.ShowModal() == wx.ID_OK:
                try:
                    self.new_book.save('areas-02932075348902.xls')
                    print "done"
                except IOError:
                    print "please close your file!"
            io.Destroy()
        else:
            print "****"          
        
        
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

        sizer.Add(t_id, 0, wx.LEFT|wx.TOP, 5)
        sizer.Add(self.id, 0, wx.LEFT, 5)
        sizer.Add(t_group, 0, wx.LEFT, 5)
        sizer.Add(self.group, 0, wx.LEFT, 5)

        sizer.AddSpacer(20)
        sizer.Add(t_count, 0, wx.ALL, 10)
        sizer.Add(t_ddg, 0, wx.LEFT, 5)
        sizer.Add(self.ddg, 0, wx.LEFT, 5)
        sizer.Add(t_vdg, 0, wx.LEFT, 5)
        sizer.Add(self.vdg, 0, wx.LEFT, 5)
        sizer.Add(t_dhil, 0, wx.LEFT, 5)
        sizer.Add(self.dhil, 0, wx.LEFT, 5)
        sizer.Add(t_vhil, 0, wx.LEFT, 5)
        sizer.Add(self.vhil, 0, wx.LEFT, 5)

        sizer.AddSpacer(20)
        sizer.Add(t_area, 0, wx.ALL, 10)
        sizer.Add(t_ddgarea, 0, wx.LEFT, 5)
        sizer.Add(self.ddgarea, 0, wx.LEFT, 5)
        sizer.Add(t_vdgarea, 0, wx.LEFT, 5)
        sizer.Add(self.vdgarea, 0, wx.LEFT, 5)
        sizer.Add(t_dhilarea, 0, wx.LEFT, 5)
        sizer.Add(self.dhilarea, 0, wx.LEFT, 5)
        sizer.Add(t_vhilarea, 0, wx.LEFT, 5)
        sizer.Add(self.vhilarea, 0, wx.LEFT, 5)

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

        print wanted_cols
        
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

def runfile():
    app = wx.App()
    Count(None)
    app.MainLoop()

if __name__ == '__main__':
    runfile()
        
