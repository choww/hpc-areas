# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import wx
from col_query import *
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

filetypes = "All Excel File (*.xls, *.xlsx)|*.xls;*.xlsx" 
class IntroPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        txt0 = "This app will:\n" \
               "Calculate counts for the doral & ventral dentate gyrus and hilus (multiply values by 10)\n" \
               "Calculate dorsal & ventral dentate gyrus and hilar volumes (using Cavalieri's principle)\n" \
               "Calculate cell density based on counts and volume data (counts/mm3)"

        txt0_i = wx.StaticText(self, label=txt0)

        txt1 = "Make sure the data in your excel file is formatted " \
              "similarly to this: "
        txt1_i = wx.StaticText(self, label=txt1)

        txt2 = "* NOTE: If you have blank spaces in your file (as shown above), "\
               "areas won't be calculated for those animals."
        txt2_i = wx.StaticText(self, label=txt2)
     
        pwd = os.getcwd()
        imgname = u"%s\\img\\format.bmp" % (pwd)
        img = wx.Image(imgname, wx.BITMAP_TYPE_ANY)
        set_img = wx.StaticBitmap(self, -1, wx.BitmapFromImage(img))

        sizer_i = wx.BoxSizer(wx.VERTICAL)
        sizer_i.Add(txt0_i, 0, wx.ALL|wx.CENTER, 10)
        sizer_i.Add(txt1_i, 0, wx.TOP|wx.CENTER, 10)
        sizer_i.AddSpacer(15)
        sizer_i.Add(set_img, 0, wx.ALL|wx.CENTER, 10)
        sizer_i.Add(txt2_i, 0, wx.ALL|wx.CENTER, 10)
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
        self.t_open = wx.StaticText(self, label="", size=(220,5))
        b_sheet = wx.Button(self, label="Step 2. Select Excel sheet")
        self.t_sheet = wx.StaticText(self, label="", size=(100,5))
        t_pixels = wx.StaticText(self, label="Step 3. Pixels pers micron:")
        self.pixels = wx.TextCtrl(self)
        b_run = wx.Button(self, label="Step 4. Start Program")

        b_open.Bind(wx.EVT_BUTTON, self.open_xls)
        b_sheet.Bind(wx.EVT_BUTTON, self.get_sheet)
        b_run.Bind(wx.EVT_BUTTON, self.start)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(b_open, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(self.t_open, 5, wx.ALL|wx.CENTER, 10)
        sizer.Add(b_sheet, 0, wx.ALL|wx.CENTER, 5)
        sizer.Add(self.t_sheet, 0, wx.ALL|wx.CENTER, 10)
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
        self.t_open.SetLabel(self.orig_xls)
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
            self.t_sheet.SetLabel(wksheets[self.wksheet])
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
            px = self.pixels.GetValue()
            if px == u"":
                msg = "Please fill in the pixel value!"
            elif not px.isdigit():
                msg = "Please enter numbers only as the pixel value"
            else:
                self.col_prompt()
                self.headings()
                self.multiply_by_10()
                self.dg_volume()
                self.cell_density()
                
            io = wx.MessageDialog(None, msg, "ERROR",
                            wx.OK|wx.ICON_EXCLAMATION)
            io.ShowModal()
            io.Destroy()
            
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

        closed_file = True
        while closed_file: 
            try: 
                self.new_book.save('areas-02932075348902.xls')
                closed_file = False
            except IOError:                    
                msg = "Please close the file areas-02932075348902.xls before proceeding" 
                io = wx.MessageDialog(None, msg, "ERROR",
                                    wx.OK|wx.ICON_EXCLAMATION)
                io.ShowModal()
                io.Destroy()
            else:
                print "****"          
