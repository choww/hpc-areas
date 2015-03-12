# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import wx
import col_dlg
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
        
        self.curr_dir = os.getcwd()

        # EXCEL VARS
        self.orig_xls= ""
        self.wkbook = ""
        self.wksheet = 0
        self.old_sheet = ""
        self.new_book = ""
        self.new_sheet = 0
        self.wanted_rows = 0
        self.wanted_cols = []
        self.new_cols = []

        # CALCULATION VARS
        self.pixels = 0     # input pixels per micron
        self.ror = ""

        b_open = wx.Button(self, label="Step 1. Select Excel file")
        self.t_open = wx.StaticText(self, label="", size=(220,-1))
        b_sheet = wx.Button(self, label="Step 2. Select Excel sheet")
        self.t_sheet = wx.StaticText(self, label="", size=(100,-1))
        t_pixels = wx.StaticText(self, label="Step 3. Pixels pers micron:")
        self.pixels = wx.TextCtrl(self)

        t_region = wx.StaticText(self, label="Step 4. Choose method of analysis")
        self.r_dv = wx.RadioButton(self, label="Dorsal/Ventral")
        self.r_tot = wx.RadioButton(self, label="Combined")
        self.r_all = wx.RadioButton(self, label="All of the above")

        self.r_dv.Bind(wx.EVT_RADIOBUTTON, self.setROR)
        self.r_tot.Bind(wx.EVT_RADIOBUTTON, self.setROR)
        self.r_all.Bind(wx.EVT_RADIOBUTTON, self.setROR)
        
        b_run = wx.Button(self, label="Step 5. Start Program")

        self.t_cols = wx.StaticText(self, label="", size=(220,40))
        
        b_open.Bind(wx.EVT_BUTTON, self.open_xls)
        b_sheet.Bind(wx.EVT_BUTTON, self.get_sheet)
        b_run.Bind(wx.EVT_BUTTON, self.start)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.AddMany([(b_open, 0, wx.TOP|wx.CENTER, 15),
                       (self.t_open, 0, wx.ALL|wx.CENTER, 5),
                       (b_sheet, 0, wx.ALL|wx.CENTER, 5),
                       (self.t_sheet, 0, wx.ALL|wx.CENTER, 5),
                       (t_pixels, 0, wx.ALL|wx.CENTER, 5),
                       (self.pixels, 0, wx.ALL|wx.CENTER, 5),
                       (t_region, 0, wx.ALL|wx.CENTER, 5),
                       (self.r_dv, 0, wx.ALL|wx.CENTER, 5),
                       (self.r_tot, 0, wx.ALL|wx.CENTER, 5),
                       (self.r_all, 0, wx.ALL|wx.CENTER, 5),
                       (b_run, 0, wx.ALL|wx.CENTER, 15),
                       (self.t_cols, 1, wx.RIGHT|wx.LEFT, 10)])

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
        self.t_open.SetForegroundColour((143,28,147))
        
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
            self.t_sheet.SetForegroundColour((143,28,147))
            sheets.Destroy()

        self.old_sheet = self.wkbook.sheet_by_index(self.wksheet)
        self.wanted_rows = self.old_sheet.nrows

        return self.wkbook, self.wksheet, self.old_sheet, self.wanted_rows

    def start(self, event):
        try:
            self.new_xls()
        except AttributeError:
            msg = "Please select an excel file, excel sheet, and enter in the pixel value!" 
            io = wx.MessageDialog(None, msg, "ERROR",
                                wx.OK|wx.ICON_EXCLAMATION)
            io.ShowModal()
            io.Destroy()
        else:
            msg = ""
            px = self.pixels.GetValue()
            if px == u"":
                msg = "Please fill in the pixel value!"
            elif not px.isdigit():
                msg = "Please enter numbers only as the pixel value"
            else:
                self.col_prompt()
                try: 
                    self.multiply_by_10()
                    self.dg_volume()
                    self.cell_density()
                    self.headings()
                except IndexError:
                    msg = "Please make sure your column IDs are correct!"
                else: 
                    self.save_xls()

            ie = wx.MessageDialog(None, msg, "ERROR",
                wx.OK|wx.ICON_EXCLAMATION)
            ie.ShowModal()
            ie.Destroy()

            self.wkbook.release_resources()
            
    ##### PROGRAM START #####

    def new_xls(self):
        """
        Create the new workbook that the calculated data will be entered in.
        """
        self.new_book = Workbook()
        self.new_sheet = self.new_book.add_sheet('CALCULATED DATA', cell_overwrite_ok = True)
        return self.new_book, self.new_sheet

    def setROR(self, event):
        """
        Ask user which regions of the hippocampus they would like to analyze.
        """
        click = event.GetEventObject()
        self.ror = click.GetLabel()
        return self.ror

    def col_prompt(self):
        """
        Ask user what columns they want. 
        """
        while True: 
            try:
                choose = ""
                ror = ""
                if self.ror == u'Dorsal/Ventral':
                    choose = col_dlg.DVDialog(None, title="Choose Columns")
                    ror = 'dv'
                elif self.ror == u'Combined':
                    choose = col_dlg.TotDialog(None, title="Choose Columns")
                    ror = 'tot'
                else:
                    choose = col_dlg.AllDialog(None, title="Choose Columns")
                    ror = 'all'

                choose.EnableLayoutAdaptation(True)
                if choose.ShowModal() == wx.ID_OK: 
                    choose.Destroy()
                    colnames = choose.GetValues(ror)[0]
                
                    # Update self.t_cols in main panel
                    txt = "Selected Columns: \n" \
                          "%s: %s; %s: %s; %s: %s"
                    keys = ['id', 'counts', 'areas']
                    t_cols = txt % (keys[0], ", ".join(colnames[keys[0]]), \
                                    keys[1], ", ".join(colnames[keys[1]]), \
                                    keys[2], ", ".join(colnames[keys[2]]))
                    self.t_cols.SetLabel(t_cols)
                    self.t_cols.SetForegroundColour((143,28,147))
                    self.wanted_cols = choose.GetIndices()
                    self.new_cols = choose.GetIndices()
                    print "col_prompt", self.wanted_cols
                break
            except IndexError:
                msg = "Please make sure all fields are filled in!"
                ie = wx.MessageDialog(None, msg, "ERROR",
                    wx.OK|wx.ICON_EXCLAMATION)
                ie.ShowModal()
                ie.Destroy()
            else:
                return self.wanted_cols
                break

    def headings(self):
        """
        Sets up column headings for the new worksheet. 
        """                  
        headings = {}
        headings['id'] = []
        headings['counts'] = []
        headings['areas'] = []
        headings['density'] = {'region': ["Dorsal", "Ventral"],
                               'hpc': ["DG density", "Hil density"]}
        h_order = ['id', 'counts', 'areas', 'density']

        for key in self.wanted_cols:
            for col in range(len(self.wanted_cols[key])):
                headings[key].append(self.old_sheet.cell(0, self.wanted_cols[key][col]).value)
            print headings[key]
        
        for key in h_order:
            if key == 'id':
                h = min(self.new_cols[key])
            elif key == 'counts':
                h = max(self.new_cols['id']) + 1
            elif key == 'areas':
                h = max(self.new_cols['counts']) + 2
            else:
                h = max(self.new_cols['areas']) + 2

            for cell in range(len(headings[key])):
                hdg = headings[key][cell]
                if key == 'density':
                    if self.ror == u'Dorsal/Ventral':
                        hdg = headings[key]['region'][cell] + headings[key]['hpc'][cell]
                    elif self.ror == u'Combined':
                        hdg = headings[key]['hpc'][cell]
                    else:
                        hdg = headings[key]['region'][cell] + headings[key]['hpc'][cell]

                self.new_sheet.write(0, h, hdg)
                h += 1
                
        for row in range(1, self.wanted_rows):
            for i in range(2):
                self.new_sheet.write(row, i, self.old_sheet.cell(row, self.wanted_cols['id'][i]).value)

    ### REFACTOR ATTEMPT 1 ###
    def enter_data(self, x, y, z):
        """
        x = key for self.wanted_cols
        y = how much to add to c
        z = key for the self.wanted_cols inside the nested for loop
        """
        for row in range(1, self.wanted_rows):
            c = max(self.new_cols[x]) + y
            for col in range(len(self.wanted_cols[z])):
                calc = ""
                self.new_sheet.write(row, c, calc)
                self.new_cols[z][col] = c
                c += 1
 
    def multiply_by_10(self):
        
        """
        Cell counts - multiply all values by 10.
        """
        for row in range(1, self.wanted_rows):
            c = max(self.wanted_cols['id']) + 1
            for col in range(len(self.wanted_cols['counts'])):
                calc = self.old_sheet.cell(row, self.wanted_cols['counts'][col]).value * 10
                self.new_sheet.write(row, c, calc)
                self.new_cols['counts'][col] = c
                c += 1
        print "counts:", self.new_cols

    def dg_volume(self):
        """
        Calculate DG & hilar volume for each brain. 
        """
        px = (self.pixels.GetValue()).encode('utf-8')
        calc = (1/float(px))**2 * 2 * 0.4
        print type(calc)
        
        for row in range(1, self.wanted_rows):
            c = max(self.new_cols['counts']) + 2
            for col in range(len(self.wanted_cols['areas'])):
                area = self.old_sheet.cell(row, self.wanted_cols['areas'][col]).value
                if  area == "":
                    pass
                else: 
                    vol = area * calc
                    self.new_sheet.write(row, c, vol)
                    self.new_cols['areas'][col] = c
                    c += 1
        print "vol:", self.new_cols

        self.new_book.save('temp-areas-file.xls')
        print "temp file created"

    def cell_density(self):
        """
        Calculate cell density for each brain. counts/mm3
        """
        temp_book = open_workbook('temp-areas-file.xls')
        self.new_book = copy(temp_book)
        self.new_sheet = self.new_book.get_sheet(-1)
        temp_sheet = temp_book.sheet_by_name('CALCULATED DATA')

        cc = self.new_cols['counts']
        ca =  self.new_cols['areas']
        self.new_cols['density'] = []
        for i in range(len(ca)):
            self.new_cols['density'].append(i)

        for row in range(1, self.wanted_rows):
            c = max(ca)+2
            for col in range(len(ca)):
                count = temp_sheet.cell(row, cc[col]).value
                area = temp_sheet.cell(row, ca[col]).value
                
                if not (count == "" or area == ""):
                    calc = count/area
                else:
                    calc = ""
                self.new_sheet.write(row, c, calc)
                self.new_cols['density'][col] = c
                c += 1
        print self.new_cols['density']
        temp_book.release_resources()
        os.remove('temp-areas-file.xls')

    def save_xls(self):
        while True: 
            try:
                save = wx.FileDialog(self, message="Save file as...",
                                     defaultDir=self.curr_dir,
                                     defaultFile="",
                                     wildcard=filetypes,
                                     style=wx.SAVE)
                if save.ShowModal() == wx.ID_OK:
                    new_file = save.GetPath()
                    self.new_book.save(new_file)
                    save.Destroy()

                    wx.MessageBox('Saved: %s' % (new_file),
                                  'Finished', wx.OK|wx.ICON_INFORMATION)
                break
            except IOError:                    
                msg = "Please close the file %s before proceeding" % (new_file) 
                io = wx.MessageDialog(None, msg, "ERROR",
                                    wx.OK|wx.ICON_EXCLAMATION)
                io.ShowModal()
                io.Destroy()
            else:
                print "****"
                break
