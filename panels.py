# !/usr/bin/python
# -*- coding: utf-8 -*-
import os, wx, col_query
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

filetypes = "All Excel File (*.xls, *.xlsx)|*.xls;*.xlsx" 
class IntroPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        
        txt0 = "*** GETTING STARTED ***\n"\
               "Step 1. Sum up the cell counts and ImageJ area data for each animal."
        txt0_i = wx.StaticText(self, label=txt0)

        txt1 = "Step 2. Make sure the data in your excel file is formatted " \
              "similarly to this: "
        txt1_i = wx.StaticText(self, label=txt1)

        txt2 = "* NOTE: If you have blank spaces in your file (as shown above), "\
               "volumes and cell density won't be calculated for those animals."
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
        self.wanted_cols = {}
        self.new_cols = {}

        # CALCULATION VARS
        self.pixels = 0     # input pixels per micron
        self.volumes = []

        b_open = wx.Button(self, label="Step 1. Select Excel file")
        self.t_open = wx.StaticText(self, label="", size=(220,-1))
        b_sheet = wx.Button(self, label="Step 2. Select Excel sheet")
        self.t_sheet = wx.StaticText(self, label="", size=(100,-1))
        t_pixels = wx.StaticText(self, label="Step 3. Pixels pers micron:")
        self.pixels = wx.TextCtrl(self)
        b_run = wx.Button(self, label="Step 4. Start Program")

        self.t_cols = wx.StaticText(self, label="", size=(220,50))
        
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

    def check_input(self, values, obj_type):
        """
        Validates user input
        """
        if obj_type == 'dict':
            for key in values:
                for item in range(len(values[key])):
                    if not values[key][item].isalpha():
                        print "%s is not valid" %(values[key][item])
                        return False
            return True
                        
        elif obj_type == 'int':
            if not values.isdigit():
                return False
            elif values == u"":
                return False
            else:
                return True

    def start(self, event):
        try:
            print self.old_sheet.cell(1,1).value
        except AttributeError:
            msg = "Please select an excel file and excel sheet!" 
            io = wx.MessageDialog(None, msg, "ERROR",
                                wx.OK|wx.ICON_EXCLAMATION)
            io.ShowModal()
            io.Destroy()
        else: 
            px = self.pixels.GetValue()
            if not self.check_input(px, 'int'):
                msg = "Please make sure your pixel value is correct!"
            else:
                self.new_xls()
                self.col_prompt()
                try: 
                    self.multiply_by_10()
                    self.dg_volume()
                    self.cell_density()
                    self.headings()
                except IndexError:
                    msg = "Please make sure your column IDs are correct!"
                except TypeError:
                    # TypeError: can't multiply sequence by non-int of type 'float'
                    msg = "Please make sure your selected columns contain only Numbers!"
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
            
    def col_prompt(self):
        """
        Ask user what columns they want. 
        """
        while True:
            try:
                choose = col_query.ColQueryDialog(None, title="Choose columns")
                choose.EnableLayoutAdaptation(True)
                if choose.ShowModal() == wx.ID_OK:
                    colnames = choose.GetValues()
                
                    # Updates self.t_cols 
                    txt = "Selected Columns: \n" \
                          "%s: %s; %s: %s; \n" \
                          "%s: %s"
                    keys = ['id', 'counts', 'areas']
                    t_cols = txt % (keys[0], ", ".join(colnames[keys[0]]), \
                                    keys[1], ", ".join(colnames[keys[1]]), \
                                    keys[2], ", ".join(colnames[keys[2]]))
                    self.t_cols.SetLabel(t_cols)
                    self.t_cols.SetForegroundColour((143,28,147))
                    if self.check_input(colnames, 'dict'):
                        self.wanted_cols = choose.GetIndices()
                        self.new_cols = choose.GetIndices()
                        break
                    else:
                        msg = "Please make sure column IDs contain letters only!"
                        te = wx.MessageDialog(None, msg, "ERROR",
                            wx.OK|wx.ICON_EXCLAMATION)
                        te.ShowModal()
                        te.Destroy()
            except IndexError:
                msg = "Please make sure all fields are filled in!"
                ie = wx.MessageDialog(None, msg, "ERROR",
                    wx.OK|wx.ICON_EXCLAMATION)
                ie.ShowModal()
                ie.Destroy()
            else:
                choose.Destroy()
                return self.wanted_cols, self.new_cols
                break

    def headings(self):
        """
        Sets up column headings for the new worksheet. 
        """
        headings = {'id': ["ID", "Group"],
                    'counts': ["Dorsal DG counts", "Ventral DG counts", "Dorsal hil counts",
                               "Ventral hil counts", "Total DG counts", "Total hil counts"],
                    'areas': ["Dorsal DG vol", "Ventral DG vol", "Dorsal hil vol",
                              "Ventral hil vol", "Total DG vol", "Total hil vol"],
                    'density': ["Dorsal DG density", "Ventral DG density", "Dorsal hil density",
                                "Ventral hil density", "Total DG density", "Total hil density"]
                    }
        h_order = ['id', 'counts', 'areas', 'density']
        for key in h_order:
            if key == 'id':
                h = min(self.new_cols[key])
            elif key == 'counts':
                h = max(self.new_cols['id']) + 1
            elif key == 'areas':
                h = max(self.new_cols['counts']) + 2
            elif key == 'density':
                h = max(self.new_cols['areas']) + 2
                print "headings", self.new_cols
            else:
                pass

            for label in range(len(headings[key])):
                self.new_sheet.write(0, h, headings[key][label])
                h += 1
                
        for row in range(1, self.wanted_rows):
            for i in range(2):
                self.new_sheet.write(row, i,
                                     self.old_sheet.cell(row, self.wanted_cols['id'][i]).value)
 
    def multiply_by_10(self):
        """
        Cell counts - multiply all values by 10.
        """
        for row in range(1, self.wanted_rows):
            c = max(self.wanted_cols['id']) + 1
            for col in range(len(self.wanted_cols['counts'])):
                count = self.old_sheet.cell(row, self.wanted_cols['counts'][col]).value
                if count == "":
                    pass
                else:
                    calc = count * 10
                    self.new_sheet.write(row, c, calc)

                self.new_cols['counts'][col] = c
                c += 1

    def dg_volume(self):
        """
        Calculate DG & hilar volume for each brain. 
        """
        px = (self.pixels.GetValue()).encode('utf-8')
        calc = (1/float(px))**2 * 2 * 0.4
        
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
                else:
                    wx.MessageBox('File not saved.',
                                  'Cancelled', wx.OK|wx.ICON_INFORMATION)
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
