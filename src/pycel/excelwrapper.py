try:
    import win32com.client
    #http://www.py2exe.org/index.cgi/IncludingTypelibs
    #win32com.client.gencache.is_readonly=False
    #win32com.client.gencache.GetGeneratePath()
    from win32com.client import Dispatch
    from win32com.client import constants 
    import pythoncom   
except Exception as e:
    print "WARNING: cant import dependent win32 packages:",e
try:
    import numpy as np
except Exception as e:
    print "WARNING: cant import dependent numpy packages:",e
try:
    import openpyxl
except Exception as e:
    print "WARNING: cant import dependent packages:",e

import os
from os import path

class ExcelComWrapper(object):
    
    def __init__(self, filename, app=None):
        
        super(ExcelComWrapper,self).__init__()
        
        self.filename = path.abspath(filename)
        self.app = app
        
        # WARNING: here is the openpyxl way to get ranged names, for the win32.com way see get_rangednames function below
        # TODO: automate detection of max string length to set up numpy array accordingly
        # TODO: discriminate between worksheet & workbook ranged names
        try:
            wb = openpyxl.load_workbook(self.filename)
            rn = wb.get_named_ranges()
            self.rangednames_openpyxl = np.zeros(shape = (int(len(rn)),1), dtype=[('id', 'int_'), ('name', 'S200'), ('formula', 'S200')])
            for i in range(0, len(rn)):
                self.rangednames_openpyxl[i]['id'] = int(i+1)       
                self.rangednames_openpyxl[i]['name'] = rn[i].name      
                self.rangednames_openpyxl[i]['formula'] = rn[i].value
        except (RuntimeError, TypeError, NameError):
            pass
            
    def get_rangednames(self):
    
        # WARNING: stored here the win32.com way to get ranged names, if range is with single column/line it will fail
        # WARNING: by default numpy array requires dtype declaration to specify character length (here 'S200', i.e. 200 characters)    
        # TODO: automate detection of max string length to set up numpy array accordingly
        # TODO: discriminate between worksheet & workbook ranged names

        self.rangednames = np.zeros(shape = (int(self.app.ActiveWorkbook.Names.Count),1), dtype=[('id', 'int_'), ('name', 'S200'), ('formula', 'S200')])
        for i in range(0, self.app.ActiveWorkbook.Names.Count):
            self.rangednames[i]['id'] = int(i+1)       
            self.rangednames[i]['name'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Name)        
            self.rangednames[i]['formula'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Value)
        return self.rangednames
            
    def connect(self):
        #http://devnulled.com/content/2004/01/com-objects-and-threading-in-python/
        # TODO: dont need to uninit?
        #pythoncom.CoInitialize()
        if not self.app:
            self.app = Dispatch("Excel.Application")
            self.app.Visible = True
            self.app.DisplayAlerts = 0
            self.app.Workbooks.Open(self.filename)
        else:
            # if we are running as an excel addin, this gets passed to us
            pass
    
    def save(self):
        self.app.ActiveWorkbook.Save()
    
    def save_as(self, filename, delete_existing=False):
        if delete_existing and os.path.exists(filename):
            os.remove(filename)
        self.app.ActiveWorkbook.SaveAs(filename)
  
    def close(self):
        self.app.ActiveWorkbook.Close(False)
  
    def quit(self):
        return self.app.Quit()
        
    def set_sheet(self,s):
        return self.app.ActiveWorkbook.Worksheets(s).Activate()
    
    def get_sheet(self):
        return self.app.ActiveWorkbook.ActiveSheet
            
    def get_range(self, range):
        #print '*',range
        if range.find('!') > 0:
            sheet,range = range.split('!')
            return self.app.ActiveWorkbook.Worksheets(sheet).Range(range)
        else:        
            return self.app.ActiveWorkbook.ActiveSheet.Range(range)

    def get_used_range(self):
        return self.app.ActiveWorkbook.ActiveSheet.UsedRange

    def get_active_sheet(self):
        return self.app.ActiveWorkbook.ActiveSheet.Name
    
    def get_cell(self,r,c):
        return self.app.ActiveWorkbook.ActiveSheet.Cells(r,c)
        
    def get_value(self,r,c):
        return self.get_cell(r, c).Value
    
    def set_value(self,r,c,val):
        self.get_cell(r, c).Value = val

    def get_formula(self,r,c):
        f = self.get_cell(r, c).Formula
        return f if f.startswith("=") else None 
    
    def has_formula(self,range):
        f = self.get_range(range).Formula
        return f and f.startswith("=")
    
    def get_formula_from_range(self,range):
        f = self.get_range(range).Formula
        if isinstance(f, (list,tuple)):
            if any(filter(lambda x: x[0].startswith("="),f)):
                return [x[0] for x in f];
            else:
                return None
        else:
            return f if f.startswith("=") else None 
    
    def get_formula_or_value(self,name):
        r = self.get_range(name)
        return r.Formula or r.Value

    def get_row(self,row):
        return [self.get_value(row,col+1) for col in range(self.get_used_range().Columns.Count)]

    def set_calc_mode(self,automatic=True):
        if automatic:
            self.app.Calculation = constants.xlCalculationAutomatic
        else:
            self.app.Calculation = constants.xlCalculationManual

    def set_screen_updating(self,update):
        self.app.ScreenUpdating = update

    def run_macro(self,macro):
        self.app.Run(macro)
