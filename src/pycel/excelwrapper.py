from openpyxl import load_workbook
from openpyxl.cell import Cell

import os
from os import path

import abc
from abc import abstractproperty, abstractmethod

class ExcelWrapper(object):
    __metaclass__ = abc.ABCMeta
    
    @abstractproperty
    def rangednames(self):
        """Array of name_range with format { 'id':..., 'name':..., 'formula':... }"""
        return
    
    @abstractmethod
    def connect(self):
        """"""
        return

    @abstractmethod
    def save(self):
        """"""
        return
    
    @abstractmethod
    def save_as(self, filename, delete_existing=False):
        """"""
        return

    @abstractmethod
    def close(self):
        """"""
        return
  
    @abstractmethod
    def quit(self):
        """"""
        return
        
    @abstractmethod
    def set_sheet(self,s):
        """"""
        return
    
    @abstractmethod
    def get_sheet(self):
        """"""
        return
            
    @abstractmethod
    def get_range(self, range):
        """"""
        return

    @abstractmethod
    def get_used_range(self):
        """"""
        return

    @abstractmethod
    def get_active_sheet(self):
        """"""
        return
    
    @abstractmethod
    def get_cell(self,r,c):
        """"""
        return
        
    @abstractmethod
    def get_value(self,r,c):
        """"""
        return
    
    @abstractmethod
    def set_value(self,r,c,val):
        """"""
        return

    @abstractmethod
    def get_formula(self,r,c):
        """"""
        return
    
    @abstractmethod
    def has_formula(self,range):
        """"""
        return   

    @abstractmethod
    def get_formula_from_range(self,range):
        """"""
        return    
    
    @abstractmethod
    def get_formula_or_value(self,range):
        """"""
        return    

    @abstractmethod
    def get_row(self,row):
        """"""
        return

    @abstractmethod
    def set_calc_mode(self,automatic=True):
        """"""
        return

    @abstractmethod
    def set_screen_updating(self,update):
        """"""
        return

    @abstractmethod
    def run_macro(self,macro):
        """"""
        return

# Excel COM wrapper implementation for ExcelWrapper interface
class ExcelComWrapper(ExcelWrapper):
    
    def __init__(self, filename, app=None):
        
        super(ExcelWrapper,self).__init__()
        
        self.filename = path.abspath(filename)
        self.app = app
        
        # WARNING: by default numpy array require dtype declaration to specify character length (here 'S200', i.e. 200 characters)
        # WARNING: win32.com cannot get ranges with single column/line, would require way to read Office Open XML
        # TODO: automate detection of max string length to set up numpy array accordingly
        # TODO: discriminate between worksheet & workbook ranged names
        
        self._rangednames = np.zeros(shape = (int(self.app.ActiveWorkbook.Names.Count),1), dtype=[('id', 'int_'), ('name', 'S200'), ('formula', 'S200')])
        for i in range(0, self.app.ActiveWorkbook.Names.Count):
            self._rangednames[i]['id'] = int(i+1)       
            self._rangednames[i]['name'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Name)        
            self._rangednames[i]['formula'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Value)
    
    @property
    def rangednames(self):
        return self._rangednames
    

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

# OpenPyXl implementation for ExcelWrapper interface
class ExcelOpxWrapper(ExcelWrapper):
    
    def __init__(self, filename):
        
        super(ExcelWrapper,self).__init__()
        
        self.filename = path.abspath(filename)

    @property
    def rangednames(self):
        if self.workbook == None:
            return None

        rangednames = []

        for named_range in self.workbook.get_named_ranges():
            for worksheet, range_alias in named_range.destinations:
                tuple_name = {
                    'id': len(rangednames)+1,
                    'name': str(named_range.name),
                    'formula': str(worksheet.title+'!'+range_alias)
                }
                rangednames.append(tuple_name)
            
        return rangednames
    

    def connect(self):
        self.workbook = load_workbook(self.filename)

    def save(self):
        self.workbook.save(self.filename)
    
    def save_as(self, filename, delete_existing=False):
        if delete_existing and os.path.exists(filename):
            os.remove(filename)
        self.workbook.save(filename)

    def close(self):
        return
  
    def quit(self):
        return
        
    def set_sheet(self,s):
        self.workbook.active = s;
        return self.workbook.active;
    
    def get_sheet(self):
        return self.workbook.active
            
    def get_range(self, range):
        if range.find('!') > 0:
            sheet,range = range.split('!')
            return self.workbook[sheet].iter_rows(range)
        else:        
            return self.workbook.active.iter_rows(range)

    def get_used_range(self):
        return self.workbook.active.iter_rows()

    def get_active_sheet(self):
        return self.workbook.active.title
    
    def get_cell(self,r,c):
        return self.workbook.active.cell(None,r,c)
        
    def get_value(self,r,c):
        return self.get_cell(r, c).value
    
    def set_value(self,r,c,val):
        self.get_cell(r, c).Value = val

    def get_formula(self,r,c):
        cell = self.get_cell(r, c)
        if cell.data_type is Cell.TYPE_FORMULA:
            return cell.value
        else:
            return None
    
    def has_formula(self,range):
        tuples = self.get_range(range)
        for row in tuples:
            for cell in row:
                if cell.data_type is Cell.TYPE_FORMULA:
                    return True
        return False   

    def get_formula_from_range(self,range):
        formulas = []
        tuples = self.get_range(range)
        for row in tuples:
            for cell in row:
                if cell.data_type is Cell.TYPE_FORMULA:
                    formulas.append(cell.value)
        return formulas    
    
    def get_formula_or_value(self,range):
        list = []
        tuples = self.get_range(range)
        for row in tuples:
            for cell in row:
                list.append(cell.value)
        return list    

    def get_row(self,row):
        return [self.get_value(row,col+1) for col in range(self.workbook.active.max_column)]

    def set_calc_mode(self,automatic=True):
        raise Exception('Not implemented')

    def set_screen_updating(self,update):
        raise Exception('Not implemented')

    def run_macro(self,macro):
        raise Exception('Not implemented')
