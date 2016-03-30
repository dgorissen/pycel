from openpyxl import load_workbook
from openpyxl.cell import Cell

import os
from os import path


class ExcelWrapper(object):
    
    def __init__(self, filename):
        
        super(ExcelWrapper,self).__init__()
        
        self.filename = path.abspath(filename)
        # TODO: discriminate between worksheet & workbook ranged names
        
        '''
        self.rangednames = np.zeros(shape = (int(self.app.ActiveWorkbook.Names.Count),1), dtype=[('id', 'int_'), ('name', 'S200'), ('formula', 'S200')])
        for i in range(0, self.app.ActiveWorkbook.Names.Count):
            self.rangednames[i]['id'] = int(i+1)       
            self.rangednames[i]['name'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Name)        
           self.rangednames[i]['formula'] = str(self.app.ActiveWorkbook.Names.Item(i+1).Value)
        '''
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
        if cell.data_type in Cell.TYPE_FORMULA:
            return cell.value
        else:
            return None
    
    def has_formula(self,range):
        tuples = self.get_range(range)
        for row in tuples:
            for cell in row:
                if cell.data_type in Cell.TYPE_FORMULA:
                    return True
        return False   

    def get_formula_from_range(self,range):
        formulas = []
        tuples = self.get_range(range)
        for row in tuples:
            for cell in row:
                if cell.data_type in Cell.TYPE_FORMULA:
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
