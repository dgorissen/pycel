from __future__ import division
from itertools import izip
import collections
import functools
import re
import string

#TODO: only supports rectangular ranges
class CellRange(object):
    def __init__(self,address,sheet=None):

        self.__address = address.replace('$','')
        
        sh,start,end = split_range(address)
        if not sh and not sheet:
            raise Exception("Must pass in a sheet")
        
        # make sure the address is always prefixed with the range
        if sh:
            sheet = sh
        else:
            self.__address = sheet + "!" + self.__address
        
        addr,nrows,ncols = resolve_range(address,sheet=sheet)
        
        # dont allow messing with these params
        self.__celladdr = addr
        self.__nrows = nrows
        self.__ncols = ncols
        self.__sheet = sheet
        
        self.value = None

    def __str__(self):
        return self.__address 
    
    def address(self):
        return self.__address
    
    @property
    def celladdrs(self):
        return self.__celladdr
    
    @property
    def nrows(self):
        return self.__nrows

    @property
    def ncols(self):
        return self.__ncols

    @property
    def sheet(self):
        return self.__sheet
    
class Cell(object):
    ctr = 0
    
    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr
    
    def __init__(self, address, sheet, value=None, formula=None):
        super(Cell,self).__init__()
        
        # remove $'s
        address = address.replace('$','')
        
        sh,c,r = split_address(address)
        
        # both are empty
        if not sheet and not sh:
            raise Exception("Sheet name may not be empty for cell address %s" % address)
        # both exist but disagree
        elif sh and sheet and sh != sheet:
            raise Exception("Sheet name mismatch for cell address %s: %s vs %s" % (address,sheet, sh))
        elif not sh and sheet:
            sh = sheet 
        else:
            pass
                
        # we assume a cell's location can never change
        self.__sheet = str(sheet)
        self.__formula = str(formula) if formula else None
        
        self.__sheet = sh
        self.__col = c
        self.__row = int(r)
        self.__col_idx = col2num(c)
            
        self.value = str(value) if isinstance(value,unicode) else value
        self.python_expression = None
        self._compiled_expression = None
        
        # every cell has a unique id
        self.__id = Cell.next_id()

    @property
    def sheet(self):
        return self.__sheet

    @property
    def row(self):
        return self.__row

    @property
    def col(self):
        return self.__col

    @property
    def formula(self):
        return self.__formula

    @property
    def id(self):
        return self.__id

    @property
    def compiled_expression(self):
        return self._compiled_expression

    # code objects are not serializable
    def __getstate__(self):
        d = dict(self.__dict__)
        f = '_compiled_expression'
        if f in d: del d[f]
        return d

    def __setstate__(self, d):
        self.__dict__.update(d)
        self.compile() 
    
    def clean_name(self):
        return self.address().replace('!','_').replace(' ','_')
        
    def address(self, absolute=True):
        if absolute:
            return "%s!%s%s" % (self.__sheet,self.__col,self.__row)
        else:
            return "%s%s" % (self.__col,self.__row)
    
    def address_parts(self):
        return (self.__sheet,self.__col,self.__row,self.__col_idx)
        
    def compile(self):
        if not self.python_expression: return
        
        # if we are a constant string, surround by quotes
        if isinstance(self.value,(str,unicode)) and not self.formula and not self.python_expression.startswith('"'):
            self.python_expression='"' + self.python_expression + '"'
        
        try:
            self._compiled_expression = compile(self.python_expression,'<string>','eval')
        except Exception as e:
            raise Exception("Failed to compile cell %s with expression %s: %s" % (self.address(),self.python_expression,e)) 
    
    def __str__(self):
        if self.formula:
            return "%s%s" % (self.address(), self.formula)
        else:
            return "%s=%s" % (self.address(), self.value)

    @staticmethod
    def inc_col_address(address,inc):
        sh,col,row = split_address(address)
        return "%s!%s%s" % (sh,num2col(col2num(col) + inc),row)

    @staticmethod
    def inc_row_address(address,inc):
        sh,col,row = split_address(address)
        return "%s!%s%s" % (sh,col,row+inc)
        
    @staticmethod
    def resolve_cell(excel, address, sheet=None):
        r = excel.get_range(address)
        f = r.Formula if r.Formula.startswith('=') else None
        v = r.Value
        
        sh,c,r = split_address(address)
        
        # use the sheet specified in the cell, else the passed sheet
        if sh: sheet = sh

        c = Cell(address,sheet,value=v, formula=f)
        return c

    @staticmethod
    def make_cells(excel, range, sheet=None):
        cells = [];

        if is_range(range):
            # use the sheet specified in the range, else the passed sheet
            sh,start,end = split_range(range)
            if sh: sheet = sh

            ads,numrows,numcols = resolve_range(range)
            # ensure in the same nested format as fs/vs will be
            if numrows == 1:
                ads = [ads]
            elif numcols == 1:
                ads = [[x] for x in ads]
                
            # get everything in blocks, is faster
            r = excel.get_range(range)
            fs = r.Formula
            vs = r.Value
            
            for it in (list(izip(*x)) for x in izip(ads,fs,vs)):
                row = []
                for c in it:
                    a = c[0]
                    f = c[1] if c[1] and c[1].startswith('=') else None
                    v = c[2]
                    cl = Cell(a,sheet,value=v, formula=f)
                    row.append(cl)
                cells.append(row)
            
            #return as vector
            if numrows == 1:
                cells = cells[0]
            elif numcols == 1:
                cells = [x[0] for x in cells]
            else:
                pass
        else:
            c = Cell.resolve_cell(excel, range, sheet=sheet)
            cells.append(c)

            numrows = 1
            numcols = 1
            
        return (cells,numrows,numcols)
    
def is_range(address):
    return address.find(':') > 0

def split_range(rng):
    if rng.find('!') > 0:
        sh,r = rng.split("!")
        start,end = r.split(':')
    else:
        sh = None
        start,end = rng.split(':')
    
    return sh,start,end
        
def split_address(address):
    sheet = None
    if address.find('!') > 0:
        sheet,address = address.split('!')
    
    #ignore case
    address = address.upper()
    
    # regular <col><row> format    
    if re.match('^[A-Z\$]+[\d\$]+$', address):
        col,row = filter(None,re.split('([A-Z\$]+)',address))
    # R<row>C<col> format
    elif re.match('^R\d+C\d+$', address):
        row,col = address.split('C')
        row = row[1:]
    # R[<row>]C[<col>] format
    elif re.match('^R\[\d+\]C\[\d+\]$', address):
        row,col = address.split('C')
        row = row[2:-1]
        col = col[2:-1]
    else:
        raise Exception('Invalid address format ' + address)
    
    return (sheet,col,row)

def resolve_range(rng, flatten=False, sheet=''):
    
    sh, start, end = split_range(rng)
    
    if sh and sheet:
        if sh != sheet:
            raise Exception("Mismatched sheets %s and %s" % (sh,sheet))
        else:
            sheet += '!'
    elif sh and not sheet:
        sheet = sh + "!"
    elif sheet and not sh:
        sheet += "!"
    else:
        pass

    # single cell, no range
    if not is_range(rng):  return ([sheet + rng],1,1)

    sh, start_col, start_row = split_address(start)
    sh, end_col, end_row = split_address(end)
    start_col_idx = col2num(start_col)
    end_col_idx = col2num(end_col);

    start_row = int(start_row)
    end_row = int(end_row)

    # single column
    if  start_col == end_col:
        nrows = end_row - start_row + 1
        data = [ "%s%s%s" % (s,c,r) for (s,c,r) in zip([sheet]*nrows,[start_col]*nrows,range(start_row,end_row+1))]
        return data,len(data),1
    
    # single row
    elif start_row == end_row:
        ncols = end_col_idx - start_col_idx + 1
        data = [ "%s%s%s" % (s,num2col(c),r) for (s,c,r) in zip([sheet]*ncols,range(start_col_idx,end_col_idx+1),[start_row]*ncols)]
        return data,1,len(data)
    
    # rectangular range
    else:
        cells = []
        for r in range(start_row,end_row+1):
            row = []
            for c in range(start_col_idx,end_col_idx+1):
                row.append(sheet + num2col(c) + str(r))
                
            cells.append(row)
    
        if flatten:
            # flatten into one list
            l = flatten(cells)
            return l,1,len(l)
        else:
            return cells, len(cells), len(cells[0]) 

# e.g., convert BA -> 53
def col2num(col):
    
    if not col:
        raise Exception("Column may not be empty")
    
    tot = 0
    for i,c in enumerate([c for c in col[::-1] if c != "$"]):
        if c == '$': continue
        tot += (ord(c)-64) * 26 ** i
    return tot

# convert back
def num2col(num):
    
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)
    
    s = ''
    q = num
    while q > 0:
        (q,r) = divmod(q,26)
        if r == 0:
            q = q - 1
            r = 26
        s = string.ascii_uppercase[r-1] + s
    return s

def address2index(a):
    sh,c,r = split_address(a)
    return (col2num(c),int(r))

def index2addres(c,r,sheet=None):
    return "%s%s%s" % (sheet + "!" if sheet else "", num2col(c), r)

def get_linest_degree(excel,cl):
    # TODO: assumes a row or column of linest formulas & that all coefficients are needed

    sh,c,r,ci = cl.address_parts()
    # figure out where we are in the row

    # to the left
    i = ci - 1
    while i > 0:
        f = excel.get_formula_from_range(index2addres(i,r))
        if f is None or f != cl.formula:
            break
        else:
            i = i - 1
        
    # to the right
    j = ci + 1
    while True:
        f = excel.get_formula_from_range(index2addres(j,r))
        if f is None or f != cl.formula:
            break
        else:
            j = j + 1
    
    # assume the degree is the number of linest's
    degree =  (j - i - 1) - 1  #last -1 is because an n degree polynomial has n+1 coefs

    # which coef are we (left most coef is the coef for the highest power)
    coef = ci - i 

    # no linests left or right, try looking up/down
    if degree == 0:
        # up
        i = r - 1
        while i > 0:
            f = excel.get_formula_from_range("%s%s" % (c,i))
            if f is None or f != cl.formula:
                break
            else:
                i = i - 1
            
        # down
        j = r + 1
        while True:
            f = excel.get_formula_from_range("%s%s" % (c,j))
            if f is None or f != cl.formula:
                break
            else:
                j = j + 1

        degree =  (j - i - 1) - 1
        coef = r - i
    
    # if degree is zero -> only one linest formula -> linear regression -> degree should be one
    return (max(degree,1),coef) 

def flatten(l):
    for el in l:
        if isinstance(el, collections.Iterable) and not isinstance(el, basestring):
            for sub in flatten(el):
                yield sub
        else:
            yield el

def uniqueify(seq):
    seen = set()
    seen_add = seen.add
    return [ x for x in seq if x not in seen and not seen_add(x)]

if __name__ == '__main__':
    pass