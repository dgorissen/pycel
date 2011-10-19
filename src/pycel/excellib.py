'''
Created on 13 Sep 2011

@author: dg2d09
'''
from __future__ import division
import numpy as np
from math import log
from pycel.excelutil import flatten
   
def value(text):
    # make the distinction for naca numbers
    if text.find('.') > 0:
        return float(text)
    else:
        return int(text)

# TODO: inline
def xlog(a):
    #print type(a)
    if isinstance(a,(list,tuple,np.ndarray)):
        return [log(x) for x in flatten(a)]
    else:
        #print a
        return log(a)

def xmax(*args):
    # ignore non numeric cells
    data = [x for x in flatten(args) if isinstance(x,(int,float))]
    
    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return max(data)

def xmin(*args):
    # ignore non numeric cells
    data = [x for x in flatten(args) if isinstance(x,(int,float))]
    
    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return min(data)

#TODO does not take into account mixture of lists & scalars
def xsum(*args):
    # ignore non numeric cells
    data = [x for x in flatten(args) if isinstance(x,(int,float))]
    
    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return sum(data)

def average(*args):
    l = list(flatten(*args))
    return sum(l) / len(l)
    
def right(text,n):
    #TODO: hack to deal with naca section numbers
    if isinstance(text, unicode) or isinstance(text,str):
        return text[-n:]
    else:
        # TODO: get rid of the decimal
        return str(int(text))[-n:]

#http://www.techonthenet.com/excel/formulas/match.php
#TODO: we only care about numeric data which is not exactly what excel does, also not the cleanest code + overlap with lookup
def match(value,array,match_type=1):
    
    data = array
    
    if match_type == 1:
        # index of the last numeric value
        lastnum = -1
        for i,v in enumerate(data):
            if isinstance(v,(int,float)):
                if v > value:
                    break
                else:
                    lastnum = i
                

        if lastnum < 0:
            raise Exception("No numeric data found in the lookup range")
        else:
            if i == 0:
                raise Exception("All values in the lookup range are bigger than %s" % value)
            else:
                if i >= len()-1:
                    return lastnum+1
                else:
                    return i-1+1  #want a 1 based index
        
        
    elif match_type == 0:
        res = [i for i,x in enumerate(data) if x == value]
        if not res:
            raise Exception('match: %s not found in array!' % value)
        
        return res[0] + 1
    
    elif match_type == -1:
        # index of the last numeric value
        lastnum = -1
        for i,v in enumerate(data):
            #TODO: we only care about numeric data which is not exactly what excel does
            if isinstance(v,(int,float)):
                if v < value:
                    break
                else:
                    lastnum = i
                

        if lastnum < 0:
            raise Exception("No numeric data found in the lookup range")
        else:
            if i == 0:
                raise Exception("All values in the lookup range are smaller than %s" % value)
            else:
                if i >= len()-1:
                    return lastnum+1
                else:
                    return i-1+1  #want a 1 based index
    else:
        raise Exception('Invalid match type ' + match_type)
    
    return res
    
    
def index(*args):
    array = args[0]
    row = args[1]
    
    if len(args) == 3:
        col = args[2]
    else:
        col = 1
        
    if isinstance(array[0],(list,tuple,np.ndarray)):
        # rectangular array
        array[row-1][col-1]
    elif row == 1 or col == 1:
        return array[row-1] if col == 1 else array[col-1]
    else:
        raise Exception("index (%s,%s) out of range for %s" %(row,col,array))
        

def lookup(value, lookup_range, result_range):
    
    # TODO
    if not isinstance(value,(int,float)):
        raise Exception("Non numeric lookups (%s) not supported" % value)
    
    # TODO: note, may return the last equal value
    
    # index of the last numeric value
    lastnum = -1
    for i,v in enumerate(lookup_range):
        if isinstance(v,(int,float)):
            if v > value:
                break
            else:
                lastnum = i
                

    if lastnum < 0:
        raise Exception("No numeric data found in the lookup range")
    else:
        if i == 0:
            raise Exception("All values in the lookup range are bigger than %s" % value)
        else:
            if i >= len(lookup_range)-1:
                # return the biggest number smaller than value
                return result_range[lastnum]
            else:
                return result_range[i-1]

def linest(*args, **kwargs):

    Y = args[0]
    X = args[1]
    
    if len(args) == 3:
        const = args[2]
        if isinstance(const,str):
            const = (const.lower() == "true")
    else:
        const = True
        
    degree = kwargs.get('degree',1)
    skip_blanks = kwargs.get('skip_blanks',False)
    
    #TODO all rather hackish
    if skip_blanks:
        X = [x for x in X if isinstance(x,(int,float))]
        Y = [y for y in Y if isinstance(y,(int,float))]
        # trim to the same length
        l = min(len(X),len(Y))
        X = X[:l]
        Y = Y[:l]
        
    # build the vandermonde matrix
    A = np.vander(X, degree+1)
    
    if not const:
        # force the intercept to zero
        A[:,-1] = np.zeros((1,len(X)))
    
    # perform the fit
    (coefs, residuals, rank, sing_vals) = np.linalg.lstsq(A, Y)
        
    return coefs #TODO: returns ndarray, not list, may cause problems

# same as linest but ignoring non numeric cells
def linestmario(*args, **kwargs):
    Y = args[0]
    X = args[1]
    degree = args[2]
    
    if len(args) == 4:
        const = args[3]
    else:
        const = True
    
    kwargs['skip_blanks'] = True
    kwargs['degree'] = degree
    
    return linest(Y,X,const,**kwargs)
    
if __name__ == '__main__':
    pass