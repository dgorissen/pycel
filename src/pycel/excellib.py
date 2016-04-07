'''
Python equivalents of various excel functions
'''
from __future__ import division
import numpy as np
from math import log
from pycel.excelutil import flatten

######################################################################################
# A dictionary that maps excel function names onto python equivalents. You should
# only add an entry to this map if the python name is different to the excel name
# (which it may need to be to  prevent conflicts with existing python functions 
# with that name, e.g., max).

# So if excel defines a function foobar(), all you have to do is add a function
# called foobar to this module.  You only need to add it to the function map,
# if you want to use a different name in the python code. 

# Note: some functions (if, pi, atan2, and, or, array, ...) are already taken care of
# in the FunctionNode code, so adding them here will have no effect.
FUNCTION_MAP = {
      "ln":"xlog",
      "min":"xmin",
      "min":"xmin",
      "max":"xmax",
      "sum":"xsum",
      "gammaln":"lgamma"
      }

######################################################################################
# List of excel equivalent functions
# TODO: needs unit testing

def value(text):
    # make the distinction for naca numbers
    if text.find('.') > 0:
        return float(text)
    else:
        return int(text)

def xlog(a):
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
    
    # build the vandermonde matrix
    A = np.vander(X, degree+1)
    
    if not const:
        # force the intercept to zero
        A[:,-1] = np.zeros((1,len(X)))
    
    # perform the fit
    (coefs, residuals, rank, sing_vals) = np.linalg.lstsq(A, Y)
        
    return coefs

def npv(*args):
    discount_rate = args[0]
    cashflow = args[1]
    return sum([float(x)*(1+discount_rate)**-(i+1) for (i,x) in enumerate(cashflow)])

def match(lookup_value, lookup_array, match_type=1):
    
    def type_convert(value):
        if type(value) == str:
            value = value.lower()
        elif type(value) == int:
            value = float(value)

        return value;

    lookup_value = type_convert(lookup_value)

    if match_type == 1:
        # Verify ascending sort
        posMax = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])

            if i is not len(lookup_array)-1 and current > type_convert(lookup_array[i+1]):
                raise Exception('for match_type 0, lookup_array must be sorted ascending')
            if current <= lookup_value:
                posMax = i 
        if posMax == -1:
            raise ('no result in lookup_array for match_type 0')
        return posMax +1 #Excel starts at 1

    elif match_type == 0:
        # No string wildcard
        return [type_convert(x) for x in lookup_array].index(lookup_value) + 1

    elif match_type == -1:
        # Verify descending sort
        posMin = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])

            if i is not len(lookup_array)-1 and current < type_convert(lookup_array[i+1]):
               raise ('for match_type 0, lookup_array must be sorted descending')
            if current >= lookup_value:
               posMin = i 
        if posMin == -1:
            raise Exception('no result in lookup_array for match_type 0')
        return posMin +1 #Excel starts at 1

def mod(nb, q):
    if not isinstance(nb, (int, long)):
        raise TypeError("%s is not an integer" % str(nb))
    elif not isinstance(q, (int, long)):
        raise TypeError("%s is not an integer" % str(q))
    else:
        return nb % q

if __name__ == '__main__':
    pass
