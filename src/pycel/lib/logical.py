"""
Python equivalents of excel logical functions (bools)
"""

from pycel.excelutil import ERROR_CODES, VALUE_ERROR


# AND function
# Returns TRUE if all of its arguments are TRUE
# Excel reference: https://support.office.com/en-us/article/
#   and-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9


# FALSE function
# Returns the logical value FALSE
# Excel reference: https://support.office.com/en-us/article/
#   false-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904


def x_if(test, true_value, false_value=0):
    # Excel reference: https://support.office.com/en-us/article/
    #   IF-function-69AED7C9-4E8A-4755-A9BC-AA8BBFF73BE2

    if test in ERROR_CODES:
        return test

    if isinstance(test, str):
        if test.lower() in ('true', 'false'):
            test = len(test) == 4
        else:
            return VALUE_ERROR

    return true_value if test else false_value


def iferror(arg, value_if_error):
    # Excel reference: https://support.office.com/en-us/article/
    #   IFERROR-function-C526FD07-CAEB-47B8-8BB6-63F3E417F611

    return value_if_error if arg in ERROR_CODES else arg


# IFNA function
# Excel 2013
# Returns the value you specify if the expression resolves to #N/A,
# otherwise returns the result of the expression
# Excel reference: https://support.office.com/en-us/article/
#   ifna-function-6626c961-a569-42fc-a49d-79b4951fd461


# IFS function
# Excel 2016
# Checks whether one or more conditions are met and returns a value that
# corresponds to the first TRUE condition.
# Excel reference: https://support.office.com/en-us/article/
#   ifs-function-36329a26-37b2-467c-972b-4a39bd951d45


# NOT function
# Reverses the logic of its argument
# Excel reference: https://support.office.com/en-us/article/
#   not-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77


# OR function
# Returns TRUE if any argument is TRUE
# Excel reference: https://support.office.com/en-us/article/
#   or-function-7d17ad14-8700-4281-b308-00b131e22af0


# SWITCH function
# Excel 2016
# Evaluates an expression against a list of values and returns the result
# corresponding to the first matching value. If there is no match, an optional
# default value may be returned.
# Excel reference: https://support.office.com/en-us/article/
#   switch-function-47ab33c0-28ce-4530-8a45-d532ec4aa25e


# TRUE function
# Returns the logical value TRUE
# Excel reference: https://support.office.com/en-us/article/
#   true-function-7652c6e3-8987-48d0-97cd-ef223246b3fb


# XOR function
# Excel 2013
# Returns a logical exclusive OR
# Excel reference: https://support.office.com/en-us/article/
#   xor-function-1548d4c2-5e47-4f77-9a92-0533bba14f37
