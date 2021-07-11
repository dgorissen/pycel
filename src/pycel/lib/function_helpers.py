# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import collections
import functools
import inspect
import sys

from pycel.excelutil import (
    AddressCell,
    AddressRange,
    coerce_to_number,
    coerce_to_string,
    ERROR_CODES,
    flatten,
    is_array_arg,
    is_number,
    NUM_ERROR,
    VALUE_ERROR,
)


FUNC_META = 'excel_func_meta'

ALL_ARG_INDICES = frozenset(range(512))

star_args = set()


def excel_helper(cse_params=None,
                 bool_params=None,
                 err_str_params=-1,
                 number_params=None,
                 str_params=None,
                 ref_params=None):
    """ Decorator to annotate a function with info on how to process params

    All parameters are encoded as:

        int >= 0: param number to check
        tuple of ints: params to check
        -1: check all params
        None: check no params

    :param cse_params: CSE Array Params.  If array are passed the function
        will be called multiple times, once for each value, and the result
        will be a CSE Array
    :param bool_params: params to coerce to bools
    :param err_str_params: params to check for error strings
    :param number_params: params to coerce to numbers
    :param str_params: params to coerce to strings
    :param ref_params: params which can remain as references
    :return: decorator
    """
    def mark(f):
        if any(param.kind == inspect.Parameter.VAR_POSITIONAL
               for param in inspect.signature(f).parameters.values()):
            star_args.add(f.__name__)

        setattr(f, FUNC_META, dict(
            cse_params=cse_params,
            bool_params=bool_params,
            err_str_params=err_str_params,
            number_params=number_params,
            str_params=str_params,
            ref_params=ref_params,
        ))
        return f
    return mark


# Decorator for generic excel function
excel_func = excel_helper()

# Decorator for generic excel math function (all params are numbers)
excel_math_func = excel_helper(
    cse_params=-1, err_str_params=-1, number_params=-1)


def apply_meta(f, meta=None, name_space=None):
    """Take the metadata applied by excel_helper and wrap accordingly"""
    meta = meta or getattr(f, FUNC_META, None)
    if meta:
        meta['name_space'] = name_space

        # find what all_params for this function should look like
        try:
            sig = inspect.signature(f)
            if any(param.kind == inspect.Parameter.VAR_KEYWORD
                   for param in sig.parameters.values()):
                raise RuntimeError(
                    f'Function {f.__name__}: **kwargs not allowed in signature.')
        except ValueError:
            # some built-ins do not have signature information
            sig = None  # pragma: no cover
        if sig and any(param.kind == inspect.Parameter.VAR_POSITIONAL
                       for param in sig.parameters.values()):
            all_params = ALL_ARG_INDICES
        else:
            all_params = set(range(getattr(getattr(f, '__code__', None), 'co_argcount', 0))
                             ) or ALL_ARG_INDICES

        # process error strings
        err_str_params = meta['err_str_params']
        if err_str_params is not None:
            f = error_string_wrapper(
                f, all_params if err_str_params == -1 else err_str_params)

        # process number parameters
        number_params = meta['number_params']
        if number_params is not None:
            f = nums_wrapper(
                f, all_params if number_params == -1 else number_params)

        # process str parameters
        str_params = meta['str_params']
        if str_params is not None:
            f = strs_wrapper(f, all_params if str_params == -1 else str_params)

        # process CSE parameters
        cse_params = meta['cse_params']
        if cse_params is not None:
            f = cse_array_wrapper(
                f, all_params if cse_params == -1 else cse_params)

        # process reference parameters
        ref_params = meta['ref_params']
        if ref_params != -1:
            if ref_params is None:
                ref_params = set()
            f = refs_wrapper(f, name_space, ref_params)

    return f, meta


def convert_params_indices(f, param_indices):
    """Given parameter indices, return a set of parameter indices to process

    :param f: function to check for arg count
    :param param_indices: params to check if CSE array
        int: param number to check
        tuple: params to check
    :return: set of parameter indices
    """
    if not isinstance(param_indices, collections.abc.Iterable):
        assert param_indices >= 0
        return {int(param_indices)}

    else:
        assert all(i >= 0 for i in param_indices)
        return set(map(int, param_indices))


def cse_array_wrapper(f, param_indices=None):
    """wrapper to take cse array input and call function once per element

    :param f: function to wrap
    :param param_indices: params to check if CSE array
        int: param number to check
        tuple: params to check
        None: check all params
    :return: wrapped function
    """
    param_indices = convert_params_indices(f, param_indices)

    def pick_args(args, cse_arg_nums, row, col):
        return (arg[row][col] if i in cse_arg_nums else arg
                for i, arg in enumerate(args))

    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        looper = (i for i in param_indices if i < len(args))
        cse_arg_nums = {arg_num for arg_num in looper if is_array_arg(args[arg_num])}

        if cse_arg_nums:
            a_cse_arg = next(iter(cse_arg_nums))
            num_rows = len(args[a_cse_arg])
            num_cols = len(args[a_cse_arg][0])

            return tuple(tuple(
                f(*pick_args(args, cse_arg_nums, row, col), **kwargs)
                for col in range(num_cols)) for row in range(num_rows))

        return f(*args, **kwargs)

    return wrapper


def nums_wrapper(f, param_indices=None):
    """wrapper for functions that take numbers, does excel style conversions

    :param f: function to wrap
    :param param_indices: params to coerce to numbers.
        int: param number to convert
        tuple: params to convert
        None: convert all params
    :return: wrapped function
    """
    param_indices = convert_params_indices(f, param_indices)

    @functools.wraps(f)
    def wrapper(*args):
        new_args = tuple(coerce_to_number(a, convert_all=True)
                         if i in param_indices else a
                         for i, a in enumerate(args))
        error = next((a for i, a in enumerate(new_args)
                      if i in param_indices and a in ERROR_CODES), None)
        if error:
            return error

        if any(i in param_indices and not is_number(a)
               for i, a in enumerate(new_args)):
            return VALUE_ERROR

        try:
            return f(*new_args)
        except ValueError as exc:
            if "math domain error" in str(exc):
                return NUM_ERROR
            raise  # pragma: no cover

    return wrapper


def strs_wrapper(f, param_indices=None):
    """wrapper for functions that take strings, does excel style conversions

    :param f: function to wrap
    :param param_indices: params to coerce to strings.
        int: param number to convert
        tuple: params to convert
        None: convert all params
    :return: wrapped function
    """
    param_indices = convert_params_indices(f, param_indices)

    @functools.wraps(f)
    def wrapper(*args):
        new_args = tuple(coerce_to_string(a)
                         if i in param_indices else a
                         for i, a in enumerate(args))
        error = next((a for i, a in enumerate(new_args)
                      if i in param_indices and a in ERROR_CODES), None)
        if error:
            return error

        return f(*new_args)

    return wrapper


def error_string_wrapper(f, param_indices=None):
    """wrapper to process error strings in arguments

    :param f: function to wrap
    :param param_indices: params to check for error strings.
        int: param number to check
        tuple: params to check
        None: check all params
    :return: wrapped function
    """
    param_indices = sorted(convert_params_indices(f, param_indices))

    @functools.wraps(f)
    def wrapper(*args):
        for arg_num in param_indices:
            try:
                arg = args[arg_num]
            except IndexError:
                break
            if isinstance(arg, str) and arg in ERROR_CODES:
                return arg
            elif isinstance(arg, tuple):
                error = next((a for a in flatten(arg)
                              if isinstance(a, str) and a in ERROR_CODES), None)
                if error is not None:
                    return error

        return f(*args)

    return wrapper


def refs_wrapper(f, name_space, param_indices=None):
    """wrapper to process references in arguments

    :param f: function to wrap
    :param param_indices: params to check for error strings.
        int: param number to check
        tuple: params to check
        None: check all params
    :return: wrapped function
    """
    param_indices = convert_params_indices(f, param_indices)

    _R_ = name_space.get('_R_')
    _C_ = name_space.get('_C_')

    def resolve_args(args):
        for arg_num, arg in enumerate(args):
            if arg_num in param_indices:
                yield arg
            elif isinstance(arg, AddressCell):
                # resolve cell if this is not reference param
                yield _C_(arg.address)
            elif isinstance(arg, AddressRange):
                # resolve range if this is not reference param
                yield _R_(arg.address)
            else:
                yield arg

    @functools.wraps(f)
    def wrapper(*args):
        return f(*tuple(resolve_args(args)))

    return wrapper


def built_in_wrapper(f, wrapper_marker, name_space):
    meta = getattr(wrapper_marker(lambda x: x), FUNC_META)  # pragma: no branch
    return apply_meta(f, meta, name_space)[0]


def load_functions(names, name_space, modules):
    # load desired functions into namespace from modules
    not_found = set()
    for name in names:
        if name not in name_space:
            funcs = ((getattr(module, name, None), module)
                     for module in modules)
            f, module = next(
                (f for f in funcs if f[0] is not None), (None, None))
            if f is None:
                not_found.add(name)
            else:
                if module.__name__ == 'math':
                    f = built_in_wrapper(
                        f, excel_math_func, name_space=name_space)
                else:
                    f, meta = apply_meta(f, name_space=name_space)
                name_space[name] = f

    return not_found


def load_to_test_module(load_from, load_to_name):
    # dynamic load the lib functions from 'load_from' and apply metadata
    load_to = sys.modules[load_to_name]
    for name in dir(load_from):
        obj = getattr(load_from, name)
        if callable(obj) and getattr(load_to, name, None) == obj:
            setattr(load_to, name, apply_meta(obj, name_space={})[0])
