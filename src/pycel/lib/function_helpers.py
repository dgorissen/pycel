import functools

from pycel.excelutil import ERROR_CODES, math_wrap

cse_array_functions = dict(
    # function_parameter which takes cse array input
    hlookup=0,
    iferror=0,
    lookup=0,
    vlookup=0,
    x_if=0,
)

error_string_functions = dict(
    index=(0, 1, 2),
)


def cse_array_wrapper(f, arg_num):
    """wrapper to take cse array input and call function once per element"""

    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if (isinstance(args[arg_num], tuple) and
                isinstance(args[arg_num][0], tuple)):
            if arg_num == 0:
                return tuple(tuple(
                    f(x, *args[arg_num + 1:], **kwargs)
                    for x in row) for row in args[arg_num])
            else:
                return tuple(tuple(
                    f(*args[:arg_num], x, *args[arg_num + 1:], **kwargs)
                    for x in row) for row in args[arg_num])

        return f(*args, **kwargs)

    return wrapper


def error_string_wrapper(f, err_str_args=None):
    """wrapper to process error strings in arguments"""

    if err_str_args is None:
        err_str_args = error_string_functions[f.__name__]

    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        for arg_num in err_str_args:
            try:
                arg = args[arg_num]
            except IndexError:
                continue
            if isinstance(arg, str) and arg in ERROR_CODES:
                return arg

        return f(*args, **kwargs)

    return wrapper


def load_functions(names, name_space, modules):
    # load desired functions into namespace from modules
    not_found = set()
    for name in names:
        if name not in name_space:
            funcs = ((getattr(module, name, None), module)
                     for module in modules)
            func, module = next(
                (f for f in funcs if f[0] is not None), (None, None))
            if func is None:
                not_found.add(name)
            else:
                if module.__name__ == 'math':
                    func = math_wrap(func)
                    func = cse_array_wrapper(func, 0)

                else:
                    if name in error_string_functions:
                        func = error_string_wrapper(func)

                    if name in cse_array_functions:
                        func = cse_array_wrapper(
                            func, cse_array_functions[name])

                name_space[name] = func

    return not_found
