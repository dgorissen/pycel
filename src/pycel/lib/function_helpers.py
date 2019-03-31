import functools

from pycel.excelutil import math_wrap

cse_array_functions = dict(
    # function_parameter which takes cse array input
    hlookup=0,
    iferror=0,
    lookup=0,
    vlookup=0,
    x_if=0,
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

                elif name in cse_array_functions:
                    func = cse_array_wrapper(func, cse_array_functions[name])

                name_space[name] = func

    return not_found
