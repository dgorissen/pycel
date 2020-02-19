"""This file sets the path so that pycel code can be imported.

The idea of a context file is from https://docs.python-guide.org/writing/structure/
"""

import pathlib
import sys

src_path = pathlib.Path('../src/')

# Note: append does not support Path
sys.path.append(str(src_path.resolve()))

import pycel

