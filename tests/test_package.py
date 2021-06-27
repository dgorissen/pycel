# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os
from distutils.version import LooseVersion
from pathlib import Path
from unittest import mock

import pytest
import restructuredtext_lint

import pycel

repo_root = Path(__file__).parents[1]


@pytest.fixture(scope='session')
def changes_rst():
    with open(repo_root / 'CHANGES.rst', 'r') as f:
        return f.readlines()


@pytest.fixture(scope='session')
def setup_py():
    with mock.patch('setuptools.setup'), mock.patch('setuptools.find_packages'):
        cwd = os.getcwd()
        os.chdir(repo_root)
        import importlib
        setup = importlib.import_module('setup')
        os.chdir(cwd)
        return setup


def test_module_version():
    assert pycel.version.__version__ == pycel.__version__


def test_module_version_components():
    loose = LooseVersion(pycel.__version__).version
    for component in loose:
        assert isinstance(component, int) or component in ('a', 'b', 'rc')


def test_docs_versions(changes_rst):
    doc_versions = [
        l1.split()[0].strip() for l1, l2 in zip(changes_rst, changes_rst[1:])
        if l2.startswith('===')
    ]

    for version in doc_versions:
        assert version[0] == '['
        assert version[-1] == ']'

    assert doc_versions[0] == '[unreleased]'
    assert pycel.version.__version__ == doc_versions[1][1:-1]

    for v1, v2 in zip(doc_versions[1:], doc_versions[2:]):
        assert LooseVersion(v1[1:-1]) > LooseVersion(v2[1:-1])


def test_binder_requirements(setup_py):
    binder_reqs_file = '../binder/requirements.txt'
    if os.path.exists(binder_reqs_file):
        with open(binder_reqs_file, 'r') as f:
            binder_reqs = sorted(line.strip() for line in f.readlines())

        setup_reqs = setup_py.setup.mock_calls[0][2]['install_requires']

        # the binder requirements also include the optional graphing libs
        assert binder_reqs == sorted(setup_reqs + ['matplotlib', 'pydot'])


def test_changes_rst(changes_rst, setup_py):
    def check_errors(to_check):
        return [err for err in to_check if err.level > 1]

    errors = restructuredtext_lint.lint('\n'.join(changes_rst))
    assert not check_errors(errors)

    errors = restructuredtext_lint.lint(setup_py.long_description)
    assert not check_errors(errors)
