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
from unittest import mock

import pytest
import restructuredtext_lint

import pycel


@pytest.fixture(scope='session')
def changes_rst():
    docs_path = os.path.join(
        os.path.dirname(__file__), '../CHANGES.rst')
    with open(docs_path, 'r') as f:
        changes = f.readlines()
    return changes


@pytest.fixture(scope='session')
def doc_versions(changes_rst):
    return [
        l1.split()[0].strip() for l1, l2 in zip(changes_rst, changes_rst[1:])
        if l2.startswith('===')
    ]


def test_module_version():
    assert pycel.version.__version__ == pycel.__version__


def test_module_version_components():
    loose = LooseVersion(pycel.__version__).version
    for component in loose:
        assert isinstance(component, int) or component in ('a', 'b', 'rc')


def test_docs_version(doc_versions):
    assert pycel.version.__version__ == doc_versions[0]


def test_docs_versions(doc_versions):
    for v1, v2 in zip(doc_versions, doc_versions[1:]):
        assert LooseVersion(v1) > LooseVersion(v2)


def test_binder_requirements():
    binder_reqs_file = '../binder/requirements.txt'
    if os.path.exists(binder_reqs_file):
        with open(binder_reqs_file, 'r') as f:
            binder_reqs = sorted(l.strip() for l in f.readlines())

        with mock.patch('setuptools.setup') as setup:
            cwd = os.getcwd()
            os.chdir('..')
            with open('setup.py', 'r') as f:
                exec(f.read())
            os.chdir(cwd)
            setup_reqs = setup.mock_calls[0][2]['install_requires']

            # the binder requirements also include the optional graphing libs
            assert binder_reqs == sorted(setup_reqs + ['matplotlib', 'pydot'])


def test_changes_rst(changes_rst):
    def check_errors(to_check):
        return [err for err in to_check if err.level > 1]

    errors = restructuredtext_lint.lint('\n'.join(changes_rst))
    assert not check_errors(errors)

    if os.path.exists('../setup.py'):
        with mock.patch('setuptools.setup') as setup:
            cwd = os.getcwd()
            os.chdir('..')
            with open('setup.py', 'r') as f:
                exec(f.read())
            os.chdir(cwd)
            setup_long_desc = setup.mock_calls[0][2]['long_description']
            errors = restructuredtext_lint.lint(setup_long_desc)
            assert not check_errors(errors)
