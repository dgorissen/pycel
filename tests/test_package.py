from distutils.version import LooseVersion
import os

import pytest

import pycel


@pytest.fixture(scope='session')
def doc_versions():
    docs_path = os.path.join(
        os.path.dirname(__file__), '../CHANGES.rst')
    with open(docs_path, 'r') as f:
        changes = f.readlines()

    return [
        l1.split()[0].strip() for l1, l2 in zip(changes, changes[1:])
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
        with open('../binder/requirements.txt', 'r') as f:
            binder_reqs = sorted(l.strip() for l in f.readlines())

        from unittest import mock
        with mock.patch('setuptools.setup') as setup:
            cwd = os.getcwd()
            os.chdir('..')
            with open('setup.py', 'r') as f:
                exec(f.read())
            os.chdir(cwd)
            setup_reqs = setup.mock_calls[0][2]['install_requires']

            # the binder requirements also include the optional graphing libs
            assert binder_reqs == sorted(setup_reqs + ['matplotlib', 'pydot'])
