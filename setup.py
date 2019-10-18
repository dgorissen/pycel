#!/usr/bin/env python

"""Setup script for packaging pycel.

To release:
    Update /src/pycel/version.py, /CHANGES.rst

Run tests with:
    tox

To build a package for distribution:
    python setup.py sdist bdist_wheel

and upload it to the PyPI with:
    twine upload --verbose dist/*

to install a link for development work:
    pip install -e .

"""

from setuptools import find_packages, setup

# see StackOverflow/458550
exec(open('src/pycel/version.py').read())


# Create long description from README.rst and docs/source/CHANGES.rst.
# PYPI page will contain complete changelog.
long_description = u'{}\n\n\nChange Log\n==========\n\n\n{}'.format(
    open('README.rst', 'r', encoding='utf-8').read(),
    open('CHANGES.rst', 'r', encoding='utf-8').read()
)

with open('test-requirements.txt') as f:
    tests_require = f.readlines()


setup(
    name='pycel',
    version=__version__,  # noqa: F821
    packages=find_packages('src'),
    package_dir={'': 'src'},
    description='A library for compiling excel spreadsheets to python code '
                '& visualizing them as a graph',
    keywords='excel compiler formula parser',
    url='https://github.com/stephenrauch/pycel',
    project_urls={
        # 'Documentation': 'https://pycel.readthedocs.io/en/stable/',
        'Tracker': 'https://github.com/stephenrauch/pycel/issues',
    },
    tests_require=tests_require,
    test_suite='pytest',
    install_requires=[
        'networkx>=2.0,<2.5',
        'numpy',
        'openpyxl>=2.6.2',
        'python-dateutil',
        'ruamel.yaml',
    ],
    python_requires='>=3.5',
    author='Dirk Gorissen, Stephen Rauch',
    author_email='dgorissen@gmail.com',
    maintainer='Stephen Rauch',
    maintainer_email='stephen.rauch+pycel@gmail.com',
    long_description=long_description,
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)
