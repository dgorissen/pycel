from setuptools import find_packages, setup

# see StackOverflow/458550
exec(open('src/pycel/version.py').read())

setup(
    name='Pycel',
    version=__version__,  # noqa: F821
    packages=find_packages('src'),
    package_dir={'': 'src'},
    description='A library for compiling excel spreadsheets to python code '
                '& visualizing them as a graph',
    url='https://github.com/dgorissen/pycel',
    tests_require=['pytest'],
    test_suite='pytest',
    install_requires=[
        'networkx>=2.0',
        'openpyxl>=2.5.12',
        'numpy',
        'ruamel.yaml',
        'pydot',
    ],
    author='Dirk Gorissen',
    author_email='dgorissen@gmail.com',
    long_description="Pycel is a small python library that can translate an "
                     "Excel spreadsheet into executable python code which "
                     "can be run independently of Excel. The python code is "
                     "based on a graph and uses caching & lazy evaluation to "
                     "ensure (relatively) fast execution. The graph can be "
                     "exported and analyzed using tools like Gephi. See the "
                     "contained example for an illustration.",
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License ::  OSI Approved ',
    ]
)
