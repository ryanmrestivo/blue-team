#!/usr/bin/env python3

from setuptools import setup, find_packages

VERSION = 1.11
DESCRIPTION = 'Access windows anitmalware interface using python'
with open('README.md', 'r', encoding = 'utf-8') as fh:
    long_description = fh.read()

setup(
    name='pyamsi',
    version=VERSION,
    license='MIT',
    author="Olorunfemi-Ojo Tomiwa",
    author_email='ot.server1@outlook.com',
    description=DESCRIPTION,
    long_description_content_type='text/markdown',
    long_description=long_description,
    package_dir={'':'src'},
    packages=find_packages(where='src'),
    include_package_data=True,
    package_data={"pyamsi": ['amsiscanner.dll']},
    url='https://github.com/Tomiwa-Ot/py-amsi',
    keywords=['amsi', 'python-amsi', 'pyamsi']
)
