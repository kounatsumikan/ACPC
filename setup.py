# coding: utf-8

from setuptools import setup, find_packages

NAME = 'pandas_2_pptx'
VERSION = '0.1.0'


setup(
    name=NAME,
    version=VERSION,
    install_requires=["pandas", "python-pptx"],
    packages=find_packages('lib'),
    package_dir={'': 'lib'},
)
