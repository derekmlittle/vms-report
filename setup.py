#py2exe setup file

from distutils.core import setup
import py2exe

setup(
    windows=[{'script': 'vms.py'}],
)
