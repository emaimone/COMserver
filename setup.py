#!/usr/bin/python3
# -*- coding: utf-8 -*-
# Created by: python.exe -m py2exe COMbuilder.py -W setup.py

from distutils.core import setup
import py2exe

class Target(object):
    def __init__(self, **kw):
        self.__dict__.update(kw)

    def copy(self):
        return Target(**self.__dict__)

    def __setitem__(self, name, value):
        self.__dict__[name] = value

COMbldrObj = Target(
    script="COMbuilder.py", 
    other_resources = []
    )

py2exe_options = dict(
    packages = [],
    optimize=0,
    compressed=False, 
    bundle_files=3,
    dist_dir='dist03',
    )

setup(name="name",
      console=[COMbldrObj],
      ctypes_com_server=["COMbuilder"],
      options={"py2exe": py2exe_options},
      )

