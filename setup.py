'''
Created on Jan 30, 2014

@author: Costa Halicea
'''
from distutils.core import setup
import py2exe, sys, os  # @UnresolvedImport @UnusedImport

sys.argv.append('py2exe')

setup(windows=[{'script': 'ekp_xml_converter.py'}], \
            options={"py2exe": {"includes": ["Tkinter", \
            "tkFileDialog", "tkMessageBox", "xlrd", "json", "codecs","os"], \
            'bundle_files': 3, 'compressed': False}}, \
            zipfile = None)
