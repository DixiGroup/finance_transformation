from distutils.core import setup
import py2exe

setup(console=[{'script':'finance.py'}],
      options={"py2exe":{"includes":["xlrd", "xlsxwriter"]}})
