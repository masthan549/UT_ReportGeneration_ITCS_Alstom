from cx_Freeze import setup, Executable
import sys
import os


EXE = '../src/UT_HTMLReportAnalysisAndReportPreparation_GUI'
filename = EXE+'.py'

includefiles = ['../src/UT_HTMLReportAnalysisAndReportPreparation.py']
packages = ['xlsxwriter']

base = None
if (sys.platform == "win32"):
    base = "Win32GUI"    # Tells the build script to hide the console.


setup(
    name = EXE ,
    version = "0.1" ,
    description = "first release" ,
    executables = [Executable(filename, base=base, icon="../Images/M.ico")],
    options = {'build_exe':{'packages':packages, 'include_files':includefiles}}
    )	
