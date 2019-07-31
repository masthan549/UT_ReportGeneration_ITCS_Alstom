@echo off

rem Copy required images from one folder to another folder
xcopy /Y ..\Images\alstom_logo.gif .\build\Images\
xcopy /Y ..\Images\M.ico .\build\Images\

rem wait for files to copy from one location to another
timeout 2

rem build the project
python.exe setup.py build