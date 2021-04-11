pyinstaller PyOffice.py
copy "mains.ini" "%tmp%\123.ini"
cd buildres
copy "serf.vbs" "%tmp%\1234.vbs"
cd..
cd dist
cd PyOffice
copy "%tmp%\123.ini" "mains.ini"
cd..
cd..
del build /s /q
del __pycache__ /s /q
del PyOffice.spec
cd build 
rmdir PyOffice
cd..
rmdir build
rmdir __pycache__
md %userprofile%\Desktop\PyOfficeBuild
cd dist
xcopy "PyOffice" "%userprofile%\Desktop\PyOfficeBuild" /E
cd..
del dist /s /q 
rmdir dist
cd %userprofile%\Desktop\PyOfficeBuild
copy "%tmp%\1234.vbs" "runpyoffice.vbs"