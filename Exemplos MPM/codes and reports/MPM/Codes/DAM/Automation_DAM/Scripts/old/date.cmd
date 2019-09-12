for /f %%a in ('cscript //nologo yester.vbs') do set yesterday=%%a
print %yesterday%

pause