@echo off

::驱动
lib\CH341SER.EXE

::复制依赖的控件
copy /Y lib\COMDLG32.OCX %windir%\system32\
copy /Y lib\MSCOMM32.OCX %windir%\system32\
mkdir %ProgramFiles%\Dashboard
copy dist\电子温控.exe %ProgramFiles%\Dashboard

::注册控件
cd %windir%\system32
regsvr32 /s commdlg32.ocx
regsvr32 /s mscomm32.ocx

::添加授权
reg add HKCR\Licenses\4D553650-6ABE-11cf-8ADB-00AA00C00905 /ve /d gfjmrfkfifkmkfffrlmmgmhmnlulkmfmqkqj /f
reg add HKCR\Licenses\4250E830-6AC2-11cf-8ADB-00AA00C00905 /ve /d kjljvjjjoquqmjjjvpqqkqmqykypoqjquoun /f

::快捷方式
cscript link.vbs
