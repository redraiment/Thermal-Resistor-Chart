@echo off

::����
lib\CH341SER.EXE

::���������Ŀؼ�
copy /Y lib\COMDLG32.OCX %windir%\system32\
copy /Y lib\MSCOMM32.OCX %windir%\system32\
mkdir %ProgramFiles%\Dashboard
copy dist\�����¿�.exe %ProgramFiles%\Dashboard

::ע��ؼ�
cd %windir%\system32
regsvr32 /s commdlg32.ocx
regsvr32 /s mscomm32.ocx

::�����Ȩ
reg add HKCR\Licenses\4D553650-6ABE-11cf-8ADB-00AA00C00905 /ve /d gfjmrfkfifkmkfffrlmmgmhmnlulkmfmqkqj /f
reg add HKCR\Licenses\4250E830-6AC2-11cf-8ADB-00AA00C00905 /ve /d kjljvjjjoquqmjjjvpqqkqmqykypoqjquoun /f

::��ݷ�ʽ
cscript link.vbs
