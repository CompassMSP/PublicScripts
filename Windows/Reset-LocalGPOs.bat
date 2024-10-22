@echo off
del /s "%windir%\system32\GroupPolicy" /F /Q
del /S "C:\ProgramData\Microsoft\Group Policy\History" /F /Q
REG DELETE "HKLM\Software\Policies\Microsoft" /F

gpupdate /force