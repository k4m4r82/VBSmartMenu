cls
echo. install komponen VBSmartMenu XP
pause
copy SmartMenuXP.dll %systemroot%\system32
copy SmartMenuXP.ocx %systemroot%\system32

regsvr32 /s %systemroot%\system32\SmartMenuXP.dll
regsvr32 /s %systemroot%\system32\SmartMenuXP.ocx