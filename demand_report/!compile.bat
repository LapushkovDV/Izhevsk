rem это год-месяц-день YYYYMMDD для WINXP
@set DD=%DATE:~6,4%%DATE:~3,2%%DATE:~0,2%

del *.res
del *.tmp
del *.crf
del *.tmp
del *.dsk
del *.fnc
del *.log
del Atlantis*.res

@set exedir=C:\Galaktika\ATL6\Update_10.1.2230\support

@%exedir%\vip.EXE SNG_QRY.prj /r:HR_643.res /c:vip.cfg /u:lapus /p:Zraeqw123 /def:Atl60

rem /c:support_res.cfg
@echo Компиляция завершена
@goto END

:END

pause
@exit

