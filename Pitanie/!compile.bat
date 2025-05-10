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

@set exedir=C:\Galaktika\AVAZ\avaz_ptz\Exe_support

@%exedir%\vip.EXE pitanie.prj /r:pitanie.res /c:vip.cfg /u:supervisor /def:Atl60
@%exedir%\FRRES.EXE /to /r:pitanie.res /source:fr3

rem /c:support_res.cfg
@echo Компиляция завершена
@goto END

:END

pause
@exit

