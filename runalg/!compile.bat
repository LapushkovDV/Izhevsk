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

@set exedir=C:\Galaktika\AVAZ\chumachenko\6.0.09.0\exe

@%exedir%\vip64x.exe runalg.prj /r:chumachenko.res /c:vip.cfg /u:SUPERVISOR /#user:SUPERVISOR /def:Atl60

rem /c:support_res.cfg
@echo Компиляция завершена
@goto END

:END

pause
@exit

