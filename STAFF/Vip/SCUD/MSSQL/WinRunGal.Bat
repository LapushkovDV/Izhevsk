rem *********************************************************
rem *  Автоматический запуск, copy stop                     *
rem *********************************************************
rem * Путь на Галактику *
set galnet="\\dep968-galsrv3\GalEMTrial\Exe\atlexec.exe"
%galnet% /c:\\dep968-galsrv3\GalEMTrial\Start\ERPDB.cfg %1 %2 /GALAXY.NOWRUN:GP_STAFF::GP_SCUD_SaldoCalc(0,0)

