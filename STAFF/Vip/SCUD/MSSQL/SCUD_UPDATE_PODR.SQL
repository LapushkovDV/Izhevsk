IF OBJECT_ID('tempdb ..#xx$locks')IS NOT NULL drop table #xx$locks
create table #xx$locks (TableNRec binary(8))

update T$GP_SCUD_PODR set T$GP_SCUD_PODR.F$EXT_KEY = cast(SC_PODR.ID_REF as varchar)
--select katpodr.F$NREC ,KATPODR.F$KOD,SC_PODR.DISPLAY_NAME
from T$GP_SCUD_PODR 
inner join T$KATPODR KATPODR on KATPODR.F$NREC= T$GP_SCUD_PODR.F$CKATPODR
--inner 
inner join OPENQUERY([SCUD],'select ID_REF,DISPLAY_NAME from SUBDIV_REF')  SC_PODR
		    on rtrim(SC_PODR.DISPLAY_NAME) = KATPODR.F$KOD



update T$GP_SCUD_PODR set T$GP_SCUD_PODR.F$EXT_KEY_UP = cast(SC_PODR.ID_REF as varchar)
--select katpodr.F$NREC ,KATPODR.F$KOD,SC_PODR.DISPLAY_NAME
from T$GP_SCUD_PODR 
inner join T$KATPODR KATPODR on KATPODR.F$NREC= T$GP_SCUD_PODR.F$CKATPODR
inner join OPENQUERY([SCUD],'select ID_REF,DISPLAY_NAME from SUBDIV_REF')  SC_PODR
		    on rtrim(subString(SC_PODR.DISPLAY_NAME,1,3)) = SubString(KATPODR.F$KOD,1,3)
