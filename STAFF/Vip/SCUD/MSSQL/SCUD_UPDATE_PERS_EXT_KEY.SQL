----------------------------------------------------
---  Обновим GP_SCUD_PERS.EXT_KEY
----------------------------------------------------
IF OBJECT_ID('tempdb ..#xx$locks')IS NOT NULL drop table #xx$locks
create table #xx$locks (TableNRec binary(8))


 update  GP_SCUD_PERS set GP_SCUD_PERS.F$EXT_KEY=cast(scud.ID_STAFF AS nVarChar)
  --select  SCUD.id_staff,SCUD.last_name+' '+SCUD.name+' '+SCUD.fathership NAME ,GP_SCUD_PERS.F$FIO

 FROM 
  OPENQUERY([SCUD],  'select
  st.last_name,st.first_name name, st.middle_name fathership
 ,st.id_staff from Staff st
 where st.last_name<>''''
 ') scud
 inner join T$GP_SCUD_PERS GP_SCUD_PERS on GP_SCUD_PERS.F$FIo = SCUD.last_name+' '+SCUD.name+' '+SCUD.fathership
    and GP_SCUD_PERS.F$EXT_KEY=''
 Left Outer join T$GP_SCUD_PERS on T$GP_SCUD_PERS.F$EXT_KEY =
  cast(scud.ID_STAFF AS nVarChar)
   where T$GP_SCUD_PERS.F$NREC is null

