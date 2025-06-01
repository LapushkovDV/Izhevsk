-- To allow advanced options to be changed.
EXEC sp_configure 'show advanced options', 1
GO
-- To update the currently configured value for advanced options.
RECONFIGURE
GO
-- To enable the feature.
EXEC sp_configure 'xp_cmdshell', 1
GO
-- To update the currently configured value for this feature.
RECONFIGURE
GO

USE ERPDB_OLD
-- удали таблицу с фото
drop table PhotoForGalFromPerco
--заполним таблицу с фото
select SCUD.portret,  SCUD.id_staff, SCUD.last_name, SCUD.name, SCUD.fathership 
into PhotoForGalFromPerco 
from
OPENQUERY([SCUD],  'select
st.portret,  st.last_name,st.first_name name, st.middle_name fathership
 ,st.id_staff from Staff st
 where st.portret is not null
') scud
inner join T$GP_SCUD_PERS on T$GP_SCUD_PERS.F$EXT_KEY=cast(scud.ID_STAFF AS nVarChar)
left outer join T$APPENDIX Appendix on 1 = Appendix.F$ObjBlock
                              and T$GP_SCUD_PERS.F$cPersons = Appendix.F$Person
                              and 200 = Appendix.F$ObjType
where (Appendix.F$Nrec is Null) or Appendix.F$CONTENTS=0
--пройдем по ней и выгрузим
declare CursIdStaff cursor for select id_staff from erpdb_old.dbo.PhotoForGalFromPerco

declare @ID_STAFF as bigint

open CursIdStaff

fetch next from CursIdStaff into @ID_STAFF
WHILE (@@fetch_status <> -1)
begin
--print cast(@ID_STAFF as varchar(20))
-- основная фича в -f c:\Galaxy\ZxTime\Export_Photo\SQL\PhotoPromPerco.fmt
-- этот файл создали так: запустил в командной строке bcp "select t.portret.... и там ввел все параметрты (что это картинка, что длина префикса 0, остально по умолчанию) теперь этот файл используем как шаблон для выгрузкм
--declare @cmd nvarchar(1000) = 'bcp "select t.portret from erpdb.dbo.PhotoForGalFromPerco t where ID_STAFF =  '+cast(@ID_STAFF as varchar(20))+'" queryout "c:\temp\PhotoPerco\'+cast(@ID_STAFF as varchar(20))+'.jpg" -T -f c:\Galaxy\ZxTime\Export_Photo\SQL\PhotoPromPerco.fmt -S dep968-erpdb\erp'
declare @cmd nvarchar(1000) = 'bcp "select t.portret from erpdb_old.dbo.PhotoForGalFromPerco t where ID_STAFF =  '+cast(@ID_STAFF as varchar(20))+'" queryout "c:\Galaxy\ZxTime\Export_Photo\SQL\Fhoto\'+cast(@ID_STAFF as varchar(20))+'.jpg" -T -f "c:\Galaxy\ZxTime\Export_Photo\SQL\garsvr2014.fmt" -S dep968-galsrv\erp'
exec xp_cmdshell @cmd

 fetch next from CursIdStaff into @ID_STAFF
END


CLOSE CursIdStaff
DEALLOCATE CursIdStaff



