create procedure GetPhotoFromPerco
as
begin
IF OBJECT_ID('erpdb.dbo.PhotoForGalFromPerco') IS NULL
     begin
          select SCUD.portret,  SCUD.id_staff, SCUD.last_name, SCUD.name, SCUD.fathership
          into erpdb.dbo.PhotoForGalFromPerco
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
         end
else
begin
 delete from erpdb.dbo.PhotoForGalFromPerco;
 insert into erpdb.dbo.PhotoForGalFromPerco (portret,  id_staff, last_name, name, fathership          )
          select SCUD.portret,  SCUD.id_staff, SCUD.last_name, SCUD.name, SCUD.fathership
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

end

declare CursIdStaff cursor for select id_staff from erpdb.dbo.PhotoForGalFromPerco

declare @ID_STAFF as bigint

open CursIdStaff

exec xp_cmdshell "del C:\SQLDB\ERPDB\Photo\Files\*.jpg /f /s /q"

fetch next from CursIdStaff into @ID_STAFF
WHILE (@@fetch_status <> -1)
begin
--print cast(@ID_STAFF as varchar(20))
-- основная фича в -f c:\Galaxy\ZxTime\Export_Photo\SQL\PhotoPromPerco.fmt
-- этот файл создали так: запустил в командной строке bcp "select t.portret.... и там ввел все параметрты (что это картинка, что длина префикса 0, остально по умолчанию) теперь этот файл используем как шаблон для выгрузкм
--declare @cmd nvarchar(1000) = 'bcp "select t.portret from erpdb.dbo.PhotoForGalFromPerco t where ID_STAFF =  '+cast(@ID_STAFF as varchar(20))+'" queryout "c:\temp\PhotoPerco\'+cast(@ID_STAFF as varchar(20))+'.jpg" -T -f c:\Galaxy\ZxTime\Export_Photo\SQL\PhotoPromPerco.fmt -S dep968-erpdb\erp'
declare @cmd nvarchar(1000) = 'bcp "select t.portret from erpdb.dbo.PhotoForGalFromPerco t where ID_STAFF =  '+cast(@ID_STAFF as varchar(20))+'" queryout "C:\SQLDB\ERPDB\Photo\Files\'+cast(@ID_STAFF as varchar(20))+'.jpg" -T -f "C:\SQLDB\ERPDB\Photo\PhotoPromPerco.fmt" -S dep968-erpdb\erp'
exec xp_cmdshell @cmd

 fetch next from CursIdStaff into @ID_STAFF
END


CLOSE CursIdStaff
DEALLOCATE CursIdStaff
end
-- exec xp_cmdshell "del C:\SQLDB\ERPDB\Photo\*.jpg /f /s /q"