-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters
-- command (Ctrl-Shift-M) to fill in the parameter
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================

IF OBJECT_ID('dbo.SCUD_UPDATE_PERS') IS NOT NULL
BEGIN
    REVOKE EXECUTE ON dbo.SCUD_UPDATE_PERS FROM public
    DROP PROCEDURE dbo.SCUD_UPDATE_PERS
    IF OBJECT_ID('dbo.SCUD_UPDATE_PERS') IS NOT NULL
        PRINT '<<< FAILED DROPPING PROCEDURE dbo.SCUD_UPDATE_PERS >>>'
    ELSE
        PRINT '<<< DROPPED PROCEDURE dbo.SCUD_UPDATE_PERS >>>'
END


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE SCUD_UPDATE_PERS
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
    SET DATEFORMAT ymd
declare @staff_max bigint;
select @staff_max = SC.id_staff From OPENQUERY([SCUD],'select max(id_staff) id_staff from staff') SC
PRINT ' staff_max   >>>'+Cast(@staff_max as nvarchar)
declare @fio_id_1 bigint
declare @fio_id_2 bigint

	
declare @str nvarchar(max) 
/*
set @str='Петров Петр Петрович'
set @fio_id_1=charindex(' ',@str,0)
set @fio_id_2=charindex(' ',@str,@fio_id_1+1)

Print  @str+', '+substring(@str,1,@fio_id_1-1)
           +', '+substring(@str,@fio_id_1+1,@fio_id_2-@fio_id_1)
           +', '+substring(@str,@fio_id_2+1,100)
*/
declare Curs_GP_SCUD_PERS cursor LOCAL FOR
(
  SELECT DISTINCT
     GP_P.F$NREC ID
    ,CAST(GP_P.F$EXT_KEY as integer ) P_ID_STAFF
	,GP_P.F$Fio FIO
	------
	,''  as TABEL_ID--   varchar(20), 3 значное число у части
    ,CASE WHEN COALESCE(Persons.F$AppDate,0)=0 THEN  GETDATE() 
	     ELSE dbo.to_SQLDATE(Persons.F$AppDate) 	END  P_DATE_BEGIN -- date,
    ,SC_PODR.ID_REF  P_ID_SUBDIV -- integer, ссылка на подразделение  SUBDIV_REF
    ------
    ,SC_APPOINT.ID_REF  P_ID_APPOINT -- integer, -- должность 
	,0 P_ID_GROUP_WT -- integer, ГРАФИК P_GRAPH_TYPE=0:GR_WORK_MAIN, 1 -EL_GR_WORK_MAIN(пустая) 
    ,0 P_ID_DOCUMENTS-- integer, 0 в STAFF_REF
    ------
    ,0 P_TEMPORARY-- integer, 1 у 10 человек
 	,cast(GETDATE() as timestamp) P_DATE_ACTION-- timestamp, Staff_ref
    ,0 P_DIN_PAY_SCHEMES_ID-- integer, staff - 0 всегда
    ------
    ,0 P_DIN_GRAPHS_ID-- integer, staff - 0 всегда
	,'' P_DOC_NUMBER  -- varchar(100),
    ,CASE WHEN COALESCE(Persons.F$BORNDATE,0)=0 THEN  null
	        ELSE dbo.to_SQLDATE(Persons.F$BORNDATE) 	END   P_BIRTHDAY_DATE --date,
    ------
    ,'' P_BIRTHDAY_PLACE-- varchar(255), STAFF_DOCS пусто у всех
	, Persons.F$SEX P_SEX-- varchar(6), STAFF_DOCS пусто у всех
    ,'' P_ISSUED_BY-- varchar(255), STAFF_DOCS пусто у всех, выдимо кем выдан
    ------
    ,CONVERT(date,'1899-12-30',120) P_ISSUED_DATE-- date, STAFF_DOCS дата вылачи документа? пусто
    ,CONVERT(date,'1899-12-30',120) P_EXPIRY_DATE-- date, STAFF_DOCS дата вылачи документа? пусто
    ,'' P_NATIONALITY --   varchar(100), STAFF_DOCS пусто у всех,
    ------
    ,'' P_DEPARTMENT_CODE --varchar(20), STAFF_DOCS пусто у всех,
    ,'' P_DOC_NAME -- varchar(100), STAFF_DOCS пусто у всех,
    ,0 P_GRAPH_TYPE -- integer, для P_ID_GROUP_WT 0 GR_WORK_MAIN
	--,'' P_PATH_ACTDIR_LOGIN --BIG_VARCHAR1024 = null,
    --,'' P_PATH_ACTDIR_DOMAIN -- VARCHAR_255 = null)
  FROM T$GP_SCUD_PERS GP_P 
  left outer join T$Persons Persons on Persons.F$Nrec = GP_P.F$cPersons
  left outer join T$Appointments App on App.F$Nrec = Persons.F$APPOINTCUR
  left outer join T$CATALOGS POST on Post.f$nrec = App.F$POST
  left outer join T$Katpodr Katpodr on  Katpodr.F$NREC = GP_P.F$CKATPODR
  left outer join T$GP_SCUD_PODR GP_PODR on GP_PODR.F$cKatpodr = GP_P.F$CKATPODR
                                or  GP_PODR.F$cKatorg = GP_P.F$cKatOrg
  left outer join OPENQUERY([SCUD],'select ID_REF,DISPLAY_NAME from SUBDIV_REF')  SC_PODR
		    on SC_PODR.DISPLAY_NAME = KATPODR.F$KOD
			or SC_PODR.ID_REf =GP_PODR.F$EXT_KEY
  left outer join OPENQUERY([SCUD],'select ID_REF,DISPLAY_NAME from APPOINT_REF')  SC_APPOINT
		    on SC_PODR.DISPLAY_NAME = POST.F$NAME
  where GP_P.F$EXT_KEY<>''
);

declare @P_LAST_NAME   varchar(100)    
declare @P_FIRST_NAME  varchar(100)
declare @P_MIDDLE_NAME varchar(100)
------------ запрос
declare @P_ID binary(8) 
declare @P_ID_STAFF integer
declare @P_FIO      nvarchar(100) -- фио
declare @P_TABEL_ID varchar(20)
declare @P_DATE_BEGIN date
declare @P_ID_SUBDIV integer
declare @P_ID_APPOINT integer
declare @P_ID_GROUP_WT integer
declare @P_ID_DOCUMENTS integer
declare @P_TEMPORARY integer
declare @P_DATE_ACTION timestamp
declare @P_DIN_PAY_SCHEMES_ID integer
declare @P_DIN_GRAPHS_ID integer
declare @P_DOC_NUMBER varchar(100)
declare @P_BIRTHDAY_DATE date
declare @P_BIRTHDAY_PLACE varchar(255)
declare @P_SEX varchar(6)
declare @P_ISSUED_BY varchar(255)
declare @P_ISSUED_DATE date
declare @P_EXPIRY_DATE date
declare @P_NATIONALITY varchar(100)
declare @P_DEPARTMENT_CODE varchar(20)
declare @P_DOC_NAME varchar(100)
declare @P_GRAPH_TYPE integer
-----------------
declare @QueryPerco nvarchar (max)
declare @QueryGal nvarchar (max)
-----------------
    
OPEN Curs_GP_SCUD_PERS
WHILE 1=1 
BEGIN 

  FETCH NEXT FROM Curs_GP_SCUD_PERS INTO 
        @P_ID             ,@P_ID_STAFF   ,@P_FIO -- фио
       ,@P_TABEL_ID       ,@P_DATE_BEGIN ,@P_ID_SUBDIV
	   ,@P_ID_APPOINT     ,@P_ID_GROUP_WT,@P_ID_DOCUMENTS
       ,@P_TEMPORARY      ,@P_DATE_ACTION,@P_DIN_PAY_SCHEMES_ID
       ,@P_DIN_GRAPHS_ID  ,@P_DOC_NUMBER ,@P_BIRTHDAY_DATE
       ,@P_BIRTHDAY_PLACE ,@P_SEX        ,@P_ISSUED_BY
       ,@P_ISSUED_DATE    ,@P_EXPIRY_DATE,@P_NATIONALITY
       ,@P_DEPARTMENT_CODE,@P_DOC_NAME   ,@P_GRAPH_TYPE
	   ;
  IF @@FETCH_STATUS<>0 
    BREAK 
  set @fio_id_1=charindex(' ',@P_FIO,0)
  set @fio_id_2=charindex(' ',@P_FIO,@fio_id_1+1)
  IF @fio_id_1>0 
   set @P_LAST_NAME   =substring(@P_FIO,1,@fio_id_1-1)
  else
   set @P_LAST_NAME=@P_FIO
  if @fio_id_2>0
  begin
    set @P_FIRST_NAME  =substring(@P_FIO,@fio_id_1+1,@fio_id_2-@fio_id_1)
    set @P_MIDDLE_NAME =substring(@P_FIO,@fio_id_2+1,100)
  end
  else
    set @P_FIRST_NAME  =substring(@P_FIO,@fio_id_1+1,100)
  set @staff_max=@staff_max+1
  set @QueryPerco  = 'STAFF_$UPD (
    '+cast(@P_ID_STAFF  as varchar) -- P_ID_STAFF integer,
    +','''+@P_LAST_NAME+'''' -- varchar(100),
    +','''+@P_FIRST_NAME+'''' -- varchar(100),
    +','''+@P_MIDDLE_NAME+'''' -- varchar(100),
    +','''+@P_TABEL_ID+'''' -- varchar(20),
    +','''+convert(varchar,@P_DATE_BEGIN,106)+'''' -- date, '10.10.2017'
    +','+cast(@P_ID_SUBDIV as varchar) -- integer,
    +','+cast(@P_ID_APPOINT as varchar) -- integer,
    +','+cast(@P_ID_GROUP_WT as varchar) -- integer,
    +','+cast(@P_ID_DOCUMENTS as varchar) -- integer,
    +','+cast(@P_TEMPORARY as varchar) -- integer,
    +','''+convert(varchar,@P_DATE_ACTION,106)+'''' -- timestamp,
    +','+cast(@P_DIN_PAY_SCHEMES_ID as varchar) -- integer,
    +','+cast(@P_DIN_GRAPHS_ID as varchar) -- integer,
    +','''+@P_DOC_NUMBER+'''' -- varchar(100),
    +','''+convert(varchar,@P_BIRTHDAY_DATE,106)+'''' -- date,
    +','''+@P_BIRTHDAY_PLACE+'''' -- varchar(255),
    +','''+@P_SEX+'''' -- varchar(6),
    +','''+@P_ISSUED_BY+'''' -- varchar(255),
    +','''+convert(varchar,@P_ISSUED_DATE,106)+'''' -- date,
    +','''+convert(varchar,@P_EXPIRY_DATE,106)+'''' -- date,
    +','''+@P_NATIONALITY+'''' -- varchar(100),
    +','''+@P_DEPARTMENT_CODE+'''' -- varchar(20),
    +','''+@P_DOC_NAME+'''' -- varchar(100),
    +','+cast(@P_GRAPH_TYPE as varchar) -- integer
  --  P_PATH_ACTDIR_LOGIN BIG_VARCHAR1024 = null,
  --  P_PATH_ACTDIR_DOMAIN VARCHAR_255 = null)
  +')'
  Exec(@QueryPerco)
  --update T$GP_SCUD_PERS set T$GP_SCUD_PERS.F$EXT_KEY=cast(@staff_max as nvarchar) 
  --    from T$GP_SCUD_PERS where T$GP_SCUD_PERS.F$NREC= @P_ID
  --PRINT 'ФИО='+@P_FIO+', установили EXT_KEY='+cast(@staff_max as nvarchar)
   -- END set @QueryPerco  =
END -- While 1=1

CLOSE Curs_GP_SCUD_PERS
DEALLOCATE Curs_GP_SCUD_PERS 

END ----------------- SCUD_UPDATE_PERS =


GO
IF OBJECT_ID('dbo.SCUD_UPDATE_PERS') IS NOT NULL
    PRINT '<<< CREATED PROCEDURE dbo.SCUD_UPDATE_PERS >>>'
ELSE
    PRINT '<<< FAILED CREATING PROCEDURE dbo.SCUD_UPDATE_PERS >>>'
GRANT EXECUTE ON dbo.SCUD_UPDATE_PERS TO public
go
