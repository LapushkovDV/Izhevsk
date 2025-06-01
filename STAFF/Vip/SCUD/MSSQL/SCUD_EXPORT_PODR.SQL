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

IF OBJECT_ID('dbo.SCUD_EXPORT_PODR') IS NOT NULL
BEGIN
    REVOKE EXECUTE ON dbo.SCUD_EXPORT_PODR FROM public
    DROP PROCEDURE dbo.SCUD_EXPORT_PODR
    IF OBJECT_ID('dbo.SCUD_EXPORT_PODR') IS NOT NULL
        PRINT '<<< FAILED DROPPING PROCEDURE dbo.SCUD_EXPORT_PODR >>>'
    ELSE
        PRINT '<<< DROPPED PROCEDURE dbo.SCUD_EXPORT_PODR >>>'
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
CREATE PROCEDURE SCUD_EXPORT_PODR
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
    SET DATEFORMAT ymd
declare @subdiv_ref bigint;
select @subdiv_ref = SC.id_staff From OPENQUERY([SCUD],'select max(id_staff) id_staff from staff') SC
PRINT ' subdiv_ref   >>>'+Cast(@subdiv_ref as nvarchar)

	
declare @str nvarchar(max) 
/*
set @str='Петров Петр Петрович'
set @fio_id_1=charindex(' ',@str,0)
set @fio_id_2=charindex(' ',@str,@fio_id_1+1)

Print  @str+', '+substring(@str,1,@fio_id_1-1)
           +', '+substring(@str,@fio_id_1+1,@fio_id_2-@fio_id_1)
           +', '+substring(@str,@fio_id_2+1,100)
*/
declare Curs_GP_SCUD_PODR cursor LOCAL FOR
(
  SELECT DISTINCT
       PODR.F$NREC
      ,CAST(PODR_UP.F$EXT_KEY as integer) -- 256674 P_ID_PARENT вышестоящее или текущего уровня
      ,0    -- P_ID_REF    новый ID, надо где то получать его ранее ...
      ,CASE when PODR.F$WTYPE =0 then  PODR.F$CODE ELSE PODR.F$NAME END      -- P_DISPLAY_NAME наименование
      ,1623000    -- P_USER_ID      под кем создаем, это Аретм Синдецький
      ,1          -- P_VISIBLE      1 - видимый
      ,0          -- P_SUBDIV_OFFICIES         по аналогии с другими записями в PERCo
      ,0          -- P_STAFF_ID_CONTANT_PERSON по аналогии с другими записями в PERCo
      ,0          -- P_STAFF_ID_ACCOMPANY      по аналогии с другими записями в PERCo
      ,''         -- P_DESCRIPTION             по аналогии с другими записями в PERCo, ну так то можно и название вставлять
      ,''         -- P_PHONES                  по аналогии с другими записями в PERCo, ну можно реальный телефон из Галки тащить
      ,0          -- P_STAFF_ACCESS_PATTERN_ID по аналогии с другими записями в PERCo
      ,0          -- P_GUEST_ACCESS_PATTERN_ID по аналогии с другими записями в PERCo
      ,1          -- P_TYPE_INS  0- подчиненное, 1 - на текущий уровеньFrom T$
	  FROM T$GP_SCUD_PODR PODR
	  inner join T$GP_SCUD_PODR PODR_UP on PODR_UP.F$NREC = PODR.F$CNODE
);

------------ запрос
declare @P_ID binary(8) 
declare @P_ID_PARENT integer
declare @P_ID_REF integer
declare @P_DISPLAY_NAME varchar(100)
declare @P_USER_ID integer
declare @P_VISIBLE integer
declare @P_SUBDIV_OFFICIES varchar(1600)
declare @P_STAFF_ID_CONTANT_PERSON integer
declare @P_STAFF_ID_ACCOMPANY integer
declare @P_DESCRIPTION varchar(1600)
declare @P_PHONES varchar(70)
declare @P_STAFF_ACCESS_PATTERN_ID integer
declare @P_GUEST_ACCESS_PATTERN_ID integer
declare @P_TYPE_INS integer

-----------------
declare @QueryPerco nvarchar (max)
-----------------
    
OPEN Curs_GP_SCUD_PODR
WHILE 1=1 
BEGIN 

  FETCH NEXT FROM Curs_GP_SCUD_PODR INTO 
        @P_ID             , @P_ID_PARENT, @P_ID_REF
       ,@P_DISPLAY_NAME   , @P_USER_ID,@P_VISIBLE
       ,@P_SUBDIV_OFFICIES, @P_STAFF_ID_CONTANT_PERSON,@P_STAFF_ID_ACCOMPANY
       ,@P_DESCRIPTION    , @P_PHONES   ,@P_STAFF_ACCESS_PATTERN_ID
       ,@P_GUEST_ACCESS_PATTERN_ID       ,@P_TYPE_INS
	   ;
  IF @@FETCH_STATUS<>0 
    BREAK 
  set @subdiv_ref=@subdiv_ref+1
  set @QueryPerco  = 'STAFF_$INS ( '
    +  cast(@P_ID_PARENT as varchar) -- integer,
    +','+cast(@P_ID_REF as varchar) -- integer,
    +','''+@P_DISPLAY_NAME+'''' -- varchar(100),
    +','+cast(@P_USER_ID as varchar) -- integer,
    +','+cast(@P_VISIBLE as varchar) -- integer,
    +','''+@P_SUBDIV_OFFICIES+'''' -- varchar(1600),
    +','+cast(@P_STAFF_ID_CONTANT_PERSON as varchar) -- integer,
    +','+cast(@P_STAFF_ID_ACCOMPANY as varchar) -- integer,
    +','''+@P_DESCRIPTION+'''' -- varchar(1600),
    +','''+@P_PHONES+'''' -- varchar(70),
    +','+cast(@P_STAFF_ACCESS_PATTERN_ID as varchar) -- integer,
    +','+cast(@P_GUEST_ACCESS_PATTERN_ID as varchar) -- integer,
    +','+cast(@P_TYPE_INS as varchar) -- integer
	+')'
  Exec(@QueryPerco)
  update T$GP_SCUD_PODR set T$GP_SCUD_PODR.F$EXT_KEY=cast(@subdiv_ref as nvarchar) 
      from T$GP_SCUD_PODR where T$GP_SCUD_PODR.F$NREC= @P_ID
  PRINT 'PODR '+@P_DISPLAY_NAME+', установили EXT_KEY='+cast(@subdiv_ref as nvarchar)
   -- END set @QueryPerco  =
END -- While 1=1

CLOSE Curs_GP_SCUD_PODR
DEALLOCATE Curs_GP_SCUD_PODR

END ----------------- SCUD_EXPORT_PODR =


GO
IF OBJECT_ID('dbo.SCUD_EXPORT_PODR') IS NOT NULL
    PRINT '<<< CREATED PROCEDURE dbo.SCUD_EXPORT_PODR >>>'
ELSE
    PRINT '<<< FAILED CREATING PROCEDURE dbo.SCUD_EXPORT_PODR >>>'
GRANT EXECUTE ON dbo.SCUD_EXPORT_PODR TO public
go
