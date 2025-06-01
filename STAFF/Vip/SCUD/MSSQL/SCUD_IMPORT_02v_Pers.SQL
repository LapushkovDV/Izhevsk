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
--USE ERPDB
-- ================================================

IF OBJECT_ID('dbo.SCUD_IMPORT_PERS') IS NOT NULL
BEGIN
    REVOKE EXECUTE ON dbo.SCUD_IMPORT_PERS FROM public
    DROP PROCEDURE dbo.SCUD_IMPORT_PERS
    IF OBJECT_ID('dbo.SCUD_IMPORT_PERS') IS NOT NULL
        PRINT '<<< FAILED DROPPING PROCEDURE dbo.SCUD_IMPORT_PERS >>>'
    ELSE
        PRINT '<<< DROPPED PROCEDURE dbo.SCUD_IMPORT_PERS >>>'
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
CREATE PROCEDURE SCUD_IMPORT_PERS
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
    SET DATEFORMAT ymd
declare @delta_date nvarchar(10)
Set @delta_date='3'

declare @maxRegEv nvarchar(10)
select @maxRegEv = max(T$GP_SCUD_EVENTS.F$NPP) from T$GP_SCUD_EVENTS
declare @maxRegEv_All nvarchar(10) 
select @maxRegEv_All = max(T$GP_SCUD_EV_ALL.F$NPP) from T$GP_SCUD_EV_ALL
 
 --re.staff_id=10411060
declare @staff_id nvarchar(20) 
set @staff_id='10411060'
----------------------
---**** ƒл€ закачки**
----------------------
set @maxRegEv_all='0'
set @maxRegEv='0'
Set @delta_date='60'
---- **** ƒозагрузка
-- set @maxRegEv_all='38095330'
--set @maxRegEv='38095330'
--Set @delta_date='50'
----========================
   PRINT ' «агружено в GP_SCUD_EVENTS  >>>'+@maxRegEv
   PRINT ' «агружено в GP_SCUD_EV_ALL  >>>'+@maxRegEv_All

----------------------
--   TEST ------------  
-- SET @maxRegEv=37000000
-- SET @maxRegEv_All=37000000



--- контролеры замка - надо убирать выходы с ареала неконтролируемой территории
declare @Controler_Zamok nvarchar(max) 
declare @Controler_Zamok_IN nvarchar(max) 
set @Controler_Zamok_IN=
'((ip.ip_addr is Null) or (
       ip.ip_addr<>''''10.18.1.107''''
   and ip.ip_addr<>''''10.18.1.109''''
   and ip.ip_addr<>''''10.18.1.125''''
   and ip.ip_addr<>''''10.18.1.103'''' )
   )'

 set @Controler_Zamok= -- Ѕыло or - не правильно
 '((ip.ip_addr is Null)
 or 
  ( ip.ip_addr <>''''10.18.1.152'''' 
  and ip.ip_addr <>''''10.18.1.139''''
  and ip.ip_addr <>''''10.18.1.153''''
  and ip.ip_addr <>''''10.18.1.103''''
 ))' 
 
----------------------------------------------------------------
--  адры  ареалы 1318780 на вход контролер ID 1318770 выход 1318774
----------------------------------------------------------
   PRINT ' Ётап 01 ¬ходы'

declare @QueryPerco nvarchar (max)
set @QueryPerco  = 'insert into V$GP_SCUD_EVENTS (
 F$NPP,F$sNPP,F$STRTABN,F$NAME
,F$CGP_SCUD_PERS,F$ENTERDATE,F$ENTERTIME
,F$CARD,F$CKATPODR,F$WEVENTS_TYPE,F$WRESOURCE
,F$CSTATUS,F$CCONTROLER,F$CAREA
)
SELECT
 SCUD_EV.Id_reg NPP
,CAST(SCUD_EV.Id_reg as nvarchar) sNPP
,T$GP_SCUD_PERS.F$CODE STRTABN, SCUD_EV.last_name+'' ''+SCUD_EV.name+'' ''+SCUD_EV.fathership NAME
,T$GP_SCUD_PERS.F$Nrec CGP_SCUD_PERS,dbo.ToAtlDate(SCUD_EV.date_ev) ENTERDATE,dbo.ToAtlTime(SCUD_EV.time_ev) ENTERTIME
,SCUD_EV.Indetifier CARD,T$GP_SCUD_PERS.F$ckatpodr CKATPODR
, 0 WEVENTS_TYPE
,case WHEN (SCUD_EV.areas_id =125937 or SCUD_EV.areas_id =1318780) THEN 1 else 0 end WRESOURCE
,T$GP_STATUS.F$NREC CSTATUS
,T$GP_SCUD_CONTROLER.F$NREC CCONTROLER
,T$GP_SCUD_AREAS.F$NREC CAREA
FROM OPENQUERY([SCUD],  ''select re.time_ev as time_ev 
 ,min(dpt.display_name) dcode,min(re.id_reg)  Id_reg, 
 max(CASE when 1318629=ct.id_configs_tree then 1318770 else ct.id_configs_tree end) AS id_configs_tree
 ,min(re.type_pass) type_pass, min(re.inner_number_ev) inner_number_ev
 ,max(sc.Identifier) as Indetifier,re.date_ev, re.areas_id 
 --,re_o2.time_ev time_ev2
 ,st.last_name,st.first_name name, st.middle_name fathership, st.id_staff
 from REG_EVENTS re  inner join staff st on st.id_staff = re.staff_id
 inner join AREAS_TREE ar  on (ar.ID_AREAS_TREE=re.AREAS_ID)
 left outer join STAFF_REF sr on sr.staff_id=st.id_staff
 left outer join STAFF_CARDS sc on sc.staff_id=st.id_staff
 and sc.DATE_END>CURRENT_DATE-10
 left outer join SUBDIV_REF dpt on dpt.id_ref=sr.subdiv_id
 left outer join CONFIGS_TREE ct on ct.id_configs_tree = re.configs_tree_id_controller
 Left outer join CONFIGS_TREE_LINKS cl on cl.configs_tree_id_child =ct.id_configs_tree 
 left outer join CONFIGS_TREE_IP ip on ip.configs_tree_id = cl.configs_tree_id_parent 
 where re.date_ev>=CURRENT_DATE-'+@delta_date+' and re.id_reg > ' +@maxRegEv+'
 and re.type_pass=1   and (re.areas_id =125937 or re.areas_id =1318780) 
 and re.staff_id='+@staff_id+'
  and '+@Controler_Zamok_IN+'
  --and SubString(ct.display_name from 1 for 2)<>'''' Ѕ'''' -- не будем брать входы с  Ѕ
  group by re.date_ev,re.areas_id ,re.time_ev -- было re_o2.time_ev
,st.last_name ,st.first_name, st.middle_name, st.id_staff
 '') scud_ev
 inner join T$GP_TYPEDOCS on T$GP_TYPEDOCS.F$WTYPE =1009
 inner join  T$GP_STATUS on T$GP_STATUS.F$CTYPEDOC =T$GP_TYPEDOCS.F$NREC
       and T$GP_STATUS.F$NAME=''— ”ƒ''
 Inner JOIN T$GP_SCUD_PERS ON T$GP_SCUD_PERS.F$EXT_KEY = cast(scud_ev.ID_STAFF AS nVarChar)
 inner JOIN T$GP_SCUD_AREAS ON T$GP_SCUD_AREAS.F$ID_AREAS_TREE=scud_ev.areas_id
 inner join T$GP_SCUD_CONTROLER ON T$GP_SCUD_CONTROLER.F$EXT_KEY = cast(scud_ev.id_configs_tree AS nVarChar)
     --and SubString(T$GP_SCUD_CONTROLER.F$name, 1, 2)<>'' Ѕ'' -- не будем брать входы с  Ѕ
	 and T$GP_SCUD_CONTROLER.F$wType<>1
 LEFT OUTER JOIN T$GP_SCUD_EVENTS ON T$GP_SCUD_EVENTS.F$CGP_SCUD_PERS =T$GP_SCUD_PERS.F$NREC
    AND T$GP_SCUD_EVENTS.F$ENTERDATE=dbo.ToAtlDate(SCUD_EV.date_ev)--F$ENTERDATE            --≥Д†в† ҐеЃ§†                       ≥Date
    AND T$GP_SCUD_EVENTS.F$ENTERTIME =dbo.ToAtlTime(SCUD_EV.time_ev)--F$ENTERTIME            --≥Fhtvd ҐлеЃ§†                     ≥Time
  where T$GP_SCUD_EVENTS.F$NREC is NULL
  ' -- END set @QueryPerco  =

--PRINT @QueryPerco
Exec(@QueryPerco)
PRINT ' Ётап 02 ¬ыходы'


set @QueryPerco  = 'insert into V$GP_SCUD_EVENTS (
 F$NPP,F$sNPP,F$STRTABN,F$NAME
,F$CGP_SCUD_PERS,F$ENTERDATE,F$ENTERTIME
,F$CARD,F$CKATPODR,F$WEVENTS_TYPE,F$WRESOURCE
,F$CSTATUS,F$CCONTROLER,F$CAREA
)
SELECT
 SCUD_EV.Id_reg NPP
,CAST(SCUD_EV.Id_reg as nvarchar) sNPP
,T$GP_SCUD_PERS.F$CODE STRTABN, SCUD_EV.last_name+'' ''+SCUD_EV.name+'' ''+SCUD_EV.fathership NAME
,T$GP_SCUD_PERS.F$Nrec CGP_SCUD_PERS,dbo.ToAtlDate(SCUD_EV.date_ev) ENTERDATE,dbo.ToAtlTime(SCUD_EV.time_ev) ENTERTIME
,SCUD_EV.Indetifier CARD,T$GP_SCUD_PERS.F$ckatpodr CKATPODR, 0 WEVENTS_TYPE
,case WHEN (SCUD_EV.areas_id =125937  or (SCUD_EV.areas_id =1318780 and SCUD_EV.id_configs_tree=1318770)) THEN 1 else 0 end WRESOURCE
,T$GP_STATUS.F$NREC CSTATUS
,T$GP_SCUD_CONTROLER.F$NREC CCONTROLER
,T$GP_SCUD_AREAS.F$NREC CAREA
FROM OPENQUERY([SCUD],  ''
 select distinct  max(re.time_ev) AS TIME_EV
 ,max(dpt.display_name) dcode,max(re.id_reg) AS ID_REG
 ,max(CASE when 1318629=ct.id_configs_tree then 1318774 else ct.id_configs_tree end) AS id_configs_tree
 ,max(re.type_pass) AS type_pass
 ,max(re.inner_number_ev) AS inner_number_ev
 ,max(sc.Identifier) as Indetifier
 ,re.date_ev,re.areas_id ,re.time_ev time_ev2
 ,st.last_name,st.first_name name, st.middle_name fathership, st.id_staff
from REG_EVENTS re
inner join staff st on st.id_staff=re.staff_id
inner join AREAS_TREE ar on (ar.ID_AREAS_TREE=re.AREAS_ID)
left outer join STAFF_REF sr on sr.staff_id=st.id_staff
left outer join STAFF_CARDS sc on sc.staff_id=st.id_staff
and sc.DATE_END>CURRENT_DATE-10
left outer join SUBDIV_REF dpt on dpt.id_ref=sr.subdiv_id
left outer join CONFIGS_TREE ct on ct.id_configs_tree=re.configs_tree_id_controller
Left outer join CONFIGS_TREE_LINKS cl on cl.configs_tree_id_child =ct.id_configs_tree 
left outer join CONFIGS_TREE_IP ip on ip.configs_tree_id = cl.configs_tree_id_parent 
where re.date_ev>=CURRENT_DATE-'+@delta_date+' and re.id_reg > ' +@maxRegEv+'
 and re.staff_id='+@staff_id+'
 and re.type_pass=1 and (re.areas_id =1 )
 and '+@Controler_Zamok+'
 group by re.date_ev,re.areas_id ,re.time_ev
 ,st.last_name,st.first_name , st.middle_name, st.id_staff
 '') scud_ev
 inner join T$GP_TYPEDOCS on T$GP_TYPEDOCS.F$WTYPE =1009
 inner join T$GP_STATUS on T$GP_STATUS.F$CTYPEDOC =T$GP_TYPEDOCS.F$NREC
       and T$GP_STATUS.F$NAME=''— ”ƒ''
 Inner JOIN T$GP_SCUD_PERS ON T$GP_SCUD_PERS.F$EXT_KEY = cast(scud_ev.ID_STAFF AS nVarChar)
 inner JOIN T$GP_SCUD_AREAS ON T$GP_SCUD_AREAS.F$ID_AREAS_TREE=scud_ev.areas_id
 inner join T$GP_SCUD_CONTROLER ON T$GP_SCUD_CONTROLER.F$EXT_KEY = cast(scud_ev.id_configs_tree AS nVarChar)
     and T$GP_SCUD_CONTROLER.F$wType<>1
 LEFT OUTER JOIN T$GP_SCUD_EVENTS ON T$GP_SCUD_EVENTS.F$CGP_SCUD_PERS =T$GP_SCUD_PERS.F$NREC
    AND T$GP_SCUD_EVENTS.F$ENTERDATE=dbo.ToAtlDate(SCUD_EV.date_ev)--F$ENTERDATE            --≥Д†в† ҐеЃ§†                       ≥Date
    AND T$GP_SCUD_EVENTS.F$ENTERTIME =dbo.ToAtlDate(SCUD_EV.time_ev)--F$ENTERTIME            --≥Fhtvd ҐлеЃ§†                     ≥Time
  where T$GP_SCUD_EVENTS.F$NREC is NULL
  ' -- END set @QueryPerco  =

Exec(@QueryPerco)
PRINT ' Ётап 03 ¬се событи€'

--************************************************
declare @QueryPerco_ALL nvarchar (max)
set @QueryPerco_ALL ='
insert into V$GP_SCUD_EV_ALL(
 F$NPP,F$sNPP,F$STRTABN,F$NAME,F$CGP_SCUD_PERS
,F$ENTERDATE,F$ENTERTIME,F$CARD
,F$CKATPODR,F$WEVENTS_TYPE,F$WRESOURCE
,F$CSTATUS,F$CCONTROLER,F$CAREA
)

SELECT DISTINCT
 SCUD_EV.Id_reg NPP                  --≥ПЃап§™ЃҐл© НЃђ•а                 ≥LongInt
,CAST(SCUD_EV.Id_reg as nvarchar(50)) sNPP
,T$GP_SCUD_PERS.F$CODE STRTABN              --≥ПЃап§™ЃҐл© НЃђ•а                 ≥String
,SCUD_EV.last_name+'' ''+SCUD_EV.name+'' ''+SCUD_EV.fathership NAME                 --≥СЃваг§≠®™ ѓЃб•в®в•Ђм             ≥String
,T$GP_SCUD_PERS.F$Nrec CGP_SCUD_PERS        --≥СблЂ™† ≠† бѓа†ҐЃз≠®™ Ђо§•©       ≥Comp
,dbo.ToAtlDate(SCUD_EV.date_ev) ENTERDATE            --≥Д†в† ҐеЃ§†                       ≥Date
,dbo.ToAtlTime(SCUD_EV.time_ev) ENTERTIME            --≥Fhtvd ҐлеЃ§†                     ≥Time
,SCUD_EV.Indetifier CARD                 --≥К†ав†                            ≥String
,T$GP_SCUD_PERS.F$ckatpodr CKATPODR             --≥ПЃ§а†І§•Ђ•≠®•                    ≥Comp
,SCUD_EV.type_pass WEVENTS_TYPE         --≥Т®ѓ бЃ°лв®п                      ≥Word
,case WHEN SCUD_EV.areas_id =125937  or (SCUD_EV.areas_id =1318780 and SCUD_EV.id_configs_tree=1318770) THEN 1 else 0 end WRESOURCE            --≥Т®ѓ а•бгаб† гбваЃ©бвҐ†           ≥Word
,T$GP_STATUS.F$NREC CSTATUS              --≥Св†вгб                           ≥Comp
,T$GP_SCUD_CONTROLER.F$NREC CCONTROLER           --≥КЃ≠ваЃЂ•а                        ≥Comp
,T$GP_SCUD_AREAS.F$NREC CAREA                --≥Аа•†Ђ                            ≥Comp
FROM OPENQUERY([SCUD],  ''
         select re.time_ev as time_ev --Ґа•ђп ҐлеЃ§†
       , dpt.display_name dcode
       , re.id_reg  Id_reg
       , CASE when 1318629=ct.id_configs_tree 
	     then  CASE when re.areas_id=1 then 1318774 else 1318770 end
	     else ct.id_configs_tree end AS id_configs_tree
       , re.type_pass type_pass
       , re.inner_number_ev inner_number_ev
       , sc.Identifier as Indetifier
       , re.date_ev, re.areas_id--,re_o2.time_ev time_ev2
       ,st.last_name,st.first_name name, st.middle_name fathership, st.id_staff
  from REG_EVENTS re  inner join staff st on st.id_staff = re.staff_id
    inner join AREAS_TREE ar  on (ar.ID_AREAS_TREE = re.AREAS_ID)
    left outer join STAFF_REF sr on sr.staff_id        = st.id_staff
    left outer join STAFF_CARDS sc on sc.staff_id        = st.id_staff
               and sc.DATE_END>CURRENT_DATE-10
    left outer join SUBDIV_REF dpt on dpt.id_ref = sr.subdiv_id
    left outer join CONFIGS_TREE ct on ct.id_configs_tree = re.configs_tree_id_controller

  where re.date_ev>=CURRENT_DATE-'+@delta_date+'  and re.id_reg > ' + Cast(@maxRegEv_All as nvarchar)+'
  and re.staff_id='+@staff_id+'
 
	'') AS scud_ev
 inner join T$GP_TYPEDOCS on T$GP_TYPEDOCS.F$WTYPE =1009
 inner join  T$GP_STATUS on T$GP_STATUS.F$CTYPEDOC =T$GP_TYPEDOCS.F$NREC
       and T$GP_STATUS.F$NAME=''— ”ƒ''
 Inner JOIN T$GP_SCUD_PERS ON T$GP_SCUD_PERS.F$EXT_KEY = cast(scud_ev.ID_STAFF AS nVarChar)
 inner JOIN T$GP_SCUD_AREAS ON T$GP_SCUD_AREAS.F$ID_AREAS_TREE=scud_ev.areas_id
 inner join T$GP_SCUD_CONTROLER ON T$GP_SCUD_CONTROLER.F$EXT_KEY = cast(scud_ev.id_configs_tree AS nVarChar)
 LEFT OUTER JOIN T$GP_SCUD_EV_ALL ON T$GP_SCUD_EV_ALL.F$CGP_SCUD_PERS =T$GP_SCUD_PERS.F$NREC
    AND CAST(SCUD_EV.Id_reg as nvarchar(50))=T$GP_SCUD_EV_ALL.F$SNPP--F$ENTERDATE            --≥Д†в† ҐеЃ§†                       ≥Date

 where T$GP_SCUD_EV_ALL.F$NREC is NULL
 ' --

Exec(@QueryPerco_ALL)

END ----------------- @QueryPerco_ALL =


GO
IF OBJECT_ID('dbo.SCUD_IMPORT_PERS') IS NOT NULL
    PRINT '<<< CREATED PROCEDURE dbo.SCUD_IMPORT_PERS >>>'
ELSE
    PRINT '<<< FAILED CREATING PROCEDURE dbo.SCUD_IMPORT_PERS >>>'
GRANT EXECUTE ON dbo.SCUD_IMPORT_PERS TO public
go
