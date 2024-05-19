USE [Russia2]
GO
/****** Object:  StoredProcedure [dbo].[S$BUFSPPLAM2]    Script Date: 19.05.2024 20:50:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER   PROCEDURE [dbo].[S$BUFSPPLAM2](@GUID VARCHAR(38),
                                       @cGRAFIK binary(8),
                                       @DateBeg_endraspr int ,
                                       @CMNPLAN binary(8),
                                       @CVIDPROD_not_need_raspr binary(8),
                                       @STATUS_PLAN binary(8)
                                      )

AS
BEGIN
--DECLARE @GUID VARCHAR(38) = '6546546546465'
--DECLARE @cGRAFIK binary(8) = 0x8000000000000001;
--DECLARE @DateBeg_endraspr int = dbo.toatldate('01.01.2024');
--DECLARE @CMNPLAN binary(8) = 0x8001000000000272;
--DECLARE @CVIDPROD_not_need_raspr binary(8) = 0x8000000000000001; -- Автомобили серия по ним не надо распределять количество по датам
--DECLARE @STATUS_PLAN binary(8) = 0x8000000000000024;

with PlanSbytSpecGroup as (
select
   mnplan.F$CROLEAN1 as roleVidpROD-- вид продукции
 , mnplan.F$WKODGR1  AS KODKAUVidProd-- вид продукции
 , mnplan.F$CANVAL1  as cVidprod-- вид продукции
 , spmnplan.f$CROLEAN1 as roleTechMarsh-- техмаршрут
 , spmnplan.F$WKODGR1  as kodkauTechMarsh-- техмаршрут
 , spmnplan.F$CANVAL1  as cTechMarch-- техмаршрут
-- , spmnpl.F$RECOMDATE  -- плановая дата
 , spmnplan.F$TYPEIZD as typeizd
 , spmnplan.F$CIZD   as cIzd
 , spmnplan.F$COTPED as cOtpEd
 , kated.F$NAME  as edName
 , case when kated.f$name  = 'шт' then 1 else kated.f$diskret end as diskret
 , coalesce(spsloj.F$CROLE,0x8000000000000000)  as roleModel
 , coalesce(spsloj.F$KODGRKAU,0) as kodkaumodel
 , case coalesce(spsloj.F$NPP,0) when 1 then hashan.F$CANALIT#1#
                                 when 2 then hashan.F$CANALIT#2#
                                 when 3 then hashan.F$CANALIT#3#
                                 when 4 then hashan.F$CANALIT#4#
                                 when 5 then hashan.F$CANALIT#5#
                                 when 6 then hashan.F$CANALIT#6#
                                 when 7 then hashan.F$CANALIT#7#
                                 when 8 then hashan.F$CANALIT#8#
                                 when 9 then hashan.F$CANALIT#9#
                                 when 10 then hashan.F$CANALIT#10#
        else 0x8000000000000000 end as сmodel
 , sum(valspmnp.F$KOL) as FullKol
 , case when spmnpl.F$RECOMDATE <= @DateBeg_endraspr and mnplan.F$CANVAL1 != @CVIDPROD_not_need_raspr -- по этому виду продукции на недо распределять по дням
                then [dbo].[ymd2date]([dbo].[dtYEAR]([dbo].[ToAtlDateTime2](spmnpl.F$RECOMDATE,0)),[dbo].[dtMONTH]([dbo].[ToAtlDateTime2](spmnpl.F$RECOMDATE,0)),1)
                else spmnpl.F$RECOMDATE end as PlanDate-- плановая дата
from t$mnplan mnplan
join T$SPMNPLAN spmnplan on spmnplan.F$CMNPLAN = mnplan.F$NREC
join T$SPMNPL spmnpl on spmnpl.F$CSPMNPLAN = spmnplan.F$NREC
join T$VALSPMNP valspmnp on valspmnp.F$CSPMNPL = spmnpl.F$NREC
join T$KATMC katmc on katmc.F$NREC = spmnplan.F$CIZD
left join T$KATOTPED katotped on katotped.F$NREC = spmnplan.F$COTPED
join T$KATED kated on kated.F$NREC = coalesce(katotped.F$CKATED, katmc.f$ced)
left join T$SLOJ sloj on sloj.F$NREC = katmc.F$CSLOJ
left join T$SPSLOJ SpSloj on spsloj.F$CVIEWSLOJ = sloj.F$NREC and spsloj.F$KODGRKAU = 20001
left join T$HASHAN hashan on hashan.F$NREC = katmc.F$CHASHAN

where mnplan.F$TYPEPLAN = 5
  and mnplan.F$CSTATUS = @STATUS_PLAN
  --and spmnpl.F$RECOMDATE >= @DateBeg_endraspr
  --and katmc.f$nrec = 0x80000000000233A9
  and valspmnp.F$KOL > 0

group by
   mnplan.F$CROLEAN1
 , mnplan.F$WKODGR1
 , mnplan.F$CANVAL1
 , spmnplan.f$CROLEAN1
 , spmnplan.F$WKODGR1
 , spmnplan.F$CANVAL1
-- , spmnpl.F$RECOMDATE  -- плановая дата
 , spmnplan.F$TYPEIZD
 , spmnplan.F$CIZD
 , spmnplan.F$COTPED
 , kated.F$NAME
 , case when kated.f$name  = 'шт' then 1 else kated.f$diskret end
 , coalesce(spsloj.F$CROLE,0x8000000000000000)
 , coalesce(spsloj.F$KODGRKAU,0)
 , case coalesce(spsloj.F$NPP,0) when 1 then hashan.F$CANALIT#1#
                                 when 2 then hashan.F$CANALIT#2#
                                 when 3 then hashan.F$CANALIT#3#
                                 when 4 then hashan.F$CANALIT#4#
                                 when 5 then hashan.F$CANALIT#5#
                                 when 6 then hashan.F$CANALIT#6#
                                 when 7 then hashan.F$CANALIT#7#
                                 when 8 then hashan.F$CANALIT#8#
                                 when 9 then hashan.F$CANALIT#9#
                                 when 10 then hashan.F$CANALIT#10#
        else 0x8000000000000000
  end
 , case when spmnpl.F$RECOMDATE <= @DateBeg_endraspr and mnplan.F$CANVAL1 != @CVIDPROD_not_need_raspr -- по этому виду продукции на недо распределять по дням
                then [dbo].[ymd2date]([dbo].[dtYEAR]([dbo].[ToAtlDateTime2](spmnpl.F$RECOMDATE,0)),[dbo].[dtMONTH]([dbo].[ToAtlDateTime2](spmnpl.F$RECOMDATE,0)),1)
                else spmnpl.F$RECOMDATE end
  ),
  PlanSbytSpecGroupRaspred as (
  select
   PlanSbytSpecGroup.roleVidpROD-- вид продукции
 , PlanSbytSpecGroup.KODKAUVidProd-- вид продукции
 , PlanSbytSpecGroup.cVidprod-- вид продукции
 , PlanSbytSpecGroup.roleTechMarsh-- техмаршрут
 , PlanSbytSpecGroup.kodkauTechMarsh-- техмаршрут
 , PlanSbytSpecGroup.cTechMarch-- техмаршрут
 , PlanSbytSpecGroup.typeizd
 , PlanSbytSpecGroup.cIzd
 , PlanSbytSpecGroup.cOtpEd
 , PlanSbytSpecGroup.edName
 , PlanSbytSpecGroup.roleModel
 , PlanSbytSpecGroup.kodkaumodel
 , PlanSbytSpecGroup.сmodel
 , PlanSbytSpecGroup.FullKol as RasprKol
 , PlanSbytSpecGroup.PlanDate-- плановая дата
  from PlanSbytSpecGroup
  where (PlanSbytSpecGroup.PlanDate > @DateBeg_endraspr and PlanSbytSpecGroup.cVidprod != @CVIDPROD_not_need_raspr)
       or (PlanSbytSpecGroup.cVidprod = @CVIDPROD_not_need_raspr)
   union all
  select
   PlanSbytSpecGroup.roleVidpROD-- вид продукции
 , PlanSbytSpecGroup.KODKAUVidProd-- вид продукции
 , PlanSbytSpecGroup.cVidprod-- вид продукции
 , PlanSbytSpecGroup.roleTechMarsh-- техмаршрут
 , PlanSbytSpecGroup.kodkauTechMarsh-- техмаршрут
 , PlanSbytSpecGroup.cTechMarch-- техмаршрут
 , PlanSbytSpecGroup.typeizd
 , PlanSbytSpecGroup.cIzd
 , PlanSbytSpecGroup.cOtpEd
 , PlanSbytSpecGroup.edName
 , PlanSbytSpecGroup.roleModel
 , PlanSbytSpecGroup.kodkaumodel
 , PlanSbytSpecGroup.сmodel
 , traspr.kol as RasprKol
 , traspr.dt as plandate
  from PlanSbytSpecGroup
  outer apply  dbo.fn_raspredKolToDays(PlanSbytSpecGroup.plandate,PlanSbytSpecGroup.diskret,@cGRAFIK,PlanSbytSpecGroup.FullKol) traspr
  where PlanSbytSpecGroup.PlanDate <= @DateBeg_endraspr and PlanSbytSpecGroup.cVidprod != @CVIDPROD_not_need_raspr
)

 INSERT INTO V$BUFSPPLAM2 (f$sRas, f$roleVidpROD, f$KODKAUVidProd, f$cVidprod, f$roleTechMarsh, f$kodkauTechMarsh, f$cTechMarch
                        , f$typeizd, f$cIzd, f$cOtpEd, f$edName, f$Rolemodel, f$kodkaumodel, f$cmodel, f$RasprKol, f$plandate)
   SELECT @GUID, roleVidpROD, KODKAUVidProd, cVidprod, roleTechMarsh, kodkauTechMarsh, cTechMarch
        , typeizd, cIzd, cOtpEd, edName, roleModel, kodkaumodel, сmodel, RasprKol, plandate
   FROM PlanSbytSpecGroupRaspred
END;

