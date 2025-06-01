USE [ERPDB]
GO

/****** Object:  View [dbo].[S$EM_NARYAD_ONLYMES_PERS]    Script Date: 05.06.2021 18:14:04 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



ALTER view [dbo].[S$EM_NARYAD_ONLYMES_PERS]
as
Select
  --Naryad.[Number] as [Number]
  Naryad.persStrTabN   as f$persStrTabN
, dbo.tocomp(Naryad.PersNrec)      as f$PersNrec
, Naryad.PersFio       as f$PersFio
--, Naryad.[Status] as [Status]
, Naryad.[PodrKod]     as [f$PodrKod]
, dbo.tocomp(Naryad.[PodrNrec])    as [f$PodrNrec]
, Naryad.[totalChasF]  as [f$totalChasF]
, Naryad.[year]                   as [f$year]
, Naryad.[month]           as [f$month]
, Naryad.[day]                   as [f$day]
, Naryad.[datecreate]                   as [f$datecreate]
, Naryad.[SystemSRC]   as [f$SystemSRC]
, Naryad.[Location]           as [f$Location]
from
(
select  top 1
  tv_d_beg.F$LONGVAL as dayBeg
 ,tv_m_beg.F$LONGVAL as MonBeg
 ,tv_y_beg.F$LONGVAL as YearBeg
 ,tv_d_end.F$LONGVAL as dayEnd
 ,tv_m_end.F$LONGVAL as MonEnd
 ,tv_y_end.F$LONGVAL as YearEnd
from X$USERS xu
join T$TUNEDEF td_d_beg on td_d_beg.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_DAY_BEG'
left join t$TUNEVAL tv_d_beg on tv_d_beg.F$CTUNE = td_d_beg.F$NREC
               and tv_d_beg.f$cUser = xu.ATL_NREC
               and tv_d_beg.f$Obj   = xu.xu$useroffice

join T$TUNEDEF td_m_beg on td_m_beg.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_MONTH_BEG'
left join t$TUNEVAL tv_m_beg on tv_m_beg.F$CTUNE = td_m_beg.F$NREC
               and tv_m_beg.f$cUser = xu.ATL_NREC
               and tv_m_beg.f$Obj   = xu.xu$useroffice

join T$TUNEDEF td_y_beg on td_y_beg.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_YEAR_BEG'
left join t$TUNEVAL tv_y_beg on tv_y_beg.F$CTUNE = td_y_beg.F$NREC
               and tv_y_beg.f$cUser = xu.ATL_NREC
               and tv_y_beg.f$Obj   = xu.xu$useroffice

join T$TUNEDEF td_d_end on td_d_end.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_DAY_END'
left join t$TUNEVAL tv_d_end on tv_d_end.F$CTUNE = td_d_end.F$NREC
               and tv_d_end.f$cUser = xu.ATL_NREC
               and tv_d_end.f$Obj   = xu.xu$useroffice

join T$TUNEDEF td_m_end on td_m_end.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_MONTH_END'
left join t$TUNEVAL tv_m_end on tv_m_end.F$CTUNE = td_m_end.F$NREC
               and tv_m_end.f$cUser = xu.ATL_NREC
               and tv_m_end.f$Obj   = xu.xu$useroffice

join T$TUNEDEF td_y_end on td_y_end.F$CODE = 'NPOEM_OWN.ZP.NARYAD.CROSSREPORTNARYAD_YEAR_END'
left join t$TUNEVAL tv_y_end on tv_y_end.F$CTUNE = td_y_end.F$NREC
               and tv_y_end.f$cUser = xu.ATL_NREC
               and tv_y_end.f$Obj   = xu.xu$useroffice


where upper(xu.XU$LOGINNAME) = upper(SUSER_NAME())

) tunes
outer apply (


select
--  wc.Number as [Number]
  twce.TabNumber as persStrTabN
, twce.ERPNREC as PersNrec
, twce.FIO     as PersFio
--, case wccs.CurrentStatus when 'C' then 'Рассчитан' when 'P' then 'Сформирован' when 'S' then 'Утвержден' when 'A' then 'На согласовании' when 'E' then 'Выявлены ошибки' else wccs.CurrentStatus end [Status]
, podr.Code as PodrKod
, podr.externalID as podrNrec
, Sum(twce.TotalHoursFact) as totalChasF
, year(wc.ValuationDate) as [year]
, month(wc.ValuationDate) as [month]
, day(CAST(wc.ValuationDate as date)) as[day]
, dbo.ToAtlDate(CAST(wc.ValuationDate as date)) as[dateCReate]
, 'MES' as [SystemSRC]
, 'В MES' as [Location]
--, wc.ShopFloor
from [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCard] wc
join [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[ProductionUnit] podr on podr.oid = wc.area
join [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCardCurrentStatus] wccs on wccs.WorkCard = wc.oid
outer apply (select sum(wcE.ActivityQtyFact) as TotalHoursFact, pers.TabNumber, pers.externalID as ERPNREC, pers.Fullname as FIO
               from [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCardExecutor] wcE
                         join [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[Personnel] pers on pers.Oid = wce.Personnel
                         where wce.WorkCard = wc.Oid
                         group by pers.TabNumber, pers.externalID, pers.Fullname
                    ) twce
outer apply (select sn.F$NMNEM as [Number] from t$sys_nar sn
                                            join T$ATTRNAM anSystem on anSystem.F$WTABLE = 16003
                                                                   and anSystem.F$NAME = '.СистемаИсточник'
                                            join T$ATTRVAL avSystem on avSystem.F$CATTRNAM = anSystem.F$NREC
                                                                   and avSystem.F$WTABLE = anSystem.F$WTABLE
                                                                   and avSystem.F$CREC = sn.F$NREC
                                           where sn.F$NMNEM =  wc.Number and  avSystem.F$VSTRING = 'MES'
                                                                                         and  sn.f$choice in(3,203,901,905,906) -- смотрим в ЕРП только оплачет, оплачен,'На проверку'  'Проверено Экономистом' Оформляемый'
            ) tisExistInERP
where wc.ValuationDate between DATEFROMPARTS(tunes.YearBeg,tunes.MonBeg,tunes.DayBeg) and DATEFROMPARTS(tunes.YearEnd,tunes.MonEnd,tunes.DayEnd)
      and tisExistInERP.[Number] is null
group by --  wc.Number as [Number]
  twce.TabNumber
, twce.ERPNREC
, twce.FIO
--, case wccs.CurrentStatus when 'C' then 'Рассчитан' when 'P' then 'Сформирован' when 'S' then 'Утвержден' when 'A' then 'На согласовании' when 'E' then 'Выявлены ошибки' else wccs.CurrentStatus end [Status]
, podr.Code
, podr.externalID
, year(wc.ValuationDate)
, month(wc.ValuationDate)
, day(CAST(wc.ValuationDate as date))
 ,dbo.ToAtlDate(CAST(wc.ValuationDate as date))
) Naryad


GO


