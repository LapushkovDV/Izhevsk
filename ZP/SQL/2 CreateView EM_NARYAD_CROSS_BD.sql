ALTER view [dbo].[S$EM_NARYAD_CROSS_BD] as
Select
  Naryad.[Number] as [Number]
, Naryad.[Status] as [Status]
, Naryad.[PodrKod]     as [PodrKod]
, Naryad.[totalChasF]  as [totalChasF]
, Naryad.[year]                   as [year]
, Naryad.[month]           as [month]
, Naryad.[day]                   as [day]
, Naryad.[SystemSRC]   as [SystemSRC]
, Naryad.[Location]           as [Location]
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
  sn.F$NMNEM as [Number]
, case sn.f$choice when 3 then 'Проверен БТЗ'
                   when 901 then 'На проверку'
                                   when 902 then 'Отклонен'
                                   when 904 then 'Ошибка'
                                   when 905 then 'Проверено Экономистом'
                                   when 906 then 'Оформляемый'
                                   when 203 then 'Оплачен'
                                   else cast(sn.f$choice as nvarchar(5)) end as [Status]
, tprn.podrKod as PodrKod
, tPrn.TotalChasF as totalChasF
, year(dbo.to_sqlDate(tPrNarDat.F$DATAN)) as [year]
, month(dbo.to_sqlDate(tPrNarDat.F$DATAN)) as [month]
, avDate.F$VDATE as [day]
, Coalesce(avSystem.F$VSTRING,'ERP') as [SystemSRC]
, 'В ERP' as [Location]
from t$sys_nar sn
join T$ATTRNAM anDate on anDate.F$WTABLE = 16003
                 and anDate.F$NAME = 'Дата создания наряда'
join T$ATTRVAL avDate on avDate.F$CATTRNAM = anDate.F$NREC
                     and avDate.F$WTABLE = anDate.F$WTABLE
                                and avDate.F$CREC = sn.F$NREC
join T$ATTRNAM anSystem on anSystem.F$WTABLE = 16003
                       and anSystem.F$NAME = '.СистемаИсточник'
left join T$ATTRVAL avSystem on avSystem.F$CATTRNAM = anSystem.F$NREC
                            and avSystem.F$WTABLE = anSystem.F$WTABLE
                                       and avSystem.F$CREC = sn.F$NREC
outer apply (select top 1 prn.f$DATAN
              from T$PRNARYAD prn
             where prn.F$MNREC = sn.F$NREC
                         ) tPrNarDat

cross apply (select Sum(prn.f$ChasF) as TotalChasF, SubString(podr.f$kod,1,3) as podrKod
              from T$PRNARYAD prn
              join t$katpodr podr on podr.f$NREC = prn.f$cexoz
             where prn.F$MNREC = sn.F$NREC
              group by SubString(podr.f$kod,1,3)
            ) tPrn
where -- tPrNarDat.F$DATAN  between tunes.YearBeg*256*256+tunes.MonBeg*256+tunes.DayBeg and tunes.YearEnd*256*256+tunes.MonEnd*256+tunes.DayEnd
 avDate.F$VDATE between tunes.YearBeg*256*256+tunes.MonBeg*256+tunes.DayBeg and tunes.YearEnd*256*256+tunes.MonEnd*256+tunes.DayEnd

union all

select
  wc.Number as [Number]
, case wccs.CurrentStatus when 'C' then 'Рассчитан' when 'P' then 'Сформирован' when 'S' then 'Утвержден' when 'A' then 'На согласовании' when 'E' then 'Выявлены ошибки' else wccs.CurrentStatus end [Status]
, podr.Code as PodrKod
, twce.TotalHoursFact as totalChasF
, year(wc.ValuationDate) as [year]
, month(wc.ValuationDate) as [month]
, dbo.ToAtlDate(CAST(wc.ValuationDate as date)) as[day]
, 'MES' as [SystemSRC]
, 'В MES' as [Location]
--, wc.ShopFloor
from [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCard] wc
join [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[ProductionUnit] podr on podr.oid = wc.shopfloor
join [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCardCurrentStatus] wccs on wccs.WorkCard = wc.oid
outer apply (select sum(wcE.ActivityQtyFact) as TotalHoursFact from [dep968-amm-mes\mes].[amm_mes_prod].[dbo].[WorkCardExecutor] wcE where wce.WorkCard = wc.Oid) twce
outer apply (select sn.F$NMNEM as [Number] from t$sys_nar sn
                                            join T$ATTRNAM anSystem on anSystem.F$WTABLE = 16003
                                                                   and anSystem.F$NAME = '.СистемаИсточник'
                                            join T$ATTRVAL avSystem on avSystem.F$CATTRNAM = anSystem.F$NREC
                                                                   and avSystem.F$WTABLE = anSystem.F$WTABLE
                                                                   and avSystem.F$CREC = sn.F$NREC
                                           where sn.F$NMNEM =  wc.Number and  avSystem.F$VSTRING = 'MES'
            ) tisExistInERP
where wc.ValuationDate between DATEFROMPARTS(tunes.YearBeg,tunes.MonBeg,tunes.DayBeg) and DATEFROMPARTS(tunes.YearEnd,tunes.MonEnd,tunes.DayEnd)
      and tisExistInERP.[Number] is null
union all

select
  wc.NUmber
, case wc.[status] when 0 then 'Сформирован' when 1 then 'На согласовании' when 2 then 'Утвержден' when 3 then 'Рассчитан' when 4 then 'Возвращен' when 5 then 'Ошибка' else cast(wc.[status] as nvarchar(5)) end [Status]
, podr.Код as podrKod
, twce.TotalHoursFact  as totalChasF
, year(wc.DateValuation) as [year]
, month(wc.DateValuation) as [month]
, dbo.ToAtlDate(CAST(wc.DateValuation as date)) as[day]
, 'AMM' as [SystemSRC]
, 'В AMM' as [Location]
from [ammdb-cluster\ammdb].[amm_prod].[dbo].[WorkCard] wc
join [ammdb-cluster\ammdb].[amm_prod].[dbo].[ProductionUnit] podr on podr.oid = wc.shopfloor
outer apply (select sum(wcE.TotalHoursFact) as TotalHoursFact from  [ammdb-cluster\ammdb].[amm_prod].[dbo].[WorkCardExecutor] wcE where wce.WorkCard = wc.Oid) twce
outer apply (select sn.F$NMNEM as [Number] from t$sys_nar sn
                                            join T$ATTRNAM anSystem on anSystem.F$WTABLE = 16003
                                                                   and anSystem.F$NAME = '.СистемаИсточник'
                                            join T$ATTRVAL avSystem on avSystem.F$CATTRNAM = anSystem.F$NREC
                                                                   and avSystem.F$WTABLE = anSystem.F$WTABLE
                                                                   and avSystem.F$CREC = sn.F$NREC
                                           where sn.F$NMNEM =  wc.Number and  avSystem.F$VSTRING = 'AMM'
            ) tisExistInERP
where wc.DateValuation between DATEFROMPARTS(tunes.YearBeg,tunes.MonBeg,tunes.DayBeg) and DATEFROMPARTS(tunes.YearEnd,tunes.MonEnd,tunes.DayEnd)
      and tisExistInERP.[Number] is null
) Naryad
GO


