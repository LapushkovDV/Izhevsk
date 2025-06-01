select
  sn.F$NMNEM as [Number]
, case sn.f$choice when 3 then 'Проверен БТЗ' when 901 then 'На проверку' when 902 then 'Отклонен' when 904 then 'Ошибка' when 905 then 'Проверено Экономистом' when 906 then 'Оформляемый' else cast(sn.f$choice as nvarchar(5)) end as [Status]
, tprn.podrKod as PodrKod
, tPrn.TotalChasF as totalChasF
, year(dbo.to_sqlDate(avDate.F$VDATE)) as [year]
, month(dbo.to_sqlDate(avDate.F$VDATE)) as [month]
, day(dbo.to_sqlDate(avDate.F$VDATE)) as [day]
, Coalesce(avSystem.F$VSTRING,'ERP') as [System]
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
cross apply (select Sum(prn.f$ChasF) as TotalChasF, SubString(podr.f$kod,1,3) as podrKod
              from T$PRNARYAD prn
              join t$katpodr podr on podr.f$NREC = prn.f$cexoz
             where prn.F$MNREC = sn.F$NREC
              group by SubString(podr.f$kod,1,3)
            ) tPrn
where avDate.F$VDATE >= 2020*256*256+03*256+1

union all

select
  wc.Number as [Number]
, case wccs.CurrentStatus when 'C' then 'Рассчитан' when 'P' then 'Сформирован' when 'S' then 'Утвержден' when 'A' then 'На согласовании' when 'E' then 'Выявлены ошибки' else wccs.CurrentStatus end [Status]
, podr.Code as PodrKod
, twce.TotalHoursFact as totalChasF
, year(wc.ValuationDate) as [year]
, month(wc.ValuationDate) as [month]
, day(wc.ValuationDate) as[day]
, 'AMM' as [system]
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
where wc.ValuationDate >= DATEFROMPARTS(2020,03,01)
      and tisExistInERP.[Number] is null
union all

select
  wc.NUmber
, case wc.[status] when 0 then 'Сформирован' when 1 then 'На согласовании' when 2 then 'Утвержден' when 3 then 'Рассчитан' when 4 then 'Возвращен' when 5 then 'Ошибка' else cast(wc.[status] as nvarchar(5)) end [Status]
, podr.Код as podrKod
, twce.TotalHoursFact  as totalChasF
, year(wc.DateValuation) as [year]
, month(wc.DateValuation) as [month]
, day(wc.DateValuation) as[day]
, 'MES' as [system]
from [ammdb-cluster\ammdb].[amm_prod].[dbo].[WorkCard] wc
join [ammdb-cluster\ammdb].[amm_prod].[dbo].[ПроизводственнаяЕдиница] podr on podr.oid = wc.shopfloor
outer apply (select sum(wcE.TotalHoursFact) as TotalHoursFact from  [ammdb-cluster\ammdb].[amm_prod].[dbo].[WorkCardExecutor] wcE where wce.WorkCard = wc.Oid) twce
outer apply (select sn.F$NMNEM as [Number] from t$sys_nar sn
                                            join T$ATTRNAM anSystem on anSystem.F$WTABLE = 16003
                                                                   and anSystem.F$NAME = '.СистемаИсточник'
                                            join T$ATTRVAL avSystem on avSystem.F$CATTRNAM = anSystem.F$NREC
                                                                   and avSystem.F$WTABLE = anSystem.F$WTABLE
                                                                   and avSystem.F$CREC = sn.F$NREC
                                           where sn.F$NMNEM =  wc.Number and  avSystem.F$VSTRING = 'AMM'
            ) tisExistInERP
where wc.DateValuation >= DATEFROMPARTS(2020,03,01)
      and tisExistInERP.[Number] is null

