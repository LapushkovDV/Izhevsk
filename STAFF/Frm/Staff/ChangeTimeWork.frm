/*************************************************************************************************\
* Наименование:  Разработка отчетной формы                                                        *
*                РПД № 60 "Изменение режима работы"                                               *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Управление персоналом->Документы->Все приказы по персоналу->                     *
*                   Окно редактирования приказа =Приказ=->                                        *
*                   Контекстное меню <Печать индивидуальной формы приказа>                        *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#172          |26/06/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/
.LinkForm 'GP_ChangeTimeWork' prototype is 'NformRD60'
.NameInList 'ЭМ."Приказ (распоряжение) об изменении режима работ" ЭМ'
.F 'NUL'
.create view s
as select
 *
from
  tmpChangeTime
;
.begin
  var ss : string; ss := WorkRegime;
  s.delete all tmpChangeTime;
  s.insert into tmpChangeTime set tmpChangeTime.cPers        := comp(PersNrec)
                                , tmpChangeTime.sFIO         := FIO
                                , tmpChangeTime.sFIO_VP      := FIO_VP
                                , tmpChangeTime.TabNum       := TabN
                                , tmpChangeTime.NumDoc       := номер_документа
                                , tmpChangeTime.DateDoc      := Replace(дата_составления, '/', '.')
                                , tmpChangeTime.Dol          := подразделение
                                , tmpChangeTime.Podr         := должность
                                , tmpChangeTime.dBeg         := Replace(DateBeg, '/', '.')
                                , tmpChangeTime.dEnd         := Replace(DateEnd, '/', '.')
                                , tmpChangeTime.WorkRegime   := ss
                                , tmpChangeTime.cause        := Found;

  RunInterface(iEMChangeTimeWorkReport, номер_документа, Replace(дата_составления, '/', '.'), comp(TitleDocNrec));
  s.delete all tmpChangeTime;
end.
.endform
