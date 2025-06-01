/*************************************************************************************************\
* Наименование:  Разработка отчетной формы                                                        *
*                РПД № 60 "Изменение режима работы"                                               *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Управление персоналом->Документы->Все приказы по персоналу->                     *
*                   Окно редактирования приказа =Приказ=->                                        *
*                   Контекстное меню <Печать групповой формы приказа>                             *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#172          |26/06/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/
.LinkForm 'GP_ChangeTimeWorkGroupGroup' prototype is 'NformRD60a'
.NameInList 'ЭМ."Приказ (распоряжение) об изменении режима работ" ЭМ'
.F 'NUL'
.create view s
as select
 *
from
  tmpChangeTime
;
.begin
  s.delete all tmpChangeTime;
end.
.{ NformRD60aCycle CheckEnter
.begin
  s.insert into tmpChangeTime set tmpChangeTime.cPers        := comp(PersNrec)
                                , tmpChangeTime.sFIO         := FIO
                                , tmpChangeTime.sFIO_VP      := FIO_VP
                                , tmpChangeTime.TabNum       := TabN
                                , tmpChangeTime.NumDoc       := номер_документа
                                , tmpChangeTime.DateDoc      := Replace(дата_составления, '/', '.')
                                , tmpChangeTime.Dol          := должность
                                , tmpChangeTime.Podr         := подразделение
                                , tmpChangeTime.dBeg         := Replace(дата_с, '/', '.')
                                , tmpChangeTime.dEnd         := Replace(дата_по, '/', '.')
                                , tmpChangeTime.WorkRegime   := Rejim
                                , tmpChangeTime.cause        := Osn;
end.
.}
.begin
  RunInterface(iEMChangeTimeWorkReport, номер_документа, Replace(дата_составления, '/', '.'), comp(TitleDocNrec));
  s.delete all tmpChangeTime;
end.
.endform

