/*************************************************************************************************\
* Наименование:  Разработка отчетной формы                                                        *
*                РПД № 72 "Привлечение сотрудника к сверхурочной работе"                          *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Управление персоналом->Документы->Все приказы по персоналу->                     *
*                   Окно редактирования приказа =Приказ=->                                        *
*                   Контекстное меню <Печать индивидуальной формы приказа>                        *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#169          |26/07/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/
.LinkForm 'GP_OverTime' Prototype is 'NformRD72'
.NameInList 'ЭМ.Приказ о сверхурочной работе ЭМ'
.F 'NUL'
.create view vWV
var 
   cc : comp;
as select
 *
from
  TitleDoc 
, PartDoc
, ContDoc
, Persons
where
((
    NRecTitleDoc           == TitleDoc.NRec
AND TitleDoc.NRec          == PartDoc.cDoc
AND PartDoc.NRec           == ContDoc.cPart
AND ContDoc.Person         == Persons.Nrec
))
;

.{ NformRD72Cycle CheckEnter
.}
.begin
  if (vWV.GetFirst TitleDoc = tsOk)
    vWV._LOOP ContDoc
      if (vWV.GetFirst Persons = tsOk AND vWV.Persons.FIO = FIO)
        {
          cc := vWV.ContDoc.NRec;
          break;
        }

  RunInterface(iEMOverTimeReport, N_DOC, DATA, comp(NRecTitleDoc), cc);
end.
.endform
