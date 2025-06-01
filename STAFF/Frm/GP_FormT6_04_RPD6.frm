/*************************************************************************************************\
* Наименование: Приказ о переносе отпуска сотрудника                                              *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#167          |28/09/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_FormT6_04_RPD6' Prototype is 'FormT6_04'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Заявление'
.Group 'Country' subGroup 'Russia'
.f 'NUL'
.{ t6_2004_Cycle1 CheckEnter
.}
.{ t6_2004_CycleVac CheckEnter
.}
.{ t6_2004_CycleDopVac CheckEnter
.}
.begin
  var Rep:GP_RepLaborContract;
  Rep.Run(41,1,ContDocNrec,PersNrec,AppointNrec);
end.
.endform
