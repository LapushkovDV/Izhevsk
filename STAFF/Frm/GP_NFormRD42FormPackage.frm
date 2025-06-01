/*************************************************************************************************\
* Наименование: Приказ о переносе отпуска сотрудника                                              *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#167          |29/09/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NFormRD42FormPackage' Prototype is 'NFormRD42FormPackage'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ о переносе отпуска сотрудника'
.f 'NUL'
.{ NformRD42PackageCycle CheckEnter
.}
.{ NformRD42InsideCycle CheckEnter
.}
.begin
  var Rep:GP_RepLaborContract;
  Rep.Run(42,1,comp(ContDocNrec),comp(PersNrec),0);
end.
.endform
