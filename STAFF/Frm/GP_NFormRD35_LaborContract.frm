/*************************************************************************************************\
* Наименование: Приказ о направлении на обучение                                                  *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#44           |08/11/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NFormRD35_LaborContract' Prototype is 'GP_NFormRD35'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ о направлении на обучение'
.group 'РПД 35'
.group 'РПД 135'
.f 'NUL'
.{ PDocMemoRD35 checkEnter
.}
.{ Persons checkEnter
.}
.{ Chiefs checkEnter
.}
.begin
  var Rep:GP_RepLaborContract;
  Rep.Run(TypeOper,1,ContDocNrec,PersNrec,0);
end.
.endform
