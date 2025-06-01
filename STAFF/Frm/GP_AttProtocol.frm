/*************************************************************************************************\
* Наименование: Прототип формы "Протокол заседания аттестационной комиссии"                       *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Сотрудники / Аттестация сотрудников / Сведения об аттестации сотрудников            *
* Примечание:                                                                                     *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#741          |20/12/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_AttProtocol
.hide
.fields
  TypeOfAtt
  CompanyName
  AttDate: Date
  AttNmb
  PredsFullFIO
  ZamPredFullFIO
  SecretarFullFIO
  Prisutsvovali
  Invited
!.{ PEmployees checkEnter
  EmplFIO
  EmplPos
  EmplDeptFull
  Recommendations
  Result
!.}
.endFields
  TypeOfAtt: ^
  CompanyName: ^
  AttDate: ^
  AttNmb: ^
  PredsFullFIO: ^
  ZamPredFullFIO: ^
  SecretarFullFIO: ^
  Prisutsvovali: ^
  Invited: ^
.{ PEmployees checkEnter
  EmplFIO: ^
  EmplPos: ^
  EmplDeptFull: ^
  Recommendations: ^
  Result: ^
.}
.endForm
