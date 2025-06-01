/*************************************************************************************************\
* Наименование: Прототип формы "Аттестационный лист"                                              *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Сотрудники / Аттестация сотрудников / Сведения об аттестации сотрудников            *
* Примечание:                                                                                     *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#154          |09/10/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_AttList
.hide
.fields
  TypeOfAtt
  CompanyName
  AttDate: Date
  AttNmb
!.{ Employees checkEnter
  EmplFIO
  EmplBornDate
  EmplSpeciality
  EmplPos
  EmplPosDate: Date
  EmplPsnAppDate: Date
  EmplTotalSen
  EmplCompanySen
  Fortunes
  Punishments
  PredAttResult
  EmplAttDate: Date
  Recom1
  Recom2
  Rezalt
  PredsFullFIO
  PredsFIO
  ZamPredFullFIO
  ZamPredFIO
  SecretarFullFIO
  SecretarFIO
  Prisutsvovali
!.{ Educ checkEnter
  ELevelCode: Integer
  ELevel
  EName
  ESpeciality
  ETopic
  EFromDate: Date
  EToDate: Date
!.}
!.{ EducAdd checkEnter
  EALevelCode: Integer
  EALevel
  EAName
  EASpeciality
  EATopic
  EAFromDate: Date
  EAToDate: Date
!.}
!.}
.endFields
  TypeOfAtt: ^
  CompanyName: ^
  AttDate: ^
  AttNmb: ^
.{ Employees checkEnter

  EmplFIO: ^
  EmplBornDate: ^
  EmplSpeciality: ^
  EmplPos: ^
  EmplPosDate: ^
  EmplPsnAppDate: ^
  EmplTotalSen: ^
  EmplCompanySen: ^
  Fortunes: ^
  Punishments: ^
  PredAttResult: ^
  EmplAttDate: ^
  Recom1: ^
  Recom2: ^
  Rezalt: ^
  PredsFullFIO: ^
  PredsFIO: ^
  ZamPredFullFIO: ^
  ZamPredFIO: ^
  SecretarFullFIO: ^
  SecretarFIO: ^
  Prisutsvovali: ^

.{ Educ checkEnter
  ELevelCode: ^
  ELevel: ^
  EName: ^
  ESpeciality: ^
  ETopic: ^
  EFromDate: ^
  EToDate: ^

.}
.{ EducAdd checkEnter
  EALevelCode: ^
  EALevel: ^
  EAName: ^
  EASpeciality: ^
  EATopic: ^
  EAFromDate: ^
  EAToDate: ^

.}
.}
.endForm

