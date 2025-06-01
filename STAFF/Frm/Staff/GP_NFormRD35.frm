/*************************************************************************************************\
* Наименование: Прототип формы "Приказ о производственной практике студентов"                     *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Документы / Все приказы по персоналу                                                *
* Примечание:                                                                                     *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#173          |22/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_NFormRD35
.hide
.var
!.{?Internal;
   boHead:boolean;
   wHead:word;
.endvar
.fields
  TypeOper
  TitleDocNRec: Comp
  ContDocNRec: Comp
  PersonsNRec: Comp
  OrgName
  OrgOKPO
  OrderNum
  OrderDate: Date
  DateBeg: Date
  DateEnd: Date
  TestDate: Date
  TestTime: Time
  Institution
  InstLongName
  InstCity
  EducLevel
  Speciality
  Subject
  Sum: Double
  EducReason

  TeacherNRec: Comp
  TeacherFIO
  TeacherTabN
  TeacherDepCode
  TeacherDepName

!.{ PDocMemoRD35 checkEnter
  PartDocMemoText
!.}

  CommonSchedule
  CommonNote
!.{ Persons checkEnter
  PersNRec: Comp
  PersFIO
  PersTabN
  PersDepCode
  PersDepName
  PersPosition
  PersSchedule
!.}
!.{ Chiefs checkEnter

  ChiefNRec: Comp
  ChiefFIO
  ChiefTabN
  ChiefDepCode
  ChiefDepName
  ChiefPosition
!.}
.endFields
  РПД (TypeOper): ^
  ссылка на TitleDoc (TitleDocNRec): ^
  ссылка на ContDoc (ContDocNRec): ^
  ссылка на Persons (PersonsNRec): ^
  Наименование организации (OrgName): ^
  ОКПО организации (OrgOKPO): ^
  Номер приказа (OrderNum): ^
  Дата приказа (OrderDate): ^
  Дата начала (DateBeg): ^
  Дата окончания (DateEnd): ^
  Дата экзамена (TestDate): ^
  Время экзамена (TestTime): ^
  Учебное заведение (Institution): ^
  Длинное наименование учебного заведения (InstLongName): ^
  Город учебного заведения (InstCity): ^

  Вид образования (EducLevel): ^
  Специальность (Speciality): ^
  Тема обучения (Subject): ^
  Стоимость (Sum): ^
  Основание для обучения (EducReason): ^

  Преподаватель:
   Ссылка на таблицу Persons (TeacherNRec): ^
   Фамилия Имя Отчество (TeacherFIO): ^
   Таб № (TeacherTabN): ^
   Код подразделения (TeacherDepCode): ^
   Наименование подразделения (TeacherDepName): ^

  Шапка раздела:
.{ PDocMemoRD35 checkEnter
   Текст (PartDocMemoText): ^
.}

  График занятий - первое заполненное поле (CommonSchedule): ^
  Примечание - первое заполненное поле (CommonNote): ^

  Список обучающихся:
.{ Persons checkEnter

   Persons.NRec (PersNRec): ^
   Фамилия Имя Отчество (PersFIO): ^
   Таб № (PersTabN): ^
   Код подразделения (PersDepCode): ^
   Наименование подразделения (PersDepName): ^
   Должность (PersPosition): ^
   График занятий (PersSchedule): ^
.}

  Список начальников подразделений обучающихся:
.{ Chiefs checkEnter

   Persons.NRec (ChiefNRec): ^
   Фамилия Имя Отчество (ChiefFIO): ^
   Таб № (ChiefTabN): ^
   Код подразделения (ChiefDepCode): ^
   Наименование подразделения (ChiefDepName): ^
   Должность (ChiefPosition): ^
.}
.endForm
