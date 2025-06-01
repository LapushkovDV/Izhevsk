/*************************************************************************************************\
* Наименование: Прототип формы "Приказ о производственной практике студентов"                     *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Документы / Все приказы по персоналу                                                *
* Примечание:                                                                                     *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#339          |14/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_NFormRD201
.hide
.fields
  TypeOper: Word
  TitleDocNRec: Comp
  PartDocNRec: Comp
  OrgName
  OrgOKPO
  OrderNum
  OrderDate: Date
  DateBeg: Date
  DateEnd: Date
!.{ PracticeTypes checkEnter
  PracticeTypeCode
  PracticeTypeName
!.}
!.{ Insts checkEnter
  InstCode
  InstName
!.{ Courses checkEnter
  Course: Byte
!.}
!.}
!.{ RaisesList checkEnter
  RLVidOplP: LongInt
  RLName
  RLSum: Double
  RLType: Byte
  RLTypeName
!.{ RaisesByMentors checkEnter
  RLMentorNRec: Comp
  RLMentorTabN
  RLMentorFIO
  RLRaisesCnt
!.}
!.}
!.{ ContDoc checkEnter
  CDNRec: Comp
  CDTabN
  CDFIO
  CDDepCode
  CDDepName
  CDDepNameFull
  CDPosition
  CDPractType
  CDInstitution
  CDDateBeg: Date
  CDDateEnd: Date
  CDCourse: Byte
  CDReason
!.{ Raises checkEnter
  RaiseVidOplP: LongInt
  RaiseName
  RaiseSum: Double
  RaiseType: Byte
  RaiseTypeName
!.}
!.{ Trainees checkEnter
  TrnNRec: Comp
  TrnTabN
  TrnFIO
  TrnDateBeg: Date
  TrnDateEnd: Date
  TrnSpeciality
!.}
!.}
.endFields
  РПД (TypeOper): ^
  ссылка на TitleDoc (TitleDocNRec): ^
  ссылка на PartDoc (PartDocNRec): ^
  Наименование организации (OrgName): ^
  ОКПО организации (OrgOKPO): ^
  Номер приказа (OrderNum): ^
  Дата приказа (OrderDate): ^
  Наименьшая дата начала практики (DateBeg): ^
  Наибольшая дата окончания практики (DateEnd): ^

.{ PracticeTypes checkEnter
  Вид практики: Код (PracticeTypeCode): ^ Наименование (PracticeTypeName): ^
.}

.{ Insts checkEnter
  Учебное заведение: Код (InstCode): ^ Наименование (InstName): ^
.{ Courses checkEnter
   Курс(Course): ^
.}
.}

 Доплаты, назначаемые в приказе:
.{ RaisesList checkEnter

  Доплата:
   Код вида оплаты из модуля зарплата (RLVidOplP): ^
   Наименование (RLName): ^
   Сумма (RLSum): ^
   Тип (RLType): ^
   Наименование типа (RLTypeName): ^
.{ RaisesByMentors checkEnter

    Руководитель:
     Ссылка на Persons (RLMentorNRec): ^
     Таб № (RLMentorTabN): ^ ФИО (RLMentorFIO): ^
     Кол-во доплат (RLRaisesCnt): ^
.}
.}

.{ ContDoc checkEnter

  Руководитель практики:
   Ссылка на Persons (CDNRec): ^
   Таб № (CDTabN): ^ ФИО (CDFIO): ^
   Подразделение: Код (CDDepCode): ^ Наименование (CDDepName): ^
    Полное наименование с кодом (CDDepNameFull): ^
   Должность (CDPosition): ^
   Вид практики (CDPractType): ^
   Учебное заведение (CDInstitution): ^
   Дата поступления (CDDateBeg): ^
   Дата окончания (CDDateEnd): ^
   Курс (CDCourse): ^
   Основание (CDReason): ^
.{ Raises checkEnter

   Доплата:
    Код вида оплаты из модуля зарплата  (RaiseVidOplP): ^
    Наименование (RaiseName): ^
    Сумма (RaiseSum): ^
    Тип (RaiseType): ^
    Наименование типа (RaiseTypeName): ^
.}
.{ Trainees checkEnter

   Практикант:
    Ссылка на Persons (TrnNRec): ^
    Таб № (TrnTabN): ^ ФИО (TrnFIO): ^
    Дата начала практики (TrnDateBeg): ^
    Дата окончания практики (TrnDateEnd): ^
    Специальность (TrnSpeciality): ^
.}
.}
.endForm
