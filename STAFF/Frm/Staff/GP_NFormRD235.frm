/*************************************************************************************************\
* Наименование: Прототип формы "Приказ о проведении обучения (наставничество)"                    *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Документы / Все приказы по персоналу                                                *
* Примечание:                                                                                     *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#338          |07/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
* Разработка  |#1166         |18/03/18|Кириллов Э.П.            |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_NFormRD235
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

  FIO
  TabN
  DtPos: Date
  Gender
  DepNRec: Comp
  DepCode
  DepName
  Position
  PsnEducLevel

  MentorNRec: Comp
  MentorFIO
  MentorTabN
  MentorGender
  MentorDepNRec: Comp
  MentorDepCode
  MentorDepName

  DateBeg: Date
  DateEnd: Date
  TestDate: Date
  TestTime: Time
  Institution
  EducLevel
  Speciality
  Subject
  Sum: Double
  EducReason
  OrderReason

  InstitutionPr
  EducLevelPr
  SpecialityPr
  SubjectPr
  DateBegPr: Date
  DateEndPr: Date

  RaiseVidOplP: LongInt
  RaiseName
  RaiseSum: Double
  RaiseType: Byte
  RaiseTypeName
!.{ PartDocMemo checkEnter
  PartDocMemoText
!.}
!.{ CommentMemo checkEnter
  CommentMemoText
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

  Фамилия Имя Отчество (FIO): ^
  Таб № (TabN): ^
  Дата приема на работу (DtPos): ^
  Пол (Gender): ^
  NRec подразделения (DepNRec): ^
  Код подразделения (DepCode): ^
  Наименование подразделения (DepName): ^
  Должность (Position): ^
  Образование (PsnEducLevel): ^

  Наставник:
   Ссылка на таблицу Persons (MentorNRec): ^
   Фамилия Имя Отчество (MentorFIO): ^
   Таб № (MentorTabN): ^
   Пол (MentorGender): ^
   Код подразделения (MentorDepNRec): ^
   Код подразделения (MentorDepCode): ^
   Наименование подразделения (MentorDepName): ^

  Дата начала (DateBeg): ^
  Дата окончания (DateEnd): ^
  Дата экзамена (TestDate): ^
  Время экзамена (TestTime): ^

  Учебное заведение (Institution): ^
  Вид образования (EducLevel): ^
  Специальность (Speciality): ^
  Тема обучения (Subject): ^
  Стоимость (Sum): ^
  Основание для обучения (EducReason): ^
  Основание для приказа (OrderReason): ^

  Образование с видом "Высшее профессиональное":
   Учебное заведение (InstitutionPr): ^
   Вид образования (EducLevelPr): ^
   Специальность (SpecialityPr): ^
   Тема обучения (SubjectPr): ^
   Дата начала (DateBegPr): ^
   Дата окончания (DateEndPr): ^

  Доплата:
   Код (RaiseVidOplP): ^
   Наименование (RaiseName): ^
   Сумма (RaiseSum): ^
   Тип (RaiseType): ^
   Наименование типа (RaiseTypeName): ^

  Шапка раздела:
.{ PartDocMemo checkEnter
   Текст (PartDocMemoText): ^
.}
  Примечание:
.{ CommentMemo checkEnter
   Текст (CommentMemoText): ^
.}
.endform
