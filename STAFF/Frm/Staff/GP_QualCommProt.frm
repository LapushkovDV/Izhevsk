/*************************************************************************************************\
* Наименование: Форма "Протокол заседания аттестационной / квалификационной комиссии"             *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Отчеты / Отчеты пользователя / Обучение и развитие персонала / Протокол заседания   *
*             аттестационной / квалификационной комиссии                                          *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#341          |30/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
* Разработка  |#342          |30/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
* Разработка  |#343          |30/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.form GP_QualCommProt
.hide
.fields
  ActEducationsNRec: Comp
  DocNum
  DocDate: Date
  ActEdFrom: Date
  ActEdTo: Date
  ActEdNote
  ActEdSubject
  ActEdType
!.{ PsnList checkEnter
  PlanEducNRec: Comp
  PersFIO
  PersTabN
  PersBornDate: Date
  PersDepCode
  PersDepName
  PersPosition
  PersEducLvl
  PersQual
  PersDiplSeries
  PersDiplNum
  PersDiplDate: Date
!.}
.endFields
  Ссылка на запись в таблице ActEducations (ActEducationsNRec): ^
  Номер протокола (DocNum): ^
  Дата протокола (DocDate): ^
  Дата начала (ActEdFrom): ^
  Дата окончания (ActEdTo): ^
  Примечание (ActEdNote): ^
  Тема (ActEdSubject): ^
  Вид повышения квалификации (ActEdType): ^

  Список работников:
.{ PsnList checkEnter

   Ссылка на запись в таблице PlanEduc (PlanEducNRec): ^
   Фамилия Имя Отчество (PersFIO): ^
   Таб № (PersTabN): ^
   Дата рождения (PersBornDate): ^
   Код подразделения (PersDepCode): ^
   Наименование подразделения (PersDepName): ^
   Должность (PersPosition): ^
   Образование (PersEducLvl): ^
   Квалификация по диплому (PersQual): ^
   Серия диплома (PersDiplSeries): ^
   Номер диплома (PersDiplNum): ^
   Дата диплома (PersDiplDate): ^
.}
.endForm
