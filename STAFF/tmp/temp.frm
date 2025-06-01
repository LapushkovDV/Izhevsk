/*************************************************************************************************\
* Наименование: Форма "Протокол заседания тарифно-квалификационной комиссии ЭМ"                   *
* Контур/Модуль: Управление персоналом                                                            *
* Пункт меню: Отчеты / Отчеты пользователя / Обучение и развитие персонала / Протокол заседания   *
*             аттестационной / квалификационной комиссии                                          * 
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#343          |30/06/17|Тищенко Р.Н.             |НПО Энергомаш                    *
\*************************************************************************************************/
.linkForm 'GP_QualCommProt_03' prototype is 'GP_QualCommProt'
.nameInList 'ЭМ.Протокол заседания тарифно-квалификационной комиссии ЭМ'
.declare
  #include GP_StrFunc.vih // Функции работы со строками
  #include GP_Signers.vih // Библиотека для вывода подписей в отчетах
.enddeclare
.var
  oExtAttr: IExtAttr;
  OGP_SF: GP_StrFunc;
  oGP_Signers: GP_Signers;
  MemberFound: Boolean;
  MemberNum: Word;
  MemberList: String;
  IsFirst: Boolean;
.endvar
.procedure AddToMemberList(Role: String);
begin
  if oGP_Signers.GetFirst(Role) then
    do begin
      if MemberList != '' then
        MemberList += ', '; 
      MemberList += oGP_Signers.GetFIO(1, 1); 
    end while oGP_Signers.GetNext;
end.
.begin
  IsFirst := true;
  // Инициализация подписантов
  oGP_Signers.Init(cgReport_KASSFRO, 'QualCommProt', false, DocDate);
  MemberList := '';
  AddToMemberList('PredCom'); // Роль "Председатель комиссии"
  AddToMemberList('ChlenCom'); // Роль "Член комиссии"
  AddToMemberList('Otvetstv'); // Роль "Ответственный"
  AddToMemberList('Sekretar'); // Роль "Секретарь комиссии"
end.
.{ PsnList checkEnter
.{?internal; not IsFirst;

.}
.begin
  IsFirst := false;
end.
.fields
  DateToStr(DocDate, 'DD')
  DateToStr(DocDate, 'mon')
  Year(DocDate)
  MemberList
  PersDepCode + ', ' + PersPosition + ', ' + OGP_SF.ShortFIO(PersFIO)
  OGP_SF.ShortFIO(PersFIO)
.endFields
Акционерное общество \'abНПО Энергомаш\'bb
(АО \'abНПО Энергомаш\'bb)
ПРОТОКОЛ
Заседания тарифно-квалификационной комиссии
\'ab^\'bb^ ^ г. 
Состав комиссии:
Присутствовали^
(ФИО членов комиссии, присутствующих на заседании)
Повестка дня: проведение квалификационных экзаменов работников 
^ 
(наименование структурных подразделений, профессия, квалификационные разрезы, ФИО работника)
Слушали:
Представления на присвоение квалификационных разрядов (категорий) ^
(ФИО работника)
Проведены квалификационные экзамены рабочих, обучавшихся по курсовой, групповой, индивидуальной формам обучения, самоподготовка (нужное подчеркнуть).
Подведены итоги квалификационных экзаменови и принято решение о присвоении квалификационных разрядов (категорий):
№ п/пФ.И.О., место работыГод рожд.ОбразованиеПрофессия, разряд (категория) до обученияЭкзаменационная оценкаРешение квалификационной комиссии
теорет. обучениепроизвод-ое обучениепрофессияразряд (категория)
123456789
.fields
  PersFIO
  if(PersBornDate != ZeroDate, Year(PersBornDate), '')
  PersEducLvl
  PersPosition
  oExtAttr.doGetAttr(coPlanEduc, PlanEducNRec, 'Оценка за теорию')
  oExtAttr.doGetAttr(coPlanEduc, PlanEducNRec, 'Оценка за практику')
  PersQual
.endFields
1^^^^^^^

.begin
  oGP_Signers.GetFirst('PredCom'); // Роль "Председатель комиссии"
end.
.fields
  oGP_Signers.GetFIO(1, 2)
.endFields
Председатель комиссии: ^
Расшифровка подиси
.begin
  MemberFound := oGP_Signers.GetFirst('ChlenCom'); // Роль "Член комиссии"
  MemberNum := 1;
end.
.{ while MemberFound
.fields
  If(MemberNum = 1, 'Члены комиссии:', '')
  oGP_Signers.GetFIO(1, 2)
.endFields
^ ^
Расшифровка подиси
.begin
  MemberFound := oGP_Signers.GetNext;
  MemberNum++;
end.
.}
.begin
  oGP_Signers.GetFirst('Sekretar'); // Роль "Секретарь комиссии"
end.
.fields
  oGP_Signers.GetFIO(1, 2)
  OGP_SF.ShortFIO(PersFIO, 2)
.endFields
Секретарь комиссии ^
Расшифровка подиси
С протоколом ознакомлен: ^ 
Расшифровка подисидата
.begin
  oGP_Signers.GetFirst('Otvetstv'); // Роль "Ответственный"
end.
.fields
  oGP_Signers.GetFIO(1, 2)
.endFields
Начальник центра подготовки персонала^
(подпись, ФИО)
М. П.
.fields
  If(DocDate != ZeroDate, DateToStr(DocDate, 'DD'), '')
  If(DocDate != ZeroDate, DateToStr(DocDate, 'mon'), '')
  If(DocDate != ZeroDate, Year(DocDate), '')
.endFields
 \'ab^\'bb^ ^ г.
.}
.endform
