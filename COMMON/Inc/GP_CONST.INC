/*************************************************************************************************\
* Наименование: Интерфейс просмотра/редактирования листа допустимых замен                         *
* Контур/Модуль: КОУ                                                                              *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |              |13/11/17|Кольцов К                |НПО Энергомаш                    *
\*************************************************************************************************/

const
  //--------------------------------------
  //#884 Распределение ДО и Накладных по схеме договоров GP_BU
  // см em_dev_bu\Vip\F_DISTPL##RAZNSOPHOZ.Vip - Init, _DoDistrSoprHoz
  // может быть 32-35; ПИР 102.173610
  // Чтобы по умолчанию (fc_gp_PLDOCDISTR_DOG_Scheme and 0111b) давало 3, берем 35
  fc_gp_PLDOCDISTR_DOG_Scheme : word= 35;
  fc_gp_PLDOCDISTR_KOL        : word= 32;//Минск

  coGP_SPXChange = 31587;
  pAttr_GP_SPXChCOMMENT = '.Комментарий';
  //---- НВП Листы согласований
  vidDocAgreed      = 60965; //Лист согласования
  vidDocSpAgreed    = 60966; //Спецификация листа согласования
  statusOtstName    = 'Требуется согласование отступлений';//Статус накладной, при котором возможно создать лист согласования
  KatKauPersonsName = 'Лист согласования ТН';//Набор аналитики "Лист согласования ТН" (элементарная аналитика набора - Сотрудники предприятия")
  extClassRoleName  = 'Роль сотрудника листа согласования';//внешняя классификация к тбл. GP_SpAgreed
  extUserDescrInitAgreed = 'Инициатор (создатель) листа согласования';

  //--------------------------------------
  LChange_TDK : word  = 51101;
  AChange_TDK : word  = 51102;
  LChangeSp_TDK : word  = 51103;
  toACTTDK    : word  =   801;
  toSPACTTDK  : word  =   802;
  toKlPrice   : word  =   803;
  TipZAM_name : string  = 'ТИП замен';
  BaseZAM_Name : string  = 'ЭМ.Основание замены';
  //--------------------------------------------
  stRoleSpecClaim  : string = 'SpecClaim';  // Код роли Сотрудника по заполнению заявки
  stRoleAdmin      : string = 'Admin';     //  Код роли Администратора
  //----------------------------------
  complDvigDM      : string = 'Комплектование двигателя ДМ(граммы)';     //Наименование назначения
  razcomplDvigDM   : string = 'Разукомплектование двигателя ДМ(граммы)'; //Наименование назначения
  //----------------------------------
  // типы заявок на базе MNPLAN
  fc_gp_LotStatus        : word = 01; //Флаг работы с документом "Лот"
  fc_gp_WorkPlanStatus   : word = 02; //Флаг работы с документом "Рабочий план исполнителя"
  fc_gp_PlanBuyStatus    : word = 03; //Флаг работы с документом "План закупок"
  fc_gp_ClaimPodrStatus  : word = 04; /***Заявки подразделений***/
  fc_gp_ClaimPodrStatusEx: word = 05;
  fc_gp_ZapasStatus      : word = 06; /***Заявки СЗ***/
  fc_gp_ZapasStatusEx    : word = 07;
  fc_gp_KPotrStatus      : word = 08; /***Консолидированная потребность***/
  fc_gp_MZakazStatus     : word = 09; /***Производственный заказ***/
  fc_gp_NeedMtStatus     : word = 10; //Потребность производства в материальных ресурсах
  fc_gp_InvNezavShablon  : word = 11;
  fc_gp_InvMPZShablon    : word = 12;
  fc_gp_OSReservStatus   : word = 13; //Инвентаризация резерва под обесценение ОС
  fc_gp_NMAReservStatus  : word = 14; //Инвентаризация резерва под обесценение НМА
  fc_gp_InvNZPShablon    : word = 17; //Инвентаризация резерва под обесценение НЗП
  fc_gp_TZRShablon       : word = 18; //Транспортно-заготовительные расходы
  fc_gp_ClaimPodrUslStatus    : word = 15; /***Заявки подразделений Услуги***/
  fc_gp_ClaimPodrUslStatusEx  : word = 16;
  fc_gp_LotEdStatus           : word = 19; //Флаг работы с документом "Лот"
  fc_gp_ClaimPodrEdStatus     : word = 20; /***Заявки подразделений***/
  fc_gp_ClaimPodrEdStatusEx   : word = 21;
  fc_gp_LotFinStatus          : word = 22; //Флаг работы с документом "Лот"
  fc_gp_PlanZipStatus         : word = 23; //Флаг работы с документом "План ЗИП"

   CmWinVisual345 : word = 345; //#3636 - Событие, при котором происходит расчет координат видимых элементов окна интерфейса

  //---- AktOFP --- Платежный календарь / Документы / Журнал обязательств
  //---- #895 ---
  VA_AktOFP_DDoc_First      ='.Первоначальная дата погашения';//дата
  VA_AktOFP_FControl_Passed ='.Финансовый контроль пройден';//строка
  VA_AktOFP_FControl_Descr  ='.Финансовый контроль описание';//строка
  //---- AktOFP <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  //-- KATMC >>>>>>>>>>>>>>>>>>>>>>>>>>>>
  VA_KATMC_1C_BKOD     = 'Номенклатурный номер 1С';
  //---#1513 --
  VA_KATMC_LockMC      = 'Запись заблокирована';
  //------- #1410
  VA_KATMC_Create_Pers = 'Создано, Сотрудник';
  VA_KATMC_Create_Date = 'Создано, Дата';
  VA_KATMC_Create_Time = 'Создано, Время';
  VA_KATMC_Confirm_Pers= 'Утвердил, Сотрудник';
  VA_KATMC_Confirm_Date= 'Утвердил, Дата';
  VA_KATMC_Confirm_Time= 'Утвердил, Время';
  VA_KATMC_GroupMC     = 'Группа МЦ для заявки';
  VA_KATMC_Zauv_KatMC_L= 'МЦ для отмененной заявки';
  VA_KATMC_Zauv_Number = 'Номер заявки';
  VA_KATMC_Reason2Refusal= 'Причина отказа';
  VA_KatMC_Marka          = 'Марка';  //к тбл. KatMC
  //#962.#1 вывести значение ВА "Профиль" к карточке МЦ
   //VA_KatMC_Razmer         = 'Размер'; //к тбл. KatMC
  VA_KatMC_Razmer         = 'Профиль'; //к тбл. KatMC
  VA_MC_BSchet         = '.Счет МЦ'; //вн. атрибут к партиям, группам МЦ, спецификациям накладных

  VA_KATMC_GrZB = '.Группы забалансового счета';// #5945 ВА к KATMC ".Группа забалансового счета"  ссылка на Пользовательскую аналитику "Группы забалансового счета".
  VA_KATMC_Proizvoditel = 'Производитель';
  VA_GRMC_NeedZB= '.Признак учета на забалансе';// #5945 ВА к GROUPMC ".Признак учета на забалансе"  вещественное число = Логический
  //-- KATMC <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  //-- KATUSL >>>>>>>>>>>>>>>>>>>>>>>>>>>>
  //---- #2580
  //----- Сотовая связь >>>>>>>>
  VA_KatUsl_NoLimit    ='.Безлимит' ;
  VA_KatUsl_Kateg      ='.Категория';
  VA_KatUsl_LSchet     ='.Лицевой счет';
  VA_KatUsl_Phone      ='.Номер телефона';
  VA_KatUsl_Operator   ='.Оператор'  ;
  VA_KatUsl_DTLock     ='.Дата блокировки';
  VA_KatUsl_Persons    ='.ЛС для печати Сотрудник';
  VA_KatUsl_Podr       ='.ЛС для печати Подразделение';
  VA_KatUsl_NoIncLimit ='.Не включать в лимит';
  //----- Сотовая связь <<<<<<<<<<
  //-- KATUSL <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  //--------- Партии товаров
  //Наименование внш. атрибутов и вкш. классификаторов для добавленных полей.
  VA_KatParty_TARAAllowRozliv  = 'Тара. Возможность розлива';
  VA_KatParty_TARAEdIzm        = 'Тара. Ед.изм';
  VA_KatParty_TARAFasovka      = 'Тара. Фасовка';

  VA_KatParty_StatusSertif   = 'Статус сертификата';
  VA_KatParty_Prior          = 'Приоритет';
  VA_KatParty_DSertif        = 'Дата выдачи сертификата';
  VA_KatParty_NAnaliz        = 'Номер лабораторного анализа';
  VA_KatParty_Zakaz          = 'Заказ';
  VA_KatParty_KatSopr        = '.Приходный документ партии (ЭМ)';
  VA_KatParty_SchFact        = '.Счет-фактура партии (ЭМ)';
  VA_KatParty_Scan           = 'Скан-копии';
  VA_KatParty_CtrlPrc        = 'Проведение входного контроля';
  VA_KatParty_IGK            = 'Идентификатор государственного контракта';
  VA_KatParty_Zakl           = 'Заключение';
  VA_KatParty_Papka          = 'Номер папки';
  VA_KatParty_Prim           = 'Примечание';
  VA_KatParty_Izgotov        = 'Изготовитель';
  VA_KatParty_DopKol         = 'Доп.количество';
  VA_KatParty_NumPlav        = 'Номер плавки';
  VA_KatParty_NumPart        = 'Номер партии';
  VA_KatParty_DPriemka       = 'Дата тех.приемки';
  VA_KatParty_RtPriemka      = 'РТ-Техприемка';
  VA_KatParty_CheckVP        = 'ВП контроль';
  VA_KatParty_Drag           = 'ТМЦ-ДМ';  //внш. атрибут к KatParty с типом "число" для ед. изм. "процент" у драг. металлов в МЦ и партиях
  VA_KatParty_Manuf          = 'Из производства';//внш. атрибут к KatParty с типом "число"
  VA_KatParty_Source         = '.ПартияИсточник';

  VA_KatParty_NumObrzType    = 'Тип номер образца';
  VA_KatParty_DateObrz       = 'Дата присвоения номера образца';
  VA_KATPARTY_NumObrz_User   = 'Номер образца, Присвоил';// заполнять ФИО пользователя, который нажал на '...' около поля "Номер образца" либо внес ручные исправления поля.

  VA_KATPARTY_Priemka_User   = 'Дата тех.приемки, Провел';// заполнять ФИО пользователя, выбравшего дату тех.приемки
  VA_KatParty_GTD            = 'Номер ГТД'           ;//#3748 sGetTune('Nalog.Rus.GTD.Attr_GTDNumber');  
  //-------------------
  EXC_KatParty_Country  = 'Страна происхождения';//#3748 sGetTune('Nalog.Rus.GTD.Class_Country');
  EXC_KATPARTY_CtrlVid  = 'Виды контроля';
  VA_Katparty_CheckSrok = 'Контроль срока годности';
  //-------- KATPARTY <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  //==========================================
  // AllMemo
  // Для партий
  VA_AllMemo_AnlzDate        = 'Дата передачи на анализ';
  VA_AllMemo_CtrlDate        = 'Дата протокола контроля';
  VA_AllMemo_CtrlCheck       = 'Несоответствие';

  //==========================================
  // Подразделение
  VA_KatPodr_PodrZatr         = '.Подразделение отнесения затрат';
  VA_KatPodr_NumObrz          = 'Номер образца'; //внш. атрибут к Katpodr с типом "строка" для хранения последнего
                                                 //использованного номера образца партии для подразделения


  //==========================================
  // Организации
  //------- #921
  VA_KatOrg_FullNAME = 'Полное наименование';
  //------- #1487
  VA_KATORG_Create_Pers  = 'Создано, Сотрудник';
  VA_KATORG_Create_Date  = 'Создано, Дата';
  VA_KATORG_Create_Time  = 'Создано, Время';
  VA_KATORG_Confirm_Pers = 'Утвердил, Сотрудник';
  VA_KATORG_Confirm_Date = 'Утвердил, Дата';
  VA_KATORG_Confirm_Time = 'Утвердил, Время';
  VA_KATORG_Group        = 'Группа для заявки';
  VA_KATORG_Zauv_KatOrg_L= 'Организация для отмененной заявки';
  VA_KATORG_Zauv_Number  = 'Номер заявки';
  VA_KatOrg_Zauv_Source  = 'Организация источник';
  VA_KatOrg_Zauv_DTIZM   = 'Заявка на изменение.Дата изменения';
  VA_KatOrg_Black        = 'Черный список'; //Стандартный функционал!!!
  VA_KATORG_Reason2Refusal='Причина отказа';
  //---------
  VA_KatOrg_Ur_Adress    = 'Юридический адрес';
  VA_KatOrg_Persons      = '.Сотрудник (ссылка на карточку)';//#219?, ссылка на Persons
  //---------
  VA_KatBank_Zauv_Source = 'Расчетный счет - источник';
  VA_SpSopr_Post_MCName  = 'Наименование поставщика'; // Использовать константу. Наименование атрибута будет переименовано


  VA_SpSopr_IntegrERPAMM  = 'Интеграция остатков ЕРП АММ';

  //------------ #971
  VA_Katsopr_ApproveDiffPrice  = '.Разрешить расхождение цен накладной и договора';
  VA_Katsopr_Dogovor     ='.Договор' ;
  VA_Katsopr_AppDogovor  ='.Соглашение';
  VA_KatSopr_CalPlan     ='.Календарный план';
  VA_KatSopr_PKP         ='.Финансовый ПКП';
  VA_KatSopr_KatOrg      ='.Организация';
  VA_KatSopr_OsnovZameny = '.Основание замены';
  //------------ #4002
  VA_KatSopr_BTK_VP      ='.Предъявить БТК ВП';// - вещественное число
  VA_KatSopr_USLOVNOST_BTK_VP      ='.Условность (БТК)';// - вещественное число

  //-------- #877
  VA_BASEDOC_CALPLANF    = '.Финансовый ПКП';
  //-------- #906
  VA_BASEDOC_NomZNP      = '.Номер ЗНП (АММ)';
  VA_BASEDOC_PKP2        = '.Финансовый ПКП';//ссылка на CalPlan Финансовый ПКП в ДО на предоплату, ПКП на готовую продукцию в ДО на давальческое сырье
  //-------- #889
  VA_SPSTEP_Comment      = '.ДО на закупку. Комментарий к спецификации';
  //-------- #906
  VA_SPSTEP_AMM          = '.Спецификация ДС (из АММ)';
  //-------- #889
  VA_SPSOPR_Comment      = '.ДО на закупку. Комментарий к спецификации'; //#2579 07.12.18 изменен с ".Комментарий по спецификации ТСД (ЭМ)"

//

  //-------- #2579
  VA_SpSopr_StZatr       = '.Статья затрат';
  VA_SpSopr_NU           = '.НУ для спецификации';
  VA_StZatr_NU           = '.НУ';
  //--------- #971
  VA_SpSopr_GrControl    ='.Группа контроля (АММ)';
  //--------- #3829
  VA_SpSopr_cSpXChange   = '.Позиция листа замены';
  VA_SpSopr_cMCDSE       = 'ДСЕ';
  VA_SpSopr_cMCpoKTD     = 'Затребованная ТМЦ по КТД'; // приходит из АММ
  VA_SpSopr_RazrOtpusk   = 'Разрешено к отпуску'; // Разрешено к отпуску.
  //--------- #4626
  VA_SpSopr_OS           = '.Объект ремонта'; //.Объект ОС
  //--------- #3642
  VA_SpSopr_PodrIsp      = '.Подразделение-исполнитель';
  VA_SpSopr_PodrFoundMarriage      ='.Цех, обнаруживший брак';
  VA_SpSopr_PodrOriginatorMarriage ='.Цех-виновник брака';
  //--------- #EM-199
  VA_SpSopr_cSpXChangeKoef         ='.Позиция листа замены - коэффициент';  

  VA_SpSopr_ZNP                    ='ЗНП';
  VA_SpSopr_SalePrice              = '.Цена реализации';
  //---------


  //-- MNPLAN
  VA_MnPlan_VidDeyat = '.Виды деятельности';
  VA_MnPlan_KodOperDeklNDS = 'Код операции для декларации по НДС';
  VA_MnPlan_Katorg = 'Контрагент';
  VA_MnPlan_Stage = '.Этап';
  VA_MnPlan_KindMnp = 'Вид заказа';
  VA_MnPlan_KindAct = '.Виды деятельности';
  VA_MnPlan_TypeMnp = '.Тип заказа';
  VA_MnPlan_Niokr = '.НИОКР';
  VA_MnPlan_Reason = '.Основание для открытия заказа';
  VA_MnPlan_Quality = '.Шифр затрат на качество';
  VA_MnPlan_CFO = 'ЦФО';
  VA_MnPlan_Program = '.Программа';

  //---------
  VA_KatVidD_SpecF3      = '.Тип выбора в спецификации договора'; //#924
  //--- #2553 -------------
  VA_MBPIN_Zakaz         = '.Шифр заказа'  ;
  VA_MBPIN_StZatr        = '.Статья затрат';
  VA_MBPIN_SrIsp         = 'Срок использ.' ;// MBPIN.SROKBEG

  //---- 2580 ------
  VA_CalPlan_Prim      ='.Примечание' ;
  VA_CalPlan_SS_Limit  ='.Сотовая связь. Лимит суммы';
  VA_CalPlan_Comment   ='.Комментарий';
  VA_SpDocs_Comment    ='.Комментарий';

  VA_SpDocs_SS_Ost     ='.Сотовая связь. Остаток на начало';
  VA_SpDocs_SS_Ost_E   ='.Сотовая связь. Остаток на конец';
  VA_SpDocs_SS_Rasx    ='.Сотовая связь. Израсходовано';
  VA_SpDocs_SS_Opl_Empl='.Сотовая связь. Оплата сотрудником';
  VA_SpDocs_SS_Opl     ='.Сотовая связь. Оплата предприятием';
  VA_SpDocs_SS_ZP      ='.Сотовая связь. Удержать из зарплаты';
  VA_SpDocs_SS_Sopr    ='.Сотовая связь. Акт';
  //----- #219
  VA_Persons_KatOrg   ='Физическое лицо (контрагент)';

  //----- #1071
  VA_GrSkl             ='Группа склада';
  //---- #3566
  VA_SpOrder_Status   ='.Статус спецификации складских ордеров';
  //---- #3692
  VA_PRIKAZ_LANDOCS   ='Признак импорта из LANDOCS';

  //===================================
  // Plpor
  //------- #3843
  Va_PlPor_Akt_ticket ='.Билет (сопроводительный документ)';
  VA_PlPor_SS_Lsc     ='.Сотовая связь - лицевой счет' ; //#884

  //===================================
  // StatLog
  VA_StatLog_Comment  = '.Комментарий'; //#3830

  // ---------- свои нужды для фиксации даты создания документов
  VA_DOC_CreateDate = '.Создание документа. Дата';
  VA_DOC_CreateTime = '.Создание документа. Время';

  //== Атрибуты для SPINVTAB EM_51 ========================
  VA_spInv_MOL ='.МОЛ_ОСТАТКИ';// Мол откуда остатки
  VA_spInv_Podr='.ПОДРАЗДЕЛЕНИЕ_ОСТАТКИ';// Склад откуда остатки

  //=================================== 
  //Маркер для пометки записей каталога единиц измерения
  GP_GetEdIzm_Marker = 'GP_GETEDIZM_MARKER';

  //************************************************
  // Наименование разделов документов c правами  на события

  Rights_AktUslB_111_Header='Шапка';
  Rights_AktUslB_111_Spec  ='Спецификация';
  Rights_ClaimSZ_Header    ='Шапка';
  Rights_ClaimSZ_Spec      ='Спецификация';
  Rights_ConsPlan_Header   ='Шапка';
  Rights_ConsPlan_Spec     ='Спецификация';
  Rights_DepClaim_Header   ='Шапка';
  Rights_DepClaim_Spec     ='Спецификация';
  Rights_DepClaimEd_Header ='Шапка';
  Rights_DepClaimEd_Spec   ='Спецификация';
  Rights_DepClaimUsl_Header='Шапка';
  Rights_DepClaimUsl_Spec  ='Спецификация';
  Rights_KatParty_Header   ='Шапка';
  Rights_KatParty_Spec     ='Спецификация - Результаты входного контроля';
  Rights_ListXchDoc_Header ='Шапка';
  Rights_ListXchDoc_Spec   ='Спецификация';
  Rights_Lot_Header        ='Шапка';
  Rights_Lot_Spec          ='Спецификация';
  Rights_LotEd_Header      ='Шапка';
  Rights_LotEd_Spec        ='Спецификация';
  Rights_Nakl501_Header    ='Шапка';
  Rights_Nakl501_Spec      ='Спецификация';
  Rights_PlanBuy_Header    ='Шапка';
  Rights_PlanBuy_Spec      ='Спецификация';
  Rights_ProdNeed_Header   ='Шапка';
  Rights_ProdNeed_Spec     ='Спецификация';
  Rights_SopInMov_Header   ='Шапка - Основная информация';
  Rights_SopInMov_Spec     ='Спецификация';
  Rights_SoprDoc_Header    ='Шапка';
  Rights_SoprDoc_Spec      ='Спецификация';
  Rights_SoprDocB_Header   ='Шапка';
  Rights_SoprDocB_Spec     ='Спецификация';
  Rights_WorkPlan_Header   ='Шапка';
  Rights_WorkPlan_Spec     ='Спецификация';

  Rights_Peresort_Header   ='Шапка';
  Rights_Peresort_Spec_1   ='Спецификация - Списание';
  Rights_Peresort_Spec_2   ='Спецификация - Оприходование';

  Rights_Inv_Prih_Header   = 'Шапка';
  Rights_Inv_Prih_Spec     = 'Спецификация';

  Rights_AktUMC_Header   = 'Шапка';
  Rights_AktUMC_Spec     = 'Спецификация';

  Rights_SoprDsOut_Header   = 'Шапка';
  Rights_SoprDsOut_Spec     = 'Спецификация';

  Rights_Nakl502_Header   = 'Шапка';
  Rights_Nakl502_Spec     = 'Спецификация';

  Rights_AktSpis_Header   = 'Шапка';
  Rights_AktSpis_Spec     = 'Спецификация';

  //************************************************
  // Цвета
  // Цвет фона для ГосЗаказа. Раньше использовался ColorError
  ColBkGOSz = ColorSysAqua; // 15; //  'Шифр заказа' :[20],Protect,{font={BackColor=if(isGOZ(PlannedReceipts.cOrderCode),ColorError,0)}};


  //========= ТИПЫ ДОКУМЕНТОВ ДЛЯ GP_File_hRef ===========
  flRefTipDoc_JDDanger : word = 10;


  //======================================================

  UniDicType_DayOff = 4; //EM_Army, EM_DayOff
  //5??
  //56??
  //108??
  //109??
  UniDicType_SCUDHistory = 113; //coAccesHistType GpSCUD_PersEdit
  //114??
  //115??
  //116??
  //118??
  UniDicType_OKPD2ImportDeny = 121;
  UniDicType_OKPD2 = 122; //Algoritm_6971, EM_OKPD2_CUR_Limit, EM_OKPD2_DIC
  //123??
  UniDicType_Rpd158 = 158;//RepRpd158
  UniDicType_Rpd158Mail = 159;//RepRpd158_EmailSender
  UniDicType_SpisReasons = 200;//EM_Reasons_Dic
  UniDicType_NDSKoefs = 300;//GP_RuNDS_Coeff (80279 - не слито!)
  UniDicType_AktOcen501 = 501;//Import_03504
  UniDicType_AktOcen502 = 502;//Import_03504
  UniDicType_AktOcen632 = 632;//Import_03504
  UniDicType_AktDefect = 715;//ActDefect
  UniDicType_AktDefectRoles = 716;//ActDefectRoles
  UniDicType_Sebestoim = 1204; //Algoritm_6967
  UniDicType_AktOcen10007 = 10007;//Import_03504
  UniDicType_FpImpRel = 22375; //cofpImpRel GP_FpImpRelEdit, Gp_FinPlanRepBuilder

end;
