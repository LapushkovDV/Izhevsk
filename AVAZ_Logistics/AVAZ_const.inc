const

//
// СФО
//
 cgKAU_BrUch =  10166; // Аналитика Бригады Участки, новшество в СФО

 cgKAU_PointOut = 10121;

 VA_STO_KT = 'КТ';
 sKauKT45  = 'КТ45';

 cgKAU_StatCalc = 10062;
 sKAU_StatCalc = 'Инструмент и оснастка';

 VA_TARAM     =  'Ведется учет по тарным местам'; // имя внешнего атрибута <Ведется учет по тарным местам>   KatMC
 VA_KatPar    =  'Партия';//   ВА к SopORDB
 VA_SO_ROB    =  'Обувь1';
 VA_SO_ROD    =  'Одежда1';
 VA_SO_ROST   =  'Рост';
 VA_SO_BRIG   =  'Бригада';
 VA_SO_UCH    =  'Участок';
 VA_SO_Spec_rab   = 'Спецификация работника';
 VA_SO_Spec_norm  = 'Спецификация нормы';
 VA_KATMC_SCHUCH  = 'Счет учета ТМЦ';
 VA_KATMC_KE      = 'КЕ';
 VA_KATMC_PROIZV  = 'Производитель';
//
// Производственные спецификации
//
PS_SHAB_Instr = 'ПрС Инструмент';
PS_SHAB_Mat   = 'ПрС Материальная';
PS_SHAB_Main  = 'ПрС Основная';
//
// Качество
//
 //Каталог ввода норматива по внутреннему технологическому браку
 cgNormInTexBrak = 21055;

 cgProtocol_Quality = 21056; //Протокол для закрытия документов качества
 //Act_on_Quality 'Акт по качеству'

 cgAct_on_Quality = 21054; // Акт по качеству

 // Карта Брака
 cgKB_Doc     =  21053;

 cgKAU_VSob  =   10054;// Вид собственности

 cgKAU_Model  =  10105;//Задача №359  10090;// Модель автомобиля   //cgKAU_Model  =  20040;// Модель автомобиля

 cgKAU_VBrak  =  10066;// Вид брака          // cgKAU_VBrak  =  20038;// Вид брака

 cgKAU_ReshB  =  10070;// Принятое решение по КБ    // cgKAU_ReshB  =  20029;// Принятое решение по КБ

 cgKAU_VinB   =  10098; // Виновник брака   cgKAU_VinB   =  20026;

 cgKAU_Brig   =  10099; // 70; Бригада      cgKAU_Brig   =  20043;

 cgKAU_Uch    =  10100;  //Участок   cgKAU_Uch    =  20042;

 cgKAU_Stadiya    = 10071; // Стадия забракования    // cgKAU_Stadiya    = 20063; // Стадия забракования

 cgKAU_Defekt     = 10068; // Кодификатор дефектов    cgKAU_Defekt     = 20064; // Кодификатор дефектов

 cgKAU_Ranking    = 10069; // Комментарии для расчета Ranking  cgKAU_Ranking    = 20065; // Комментарии для расчета Ranking

 cgKAU_ZtrRekl    = 10067;// Затраты по обработке рекламации   cgKAU_ZtrRekl    = 20066;// Затраты по обработке рекламации

 cgKAU_ReshBrak   = cgKAU_ReshB; // 'это тоже самое что и 20062; // Принятое решение по браку

 cgKAU_KE         = 10105; // Калькулируемая единица
 
 cgKAU_ColorAvaz  = 10058; // Цвет автомобиля

 cgKAU_VINCode    = 5023; // Серийный номер

 cgKAU_Sostoyanie    = 10055;// Состояние детали


 VA_CALCED   =  'КЕ';
 VA_Q_N_TRAN = 'Номер транзитной накладной';
 VA_Q_D_TRAN = 'Дата транзитной накладной';
// 'Номер договора Грузоотправителя по транзитной схеме';

//
// Недвижимость
//

  VA_RLT_NAME    =  'Наименование';
  roRoomWorkPlace  =250;                // рабочие места к комнате
  roRoomWorkPlaceP =251;                // рабочие места к комнате - люди
  ORD_No_Protocol       = 256; // последний бит !
  ORD_No_DelControl     = 512; // не контролировать доступ к документам при удалении ордеров.

//
//  APOL
//

 VA_APOL_LINE_kol    =  'АПОЛ_ЛК_Кол/поз';
 VA_APOL_LINE_Temp   =  'АПОЛ_ЛК_Темп';
 //VA_APOL_RPT_step    =  'АПОЛ_РПТ_Шаг';
 APOL_FPL_TYPEALG             =  30;
 APOL_FPL_WSYSNUM             =  cgKB_Doc ;

 toInvTab   = 21301;     // Для связей между документами входящие\исходящие
 toSpInvTab = 21302;


end;
