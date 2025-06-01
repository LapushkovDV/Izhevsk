/*************************************************************************************************\
* Наименование:  Разработка печатной формы                                                        *
*                      Анализ обеспеченности на период мобилизации и военное время трудовыми      * 
*                      ресурсами (руководителями, специалистами, квалифицированными рабочими и    * 
*                      служащими) из числа ГПЗ"                                                   *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Управление персоналом -> Отчеты -> Отчеты пользователя -> Ведение картотеки ->   *
*                    Анализ обеспеченности на период мобилизации и военное время трудовыми        *
*                    ресурсами (руководителями, специалистами, квалифицированными рабочими и      *
*                    служащими) из числа ГПЗ                                                      *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#398          |12/08/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/

!.linkform 'VoinReport' prototype is F6Reserv2010
.linkform 'F6Reserv2010_EM_04' prototype is F6Reserv2010
.nameinlist 'ЭМ.Сведения об обеспеченности трудовыми ресурсами на период мобилизации (форма 19)'
.F 'NUL'
.declare
#include xlReport.vih
#include History.vih
.endDeclare
.group 'Отчет по забронированным военнообязанным в формате Word от 30.12.2010 г '
.var
  dm_Procent: array[1..4] of double;
  dDateFiltr:date; // дата отбора в фильтре
.endvar
.create view s
var 
  ID
, sXLSFileName
, SXLTFileName       : string;

  pxlrepsetup        : xlrepsetup;
  pXL                : XlsRepBuilder;
;
.Function MyDouble(w:string):string;
begin
 Result:=0;
 Result:=double(Replace(w,'''',''));
 end.
.procedure PrintLine(name : string; level : word;
     dSum1, dSum2, dSum3, dSum4, dSum5, dSum6
   , dSum7, dSum8, dSum9, dSum10, dSum11, dSum12 : double;
    Procent:double=0
   );
begin
//-01 Всего работающих
//-Из численности всего работающих ГПЗ
//02 --всего	в том числе
//03 ---офицеров
//04 ---прапорщиков, мичманов, сер-жантов, старшин, солдат и матросов
//05 --Численность прапорщиков, мичманов, сер-жантов, старшин, солдат и матросов запаса, ограниченно годных к воен-ной службе
//
//= Из численности ГПЗ забронировано
//06 -- всего	в том числе
//07 --- офицеров
//08 --- прапорщиков, мич┐манов, сер-жантов, старшин, солдат и матросов
//09 Численность неза-брониро-ванных ГПЗ, не имеющих мобилиза-ционных предписа┐ний
//10 Численность ГПЗ, имеющих мобилизацион-ные предписания
//11 Из численности всего работаю-щих - граждан, подлежащих призыву на военную службу
//12 Примечание

  pXL.ClearTblBuffer;
  pXL.SetTblStringFldValue('Level'  , level);
  pXL.SetTblStringFldValue('Name'  , name);
  pXL.SetTblNumberFldValue('dSum1' , dSum1);
  pXL.SetTblNumberFldValue('dSum2' , dSum2);
  pXL.SetTblNumberFldValue('dSum3' , dSum3);
  pXL.SetTblNumberFldValue('dSum4' , dSum4);
  pXL.SetTblNumberFldValue('dSum5' , dSum5);
  pXL.SetTblNumberFldValue('dSum6' , dSum6);
  pXL.SetTblNumberFldValue('dSum7' , dSum7);
  pXL.SetTblNumberFldValue('dSum8' , dSum8);
  pXL.SetTblNumberFldValue('dSum9' , dSum9);
  pXL.SetTblNumberFldValue('dSum10', dSum10);
//призывники + надо добавить 2-6 (3-4 столбец отчета)
  pXL.SetTblNumberFldValue('dSum11', dSum11);// +dSum2-dSum6);
  pXL.SetTblNumberFldValue('dSum12', dSum12);
  pXL.SetTblNumberFldValue('Procent',dSum1*Procent/100);
  pXL.InsTblRow;
end.

.begin
  var ii:word;
  for(ii:=1;ii<=4;ii++)
    if (NOT ReadMyDsk(dm_Procent[ii], 'F6ReservReport2010_Procent'+String(ii), True))
        dm_Procent[ii]:=0.8;
  if not ReadMyDsk (dDateFiltr, 'CommonFiltr_dDateFiltr',false)
     dDateFiltr:=Cur_Date;
  var oHistory : iHistory;
  ID := 'F6Reserv2010_EM_04';
  //if (NOT ReadMyDsk(sXLTFileName,ID,true))
  set sXLTFileName := TranslatePath('%ClientStartPath%') + 'XLS_ЭМ\Staff\' + ID + '.xlt';

  //StartNewVisual(vtIndicatorVisual, vfTimer + vfBreak + vfConfirm, 'Идет формирование отчета', RecordsInTable(#tmpOvertime));

  //Set sXLSFileName := pXL.CreateXLT(sXLTFileName, True);
  Set sXLSFileName := pXL.CreateReport(sXLTFileName, True);

  if not pxlrepsetup.checkParam(0, ID, sXLTFileName)
   {
     Runinterface('XlRepSetup', 1, ID, sXLTFileName);
     ReadMyDsk(sXLTFileName, ID, true);
   }

  pXL.CreateVar(sXLSFileName);
  pXL.SetStringVar('CFH'           , CommonFormHeader);

  var sUtve    : string; sUtve    := '';
  var sUtveDol : string; sUtveDol := '';

  ReadMyDsk(sUtve      , 'F6ReservReport2010_sUtve'   , true);
  ReadMyDsk(sUtveDol   , 'F6ReservReport2010_sUtveDol', true);


  var sNameCity : string; sNameCity := '';
  pXL.SetStringVar('org'           , oHistory.sGetField(coKatOrg, coGetTune('MyOrg'), 'REP.KATORGNAME', Cur_Date));
  pXL.SetStringVar('Date'          , DateToStr(dDateFiltr, 'DD mon YYYY'));
  pXL.SetStringVar('sUtve'         , sUtve   );
  pXL.SetStringVar('sUtveDol'      , sUtveDol);
  pXL.PublishVar;  // Excel

  pXL.CreateTbls(sXLSFileName);
  pXL.CreateTbl('Report');
  pXL.CreateTblFld('Level');
  pXL.CreateTblFld('Name');
  pXL.CreateTblFld('dSum1');
  pXL.CreateTblFld('dSum2');
  pXL.CreateTblFld('dSum3');
  pXL.CreateTblFld('dSum4');
  pXL.CreateTblFld('dSum5');
  pXL.CreateTblFld('dSum6');
  pXL.CreateTblFld('dSum7');
  pXL.CreateTblFld('dSum8');
  pXL.CreateTblFld('dSum9');
  pXL.CreateTblFld('dSum10');     //=Gal_TblSheet!$A$1:$N$2
  pXL.CreateTblFld('dSum11');
  pXL.CreateTblFld('dSum12');
  pXL.CreateTblFld('Procent');

  PrintLine('Руководители'                                                  , 0, double(D_1_1),  double(D_1_2),  double(D_1_3),  double(D_1_4),  double(D_1_5),  double(D_1_6),  double(D_1_7),  double(D_1_8),  double(D_1_9),  double(D_1_10),  double(D_1_11),  double(D_1_12),dm_Procent[1] );
  PrintLine('Специалисты'                                                   , 0, Mydouble(D_2_1),  Mydouble(D_2_2),  Mydouble(D_2_3),  Mydouble(D_2_4),  Mydouble(D_2_5),  Mydouble(D_2_6),  Mydouble(D_2_7),  Mydouble(D_2_8),  Mydouble(D_2_9),  Mydouble(D_2_10),  Mydouble(D_2_11),  Mydouble(D_2_12),dm_Procent[2] );
  //PrintLine('А. Сельское хозяйство'                                         , 0, Mydouble(D_3_1),  Mydouble(D_3_2),  Mydouble(D_3_3),  Mydouble(D_3_4),  Mydouble(D_3_5),  Mydouble(D_3_6),  Mydouble(D_3_7),  Mydouble(D_3_8),  Mydouble(D_3_9),  Mydouble(D_3_10),  Mydouble(D_3_11),  Mydouble(D_3_12) );
  //PrintLine('С. Добыча полезных ископаемых'                                 , 0, Mydouble(D_4_1),  Mydouble(D_4_2),  Mydouble(D_4_3),  Mydouble(D_4_4),  Mydouble(D_4_5),  Mydouble(D_4_6),  Mydouble(D_4_7),  Mydouble(D_4_8),  Mydouble(D_4_9),  Mydouble(D_4_10),  Mydouble(D_4_11),  Mydouble(D_4_12) );
  //PrintLine('D. Обрабатывающие производства'                                , 0, Mydouble(D_5_1),  Mydouble(D_5_2),  Mydouble(D_5_3),  Mydouble(D_5_4),  Mydouble(D_5_5),  Mydouble(D_5_6),  Mydouble(D_5_7),  Mydouble(D_5_8),  Mydouble(D_5_9),  Mydouble(D_5_10),  Mydouble(D_5_11),  Mydouble(D_5_12) );
  //PrintLine('Е. Производство и распределение электроэнергии, газа и воды'   , 0, Mydouble(D_6_1),  Mydouble(D_6_2),  Mydouble(D_6_3),  Mydouble(D_6_4),  Mydouble(D_6_5),  Mydouble(D_6_6),  Mydouble(D_6_7),  Mydouble(D_6_8),  Mydouble(D_6_9),  Mydouble(D_6_10),  Mydouble(D_6_11),  Mydouble(D_6_12) );
  //PrintLine('F. Строительство'                                              , 0, Mydouble(D_7_1),  Mydouble(D_7_2),  Mydouble(D_7_3),  Mydouble(D_7_4),  Mydouble(D_7_5),  Mydouble(D_7_6),  Mydouble(D_7_7),  Mydouble(D_7_8),  Mydouble(D_7_9),  Mydouble(D_7_10),  Mydouble(D_7_11),  Mydouble(D_7_12) );
  //PrintLine('I. Транспорт и связь'                                          , 0, Mydouble(D_8_1),  Mydouble(D_8_2),  Mydouble(D_8_3),  Mydouble(D_8_4),  Mydouble(D_8_5),  Mydouble(D_8_6),  Mydouble(D_8_7),  Mydouble(D_8_8),  Mydouble(D_8_9),  Mydouble(D_8_10),  Mydouble(D_8_11),  Mydouble(D_8_12) );
  //PrintLine('М. Образование'                                                , 0, Mydouble(D_9_1),  Mydouble(D_9_2),  Mydouble(D_9_3),  Mydouble(D_9_4),  Mydouble(D_9_5),  Mydouble(D_9_6),  Mydouble(D_9_7),  Mydouble(D_9_8),  Mydouble(D_9_9),  Mydouble(D_9_10),  Mydouble(D_9_11),  Mydouble(D_9_12) );
  //PrintLine('N. Здравоохранение и предоставление социальных услуг'          , 0, Mydouble(D_10_1), Mydouble(D_10_2), Mydouble(D_10_3), Mydouble(D_10_4), Mydouble(D_10_5), Mydouble(D_10_6), Mydouble(D_10_7), Mydouble(D_10_8), Mydouble(D_10_9), Mydouble(D_10_10), Mydouble(D_10_11), Mydouble(D_10_12));
  //PrintLine('Прочие виды экономической деятельности'                        , 0, Mydouble(D_11_1), Mydouble(D_11_2), Mydouble(D_11_3), Mydouble(D_11_4), Mydouble(D_11_5), Mydouble(D_11_6), Mydouble(D_11_7), Mydouble(D_11_8), Mydouble(D_11_9), Mydouble(D_11_10), Mydouble(D_11_11), Mydouble(D_11_12));
  PrintLine('Другие служащие'                                               , 0, Mydouble(D_12_1), Mydouble(D_12_2), Mydouble(D_12_3), Mydouble(D_12_4), Mydouble(D_12_5), Mydouble(D_12_6), Mydouble(D_12_7), Mydouble(D_12_8), Mydouble(D_12_9), Mydouble(D_12_10), Mydouble(D_12_11), Mydouble(D_12_12),dm_Procent[3]);
  PrintLine('Рабочие:'                                                       , 0, Mydouble(D_13_1), Mydouble(D_13_2), Mydouble(D_13_3), Mydouble(D_13_4), Mydouble(D_13_5), Mydouble(D_13_6), Mydouble(D_13_7), Mydouble(D_13_8), Mydouble(D_13_9), Mydouble(D_13_10), Mydouble(D_13_11), Mydouble(D_13_12),dm_Procent[4]);
  //PrintLine('в том числе: имеющие тарифные разряды'                         , 0, Mydouble(D_14_1), Mydouble(D_14_2), Mydouble(D_14_3), Mydouble(D_14_4), Mydouble(D_14_5), Mydouble(D_14_6), Mydouble(D_14_7), Mydouble(D_14_8), Mydouble(D_14_9), Mydouble(D_14_10), Mydouble(D_14_11), Mydouble(D_14_12));
  //PrintLine('не имеющие тарифных разрядов'                                  , 0, Mydouble(D_15_1), Mydouble(D_15_2), Mydouble(D_15_3), Mydouble(D_15_4), Mydouble(D_15_5), Mydouble(D_15_6), Mydouble(D_15_7), Mydouble(D_15_8), Mydouble(D_15_9), Mydouble(D_15_10), Mydouble(D_15_11), Mydouble(D_15_12));
  //PrintLine('сельскохозяйственного производства'                            , 0, Mydouble(D_16_1), Mydouble(D_16_2), Mydouble(D_16_3), Mydouble(D_16_4), Mydouble(D_16_5), Mydouble(D_16_6), Mydouble(D_16_7), Mydouble(D_16_8), Mydouble(D_16_9), Mydouble(D_16_10), Mydouble(D_16_11), Mydouble(D_16_12));
  //PrintLine('локомотивных бригад'                                           , 0, Mydouble(D_17_1), Mydouble(D_17_2), Mydouble(D_17_3), Mydouble(D_17_4), Mydouble(D_17_5), Mydouble(D_17_6), Mydouble(D_17_7), Mydouble(D_17_8), Mydouble(D_17_9), Mydouble(D_17_10), Mydouble(D_17_11), Mydouble(D_17_12));
  PrintLine('из них водители'                                               , 1, Mydouble(D_18_1), Mydouble(D_18_2), Mydouble(D_18_3), Mydouble(D_18_4), Mydouble(D_18_5), Mydouble(D_18_6), Mydouble(D_18_7), Mydouble(D_18_8), Mydouble(D_18_9), Mydouble(D_18_10), Mydouble(D_18_11), Mydouble(D_18_12),dm_Procent[4]);
  //PrintLine('трактористы'                                                   , 0, Mydouble(D_18_1), Mydouble(D_18_2), Mydouble(D_18_3), Mydouble(D_18_4), Mydouble(D_18_5), Mydouble(D_18_6), Mydouble(D_18_7), Mydouble(D_18_8), Mydouble(D_18_9), Mydouble(D_18_10), Mydouble(D_18_11), Mydouble(D_18_12));
  //PrintLine('Из численности руководителей, специалистов и рабочих:'         , 0, Mydouble(D_19_1), Mydouble(D_19_2), Mydouble(D_19_3), Mydouble(D_19_4), Mydouble(D_19_5), Mydouble(D_19_6), Mydouble(D_19_7), Mydouble(D_19_8), Mydouble(D_19_9), Mydouble(D_19_10), Mydouble(D_19_11), Mydouble(D_19_12));
  //PrintLine('летно-подъемный состав'                                        , 0, Mydouble(D_20_1), Mydouble(D_20_2), Mydouble(D_20_3), Mydouble(D_20_4), Mydouble(D_20_5), Mydouble(D_20_6), Mydouble(D_20_7), Mydouble(D_20_8), Mydouble(D_20_9), Mydouble(D_20_10), Mydouble(D_20_11), Mydouble(D_20_12));
  //PrintLine('плавающий состав'                                              , 0, Mydouble(D_21_1), Mydouble(D_21_2), Mydouble(D_21_3), Mydouble(D_21_4), Mydouble(D_21_5), Mydouble(D_21_6), Mydouble(D_21_7), Mydouble(D_21_8), Mydouble(D_21_9), Mydouble(D_21_10), Mydouble(D_21_11), Mydouble(D_21_12));
  //PrintLine('Итого (сумма строк 1+2+12+13)'                                 , 0, Mydouble(D_22_1), Mydouble(D_22_2), Mydouble(D_22_3), Mydouble(D_22_4), Mydouble(D_22_5), Mydouble(D_22_6), Mydouble(D_22_7), Mydouble(D_22_8), Mydouble(D_22_9), Mydouble(D_22_10), Mydouble(D_22_11), Mydouble(D_22_12));
  //PrintLine('По небронируемым организациям'                                 , 0, Mydouble(D_23_1), Mydouble(D_23_2), Mydouble(D_23_3), Mydouble(D_23_4), Mydouble(D_23_5), Mydouble(D_23_6), Mydouble(D_23_7), Mydouble(D_23_8), Mydouble(D_23_9), Mydouble(D_23_10), Mydouble(D_23_11), Mydouble(D_23_12));
  //PrintLine('Всего (сумма строк 22+23)'                                     , 0, Mydouble(D_24_1), Mydouble(D_24_2), Mydouble(D_24_3), Mydouble(D_24_4), Mydouble(D_24_5), Mydouble(D_24_6), Mydouble(D_24_7), Mydouble(D_24_8), Mydouble(D_24_9), Mydouble(D_24_10), Mydouble(D_24_11), Mydouble(D_24_12));

  pXL.PublishTbl('Report');

  pXL.LoadReport(sXLSFileName);
  pXL.DisConnectExcel;

end.
.endform
 
