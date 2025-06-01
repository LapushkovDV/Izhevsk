.LinkForm 'GP_REESBGPB' Prototype is 'REESBVIP'
.Group 'Реестры по перечислению в сбербанк'
.NameInList 'ЭМ.Реестр в Газпром Банк'
.var
  ExcelFileName:String
  NewPath:String
  NewFileName:String
  bool:Boolean
  Line:LongInt
  ii:Word
  lastBSlach:Word
  _sumUder : double;
.endvar
.create view SumVzisk
as select ZAROPER.SUMALIM
 from ZAROPER
 where ((comp(SPPLBANNREC) == ZAROPER.CPLPOR and
coSPPLBAN == ZAROPER.PRIZPR
))
;
.create view vSumUder
as select uder.sumud
 from uder, SPPLBAN
 where ((
                 comp(SPPLBANNREC) == SPPLBAN.NREC
             and SPPLBAN.NREC == ZAROPER.CPLPOR
             and coSPPLBAN == ZAROPER.PRIZPR
/*
  and year(dgettune('UP.DATOTCH')) == uder.yearn
 and month(dgettune('UP.DATOTCH')) == uder.mesn
                 and SPPLBAN.cLsch == uder.clsch
                           and 232 == uder.vidud (NoIndex)
                           */
))
;
.fields
  NewPath + NewFileName
.endfields
.{
.begin
  ExcelFileName := 'GP_ReestrToGazprom.xls';
  ExcelFileName := TranslatePath('%ClientStartPath%') + 'XLS_ЭМ\Staff\' + ExcelFileName ;


  if (not ExistFile(ExcelFileName))
  { message('Не найден шаблон файла "'+ExcelFileName+'"',error);
    exit;
  }

  if (not xlCreateExcelWithTemplate(ExcelFileName,true))
  { message('Ошибка создания нового файла в Excel"'+ExcelFileName+'".',error);
    exit;
  }

//  Путь для сохранения NewPath
  xlgetcellvalue(2,3, NewPath);
    if NewPath = ''
       NewPath := TranslatePath('%HOMEDRIVE%%HOMEPATH%') + 'Documents\'
    else
       NewPath := TranslatePath(NewPath);
end.

.{
.{
.begin
Line := longint(Npp) + 1;
  FIO := Replace(UpCase(trim(FIO)),'Ё','Е');
// LogStrToFile('gpgpgpgpgpgpgpgpg',Line + '  ' + Npp + '    FIO='+FIO);
  xlSetCellStringValue(Npp                      ,Line , 1,Line , 1); // Номер п/п
  xlSetCellStringValue('1337'                   ,Line , 2,Line , 2); // Код организации
  xlSetCellStringValue('Паспорт гражданина РФ'  ,Line , 3,Line , 3); // Тип паспорта
  xlSetCellStringValue(P_seria                  ,Line , 4,Line , 4); // Серия паспорта
  xlSetCellStringValue(''''+P_number            ,Line , 5,Line , 5); // Номер паспорта
  xlSetCellStringValue(ExtractWord(1,FIO,' ')   ,Line , 6,Line , 6); // Фамилия
  xlSetCellStringValue(ExtractWord(2,FIO,' ')   ,Line , 7,Line , 7); // Имя
  xlSetCellStringValue(ExtractWord(3,FIO,' ')   ,Line , 8,Line , 8); // Отчество
  xlSetCellStringValue(''''+Nom_lsh             ,Line , 9,Line , 9); // Номер карточного счёта Держателя

  xlSetCellStringValue(DoubleToStr(Round(Sum_Per,2),''''+'[|-]6666666666.88') ,Line ,10,Line ,10); // Сумма
/*
 if SumVzisk.getfirst ZAROPER = tsOK
 {
  if SumVzisk.ZAROPER.SUMALIM <> 0
   xlSetCellStringValue(DoubleToStr(Round(SumVzisk.ZAROPER.SUMALIM,2),''''+'[|-]6666666666.88') ,Line ,12,Line ,12); // Сумма
 }
*/
_sumUder := 0;
 vSumUder._loop Spplban
 vSumUder._loop zaroper
 {
   _sumUder := _sumUder + vSumUder.ZAROPER.SUMALIM
 }

  if _sumUder <> 0
   xlSetCellStringValue(DoubleToStr(Round(_sumUder,2),''''+'[|-]6666666666.88') ,Line ,12,Line ,12); // Сумма

end.
.}
.}
.if BRPAGE
.else
.end
.begin
    xlGetActiveWorkBookName(ExcelFileName); // Узнать текущее имя Excel-файла

//  Имя файла в формате  "VYYMMDDN.1337.xlsx" N - пусто для первого, затем A,B,C,D,E...
//    ReesDForm :=  STRTODATE(Cur_Date,'DD/MM/YYYY');
    NewFileName := DateToStr(Cur_Date,'VYYMMDD'); // 15/08/2018 Заказчик Нестерова текущая дата //  'V'+substr(ReesDForm,9,2) + substr(ReesDForm,4,2)+ substr(ReesDForm,1,2); // VYYMMDD

    if not ExistFile(NewPath + NewFileName + '.1337.xlsx')  NewFileName := NewFileName + '.1337.xlsx';
    else
    {
       ii := 65;
       While ExistFile(NewPath + NewFileName + CHR(ii) + '.1337.xlsx')
       ii++;

       NewFileName := NewFileName + CHR(ii) + '.1337.xlsx';

    }
    sSetAttr(16060,SBBANNREC,'Имя файла',NewFileName);
    xlSaveAsWorkBookByName(ExcelFileName,NewPath + NewFileName);

end.
  ^
.begin
Line++;
  xlSetCellStringValue('ИТОГО ЗАЧИСЛЕНИЙ:' + Npp        ,Line, 2,Line, 2);
  xlSetCellStringValue('ИТОГО СУММА:' +ItogSum_Reestr1  ,Line, 3,Line, 3);

  xlSaveWorkBookByName(NewPath+ NewFileName);
xlkillExcel;
end.
.}

.endform
