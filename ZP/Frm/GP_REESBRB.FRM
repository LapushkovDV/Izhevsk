/*************************************************************************************************\
* Наименование:  Реестр по перечислению для РосБанк                                               *
*                                                                                                 *
* Контур/Модуль: Заработная плата                                                                 *
* Примечание:    Link-форма печать реестра                                                        *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#800          |13/12/17|Шунин А.Г.               |НПО Энергомаш                    *
\*************************************************************************************************/
.Linkform 'GP_REESBRB' Prototype is 'reesBvip'
.nameinlist 'ЭМ.Реестр в РосБанк'
.var
ExcelFileName, NewPath, NewFileName: string;
  _sumUder : double;
bool: boolean;
Line, RegNum: longint;
ii, lastBSlach: longint;
.endvar
.group 'Реестры по перечислению в сбербанк'
.create view SumVzisk
as select ZAROPER.SUMALIM
 from ZAROPER
 where ((comp(SPPLBANNREC) == ZAROPER.CPLPOR and
coSPPLBAN == ZAROPER.PRIZPR
))
;
.create view vSumUder
as select uder.sumud
 from uder , SPPLBAN
 where ((
                 comp(SPPLBANNREC) == SPPLBAN.NREC
  and year(dgettune('UP.DATOTCH')) == uder.yearn
 and month(dgettune('UP.DATOTCH')) == uder.mesn
                 and SPPLBAN.cLsch == uder.clsch
                           and 232 == uder.vidud (NoIndex)
))
;
.{
.begin
  ExcelFileName := 'GP_ReestrToRosBank.xls';
  ExcelFileName := TranslatePath('%ClientStartPath%') + 'XLS_ЭМ\Staff\' + ExcelFileName ;


  if (not ExistFile(ExcelFileName))
  { message('Не найден шаблон файла "'+ExcelFileName+'"',error);
    exit;
  }

  if (not xlCreateExcelWithTemplate(ExcelFileName,true))
  { message('Ошибка создания нового файла в Excel"'+ExcelFileName+'".',error);
    exit;
  }

/*
    var path2: string;
    if xlOpenWorkBook(ExcelFileName) // открываем сам шаблон для хранения номера реестра в 13,4
    {
       xlGetActiveWorkBookName(path2);
       xlgetcellvalue(13,4, RegNum);
       RegNum := RegNum + 1 ;

       xlSetCellStringValue(RegNum         ,13 , 4, 13, 4);
       xlSaveWorkBookByName (path2);
       xlCloseWorkBookByName(path2);
    }

*/
  xlGetActiveWorkBookName(ExcelFileName);
//  Путь для сохранения NewPath
  xlgetcellvalue(2,3, NewPath);
  xlSetCellStringValue('' ,2 , 2, 2, 3); // Очистить ячейки с вспомогательной информацией про путь для сохранения

    if NewPath = ''
       NewPath := TranslatePath('%HOMEDRIVE%%HOMEPATH%') + 'Documents\'
    else
       NewPath := TranslatePath(NewPath);

    Line := 20;
end.

.{
.[h OverWrite
.]h
.{
.begin
Line++;
xlInsertRange(xlShiftDown, Line, 1, Line, 11);
  xlSetCellStringValue(UpCase(FIO )    ,Line , 1,Line , 1); // Фамилия, имя ,отчество
  xlSetCellStringValue(RTabN           ,Line , 2,Line , 2); // Табельный номер

  xlSetCellStringValue(Nom_lsh         ,Line , 4,Line , 4); // Счёт сотрудника
  if KodVal = comp(0)
  xlSetCellStringValue('RUB'           ,Line , 5,Line , 5); // Валюта
  xlSetCellStringValue(Sum_Per         ,Line , 6,Line , 6); // Сумма
//  xlSetCellStringValue(DoubleToStr(Sum_Per,'[|-]6666666666,88') ,Line , 6,Line , 6); // Сумма

if SumVzisk.getfirst ZAROPER = tsOK
 {
  if SumVzisk.ZAROPER.SUMALIM <> 0
  xlSetCellStringValue(DoubleToStr(Round(SumVzisk.ZAROPER.SUMALIM,2),''''+'[|-]6666666666.88'),Line ,10,Line ,10); // Сумма
 }

/*
_sumUder := 0;
 vSumUder._loop Spplban
 vSumUder._loop uder
 {
    _sumUder := _sumUder + vSumUder.uder.sumud
 }

  if _sumUder <> 0
   xlSetCellStringValue(DoubleToStr(Round(_sumUder,2),''''+'[|-]6666666666.88') ,Line ,10,Line ,10); // Сумма
*/
end.
.fields
ReesDForm
STRTODATE(ReesDForm,'DD/MM/YYYY')
.endfields
^  @@@@@@@@@@@@@@
.}
.[f OverWrite
.]f
.}
.if BrPage
.end
.begin
    ReesDForm := TRIM(ReesDForm);
    xlGetActiveWorkBookName(ExcelFileName);

// "F_РБDDMMYYYY_NN РосБанк.xls"
    NewFileName := DateToStr(Cur_Date,'РБDDMMYYYY'); // 15/08/2018 Заказчик Нестерова текущая дата, "F_" убрать


    ii := 1;
    While ExistFile(NewPath + NewFileName + '_'+ if(ii < 10, '0','') + string(ii) + '.xlsx')
     ii++;

//  NewFileName := NewFileName + '_'+ if(ii < 10, '0','') + string(ii) + '.xls';
    NewFileName := NewFileName + '_'+ if(ii < 10, '0','') + string(ii);
    sSetAttr(16060,SBBANNREC,'Имя файла',NewFileName);
//    xlSetCellStringValue(NewFileName,1, 1, 1, 1);

    NewFileName := NewFileName + '.xlsx';

    RegNum := string(ii);

    xlSaveAsWorkBookByName(ExcelFileName,NewPath + NewFileName);
end.
.fields
 NewPath + NewFileName
.endfields
  ^
.begin
  xlDeleteRange(xlShiftUp, 20, 1, 20, 11);
// Line++;
  xlSetCellStringValue(ItogSum_Reestr1  ,Line, 6,Line, 6);

  xlSetCellDATEValue(STRTODATE(ReesDForm,'DD/MM/YYYY') ,12 , 4, 12, 4); // Дата платежного реестра / Payroll list date

  xlSetCellStringValue(RegNum          ,13 , 4, 13, 4); // Номер платежного реестра / Payroll list number

  xlSetCellStringValue(PlPorNom        ,14 , 4, 14, 4); // К платёжному поручению № /  To the payment order #
  xlSetCellStringValue(''''+PlPorData       ,15 , 4, 15, 4); // От/ As of

  xlSaveWorkBookByName(NewPath + NewFileName);

xlkillExcel;
end.
.}

.endForm
