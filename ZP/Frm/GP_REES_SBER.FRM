/*************************************************************************************************\
* Наименование:  Excel-Файл для Сбер                                                              *
*                                                                                                 *
* Контур/Модуль: Заработная плата                                                                 *
* Примечание:    Link-форма печать реестра                                                        *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#000          |29/01/25|Dmitrii.Lapygin          |НПО Энергомаш                    *
\*************************************************************************************************/
.Linkform 'GP_REES_SBER' Prototype is 'reesBvip'
.nameinlist 'ЭМ. Excel-Файл для Сбер'
.var
ExcelFileName, NewPath, NewFileName: string;
  _sumUder : double;
bool: boolean;
Line, RegNum: longint;
ii, lastBSlach: longint;
sNumRees:string;

.endvar
.group 'Реестры по перечислению в сбербанк'
.{
.begin
  ExcelFileName := 'GP_Reestr_Sber.xlsx';
  ExcelFileName := TranslatePath('%ClientStartPath%') + 'XLS_ЭМ\ZP\' + ExcelFileName ;


  if (not ExistFile(ExcelFileName))
  { message('Не найден шаблон файла "'+ExcelFileName+'"',error);
    exit;
  }

  if (not xlCreateExcelWithTemplate(ExcelFileName,true))
  { message('Ошибка создания нового файла в Excel"'+ExcelFileName+'".',error);
    exit;
  }
  xlGetActiveWorkBookName(ExcelFileName);

    if NewPath = ''
       NewPath := TranslatePath('%HOMEDRIVE%%HOMEPATH%') + 'Documents\'
    else
       NewPath := TranslatePath(NewPath);

    Line := 3;

    sNumRees:= sGetAttr(16060,SBBANNREC,'Имя файла');
    sNumRees:= substr(sNumRees,Length(sNumRees)-8,2);

end.

.{
.[h OverWrite
.]h
.{
.begin
Line++;
xlInsertRange(xlShiftDown, Line, 1, Line, 9);
  xlSetCellStringValue(sNumRees ,Line , 1,Line , 1); // Номер Реестра
  xlSetCellStringValue(DateToStr(STRTODATE(ReesDForm,'DD/MM/YYYY'),'DD.MM.YYYY')       ,Line , 2,Line , 2); // Дата формирования
  xlSetCellStringValue(NPP                                     ,Line , 3,Line , 3); // Нпп
  xlSetCellStringValue(ExtractWord(1,FIO,' ')                  ,Line , 4,Line , 4); // Фамилия
  xlSetCellStringValue(ExtractWord(2,FIO,' ')                  ,Line , 5,Line , 5); // Имя
  xlSetCellStringValue(ExtractWord(3,FIO,' ')                  ,Line , 6,Line , 6); // Отчество
  xlSetCellStringValue(Sum_Per                                 ,Line , 7,Line , 7); // Сумма
  xlSetCellStringValue(0                                       ,Line , 8,Line , 8); // Кол записей
  xlSetCellStringValue(0                                       ,Line , 9,Line , 9); // Сумма Итого

end.
.}
.[f OverWrite
.]f
.}
.if BrPage
.end
.begin

  xlSetCellStringValue(Npp ,4, 8,Line, 8);
  xlSetCellStringValue(ItogSum_Reestr1  ,4, 9,Line, 9);
  Line:=Line+2;
  xlSetCellStringValue('ИТОГО : ' ,Line, 8,Line, 8);
  xlSetCellStringValue(ItogSum_Reestr1  ,Line, 9,Line, 9);

  Line:=Line+2;
  xlSetCellStringValue('Начальник отдела расчетов с персоналом' ,Line, 1,Line, 1);
  line++;
  xlSetCellStringValue('управления по бухгалтерскому сопровождению АО "НПО Энергомаш"' ,Line, 1,Line, 1);
  xlSetCellStringValue('Моисеева Е.И.' ,Line, 7,Line, 7);

  xlSetCellStringValue('АО НПО Энергомаш перечисление за '+DateToStr(STRTODATE(ReesDForm,'DD.MM.YYYY'),'month YYYY')+' ,реестр № '+sNumRees+' от '+STRTODATE(ReesDForm,'DD/MM/YYYY')   ,1, 1,1, 1);
xlkillExcel;
end.
.}
.endForm
