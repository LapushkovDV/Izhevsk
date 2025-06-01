! 1389 Протокол расчета среднего заработка
.linkform 'GP_Zarhelz101' prototype is 'zarhelz'
.NameInList 'ЭМ.Протокол cреднего (Excel)'
.group 'Standart'
.Group 'ForNeyav'
.f "NUL"
! 1389 Протокол расчета среднего заработка
#include Excel_fun.frn
#include GP_Zarhelz.Frn
.fields
 MyxlWrite(Substr(naim,19,80),rowcur,1,3)
.endfields
.declare
#include StrFunc.vih
.enddeclare
.var
  iStrFunc : StringFunctions;
.endvar
.Begin
  timelog_init;
  wKlNotSave:=True;// гасим автосохранение , поскольку есть макрос и ругается.
  MyXlInit('GP_Zarhelz.xlsm');
  rowcur:=0;
  //wexcel_num_0:=1;// гасим прочерки
  //InitWorkingTable;
  if not CL.Add(CS, 'IntInit') or not CL.Add(CS, 'WorkingTable')
  {
  }
  kl_first:=true;
 end.
.{
.begin
  if not kl_first
  { rowcur:=rowcur+1;
    XlRunMacro('New_Page'+chr(40)+string(rowcur)+chr(41))
  }
  kl_first:=false;
  rowcur:=rowcur+2;

  MyxlWrite('АО "НПО Энергомаш"',rowcur,3,3);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,7);
  xlSetFontSize(11, rowcur,1,rowcur,7);
 end.
		            АО ''НПО Энергомаш''
.begin
  rowcur++;
  xlSetFontSize(11, rowcur,1,rowcur,7);
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
end.
^
.begin
 rowcur++;
 MyxlWrite('Подразделение',rowcur,1,3);
 MyxlWrite('Табельный номер',rowcur,3,3);
 MyxlWrite('Фамилия И.О.',rowcur,5,3);
 xlMergeCells( rowcur, 1, rowcur, 2);
 xlMergeCells( rowcur, 3, rowcur, 4);
 xlMergeCells( rowcur, 5, rowcur, 6);
 MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,7);
  xlSetFontSize(11, rowcur,1,rowcur,7)

  end.
Подразделение		Табельный номер		Фамилия И.О.
.begin
  rowcur++;
  colcur:=1;
  xlSetFontSize(10, rowcur,1,rowcur,7)
end.
.fields
 MyxlWrite(Kod ,rowcur,1,3)
 MyxlWrite(chr(39)+TabN,rowcur,3,3)
 MyxlWrite(Fio ,rowcur,5,3)
.endfields
^		^		^
.begin
 xlMergeCells( rowcur, 3, rowcur, 4);
 xlMergeCells( rowcur, 1, rowcur, 2);
 xlMergeCells( rowcur, 5, rowcur, 6);
 MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
 end.
.{
.begin
v.delete all from GP_tZarHelZ;
Premi1 := 0;
PremiSum := 0;
PremiSum2 := 0;
 It_su1:=0; it_su2:=0;
a:= 0;
b:= 0;
VsegoZar := 0;
v.delete all from GP_tZarHelz;
c:= 0;
end.
.Fields
// Раб_ч
.endfields
.begin
 rowcur++;
 colcur:=1;
 MyxlWrite('Месяц' ,rowcur,colcur,3);
 MyxlWrite(Раб_ч ,rowcur,colcur,3);
 MyxlWrite('График' ,rowcur,colcur,3);
 MyxlWrite('Заработок' ,rowcur,colcur,3);
 MyxlWrite('Премия' ,rowcur,colcur,3);
 MyxlWrite('Коэффициент' ,rowcur,colcur,3);
 MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
 xlSetFontStyle(xlBold, rowcur,1,rowcur,7);
 xlSetFontSize(11, rowcur,1,rowcur,7)

 end.
!Месяц	^	График	Заработок	Премия	Коэффициент
.{
.begin
  InsTmp
  end.
#IFDEF _2018_08_12_
.if Print2Fond
.{
.}
.end
.if Print2Fond2
.end
#ENDIF
.}
.begin
 // вдруг расчет в часах
 // пока весь перерасче часов отрублен - это все решено в стандарте
 CalcItog_CH;
 end.
.{table 'v.GP_tZarHelZ' by v.GP_tZarHelZ._year, v.GP_tZarHelZ._mes;
.begin
 CalcPrem;
 end.
.fields
 MyxlWrite(chr(39)+string(v.GP_tZarHelZ._mes)+'/'+string(v.GP_tZarHelZ._year),rowcur,colcur,3)
 MyxlWrite(v.GP_tZarHelZ._chasf,rowcur,colcur,3)
 MyxlWrite(v.GP_tZarHelZ._chasGr,rowcur,colcur,3)
 MyxlWrite(v.GP_tZarHelZ._SredSumma,rowcur,colcur,2)
 MyxlWrite(doubletostr(round(Premi1+Premi2,2),'666666666666.88'),rowcur,colcur,2)
 MyxlWrite(v.GP_tZarHelZ._SredKoef,rowcur,colcur,2)
.endfields
.begin
  rowcur++;  colcur:=1;
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);

 end.
^	^	^	^	^	^
.}
.fields
 MyxlWrite(string(Sredchasf,0,0),rowcur,colcur,2)
 MyxlWrite(string(SredchasGr,0,0),rowcur,colcur,2)
 MyxlWrite(doubletostr(round(It_su1,2),'666666666666.88'),rowcur,colcur,2)//Summaw
 MyxlWrite(doubletostr(round(PremiSum+PremiSum2,2),'666666666666.88'),rowcur,colcur,2)
.endfields
.begin
 a:= It_su1;//Summaw;
 b:= PremiSum+PremiSum2;
 kl_Prem:=false;
 VsegoZar:= a +  b;
 skipIndex := 0;
 rowcur++;  colcur:=1;
 MyxlWrite('Всего',rowcur,colcur,3);
 MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
 xlSetFontStyle(xlBold, rowcur,1,rowcur,6);

end.

Всего 	^	^	^	^
.begin
  rowcur:=rowcur+2;
  colcur:=1;
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
  MyxlWrite('Всего заработок',rowcur,1,3);
  xlMergeCells( rowcur, 1, rowcur, 2);
  MyxlWrite('Средний дневной',rowcur,4,3);
  xlMergeCells( rowcur, 4, rowcur, 5);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,1);
  xlSetFontStyle(xlBold, rowcur,4,rowcur,4);
  end.
.fields
 MyxlWrite(doubletostr(round(VsegoZar,2),'666666666666.88'),rowcur,3,2)
 MyxlWrite(DaySM,rowcur,6,2)
.endfields

Всего заработок		^	Средний дневной		^
.begin
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
  rowcur:=rowcur+1;
  colcur:=1;
  MyxlWrite('С перерасчетом премии',rowcur,1,3);
  xlMergeCells( rowcur, 1, rowcur, 2);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,2);
  MyxlFrameCells(rowcur,1,rowcur,3,1 or 2 or 4 or 8 or 16 or 32);

 end.
.fields
 MyxlWrite(doubletostr(round(It_su1,2)
     +Round((PremiSum+PremiSum2)*double(_Sredchasf)/double(_SredchasGr),2),'666666666666.88'),rowcur,3,2)
.endfields
С перерасчетом премии		^
.{?Internal; GetVidOpl = 77;
.begin
  rowcur:=rowcur+1;
  colcur:=1;
  MyxlWrite('Период, в котором выплачен',rowcur,1,3);
  MyxlWrite('Среднемесячное количество рабочих дней',rowcur,4,3);
  MyxlWrite('Сумма оплаты',rowcur,5,3);
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
  xlMergeCells( rowcur, 1, rowcur, 3);
  xlMergeCells( rowcur, 5, rowcur, 6);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,6);
 end.
              Период		  Дни   	Сум. оплаты
.begin
  rowcur:=rowcur+1;
  colcur:=1;
  xlMergeCells( rowcur, 1, rowcur, 3);
  xlMergeCells( rowcur, 5, rowcur, 6);
  xlAlignCells(xlCenter, rowcur, 1, rowcur, 4);
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);

  var dn: date;
  dn := StrToDate(DataN,'DD/MM/YYYY');

  var rsltAlg30107: double;
  rsltAlg30107 := month(if(Date( 1, Month( Sub_Months( dn, 12 ) ), Year( Sub_Months( dn, 12 ) ) )<lsc_datpos,Date( 1, Month( lsc_datpos), Year( lsc_datpos)),Date( 1, Month( Sub_Months( dn, 12 ) ), Year( Sub_Months( dn, 12 ) ) )));

  var rsltAlg30108: double;
  rsltAlg30108 := year (if(Date( 1, Month( Sub_Months( dn, 12 ) ), Year( Sub_Months( dn, 12 ) ) )<lsc_datpos,Date( 1, Month( lsc_datpos), Year( lsc_datpos)),Date( 1, Month( Sub_Months( dn, 12 ) ), Year( Sub_Months( dn, 12 ) ) )));

  var rsltAlg30102: double;
  rsltAlg30102 := WT_GetDaysBwDatesEx( 2, 0, date(1, rsltAlg30107, rsltAlg30108), Date( Last_Day( Sub_Months( dn, 1 ) ), Month( Sub_Months( dn, 1 ) ), Year( Sub_Months( dn, 1 ) ) ),64);

  var rsltAlg30106: double;
  rsltAlg30106 := if(Months_Between(LSC_DATPOS, dn) < 12, Months_Between(LSC_DATPOS, dn), 12);

  rsltAlg30103 := round(rsltAlg30102 / rsltAlg30106, 2);

  //message('rsltAlg30107='+rsltAlg30107+'; rsltAlg30108='+rsltAlg30108+'; rsltAlg30102='+rsltAlg30102+'; rsltAlg30106='+rsltAlg30106+'; rsltAlg30103='+rsltAlg30103);
 end.
.fields
 MyxlWrite(''''+DateToStr(StrToDate(Datan,'DD/MM/YYYY'),'MM.YYYY'),rowcur,1,3)
 MyxlWrite(rsltAlg30103,rowcur,4,3)
 MyxlWrite(string(GetPrVidOpl77Sum,0,2),rowcur,5,3)
.endfields
               ^ ^ ^
.}
.{?Internal; GetVidOpl <> 77;
.begin
  rowcur:=rowcur+1;
  colcur:=1;
  MyxlWrite('Период',rowcur,1,3);
  MyxlWrite('Дата назнач.',rowcur,3,3);
  MyxlWrite('Дни' ,rowcur,4,3);
  MyxlWrite('Часы',rowcur,5,3);
  MyxlWrite('Сум. оплаты',rowcur,6,3);
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
  xlMergeCells( rowcur, 1, rowcur, 2);
  xlSetFontStyle(xlBold, rowcur,1,rowcur,5);
 end.
              Период		Дата назнач.	Дни	Часы	Сум. оплаты
.begin
  rowcur:=rowcur+1;
  colcur:=1;
  MyxlFrameCells(rowcur,1,rowcur,6,1 or 2 or 4 or 8 or 16 or 32);
 end.
.fields
 MyxlWrite('С '+Datan,rowcur,colcur,3)
 MyxlWrite('По '+Datok,rowcur,colcur,3)
 MyxlWrite(chr(39)+DatZ,rowcur,colcur,3)
 MyxlWrite(KolDn,rowcur,colcur,3)
 MyxlWrite(Chasu,rowcur,colcur,3)
 MyxlWrite(string(round(Double(Summa)*GetProc,2),0,2),rowcur,colcur,3)
.endfields
С ^ По ^		^	^	^	^
.}
.}
.fields
 MyxlWrite(/*GlavBuh*/ 'Бухгалтер  ________________ / ' + iStrFunc.GetFIO_ZAR(sGetTune('USER.FIO')),rowcur,1,3)
 MyxlWrite(DateToStr(/*DataVipl*/ Cur_Date(), 'DD.MM.YYYY'),rowcur,1,3)
.endfields
.begin  rowcur:=rowcur+2; end.
^
.begin  rowcur:=rowcur+2; end.
^
! Макросом добавим
!.{while skipindex < (35-c)
!.begin
!  skipindex++;
!  rowcur++;
!end.
!.}
.}
.begin
  //MyxlWriteMatrix(rowcur) ;
  //xlFreeMatrix ;
  //xlWrapText(headerstrcount+1,1,rowcur,44)
  //MyxlFrameCells(headerstrcount,1,rowcur,44,1 or 2 or 4 or 8 or 16 or 32);
  matrixkl:=false;
  //-----------
  //MyXlEnd;
  xlKillExcel;
  //-----------
  zarhelz.Fexit;
  //DoneWorkingTable;
end.
.endform
