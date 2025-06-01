/*************************************************************************************************\
* Наименование:  Разработка отчетной формы                                                        *
*                "Табель учета рабочего времени Т-12"                                             *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Учет времени->Табели учета рабочего времени->Типовая форма Т-13                  *
*                                                           ->Табель учета рабочего времени Т-12  *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#166          |20/04/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/
.linkform 'GP_TabWorkTimeT12_EM' prototype is 'TabT13' // Есть клон TabWorkTimeT12_2.frm отличие только в RunInterface(iEMTabWorkTime, sPodr, dBeg, dEnd,1);
.Group 'Т-13' 
.nameinlist'*ЭМ.Табель учета рабочего времени Т-12 НПО Энергомаш'
.F 'NUL'
.var
  marker: longint
  sPodr : string;
  dBeg
, dEnd  : date;
  prevLstabNrec: comp;
  boSeparateTablePrinting:boolean;
  logfilenm:string;
  logfilekl:boolean;
  dPeriodBeginning, dPeriodEnding:date;
  pCex_Old,pAppoint_Old:comp;
  pCex_New,pAppoint_New:comp;
  DolgOldEm : string;
.endvar
.create view s
var
 pLsTab:comp;
 pPerexod:comp;
as select
 Perexod.NRec
from
  Perexod
//, ContDoc
, tmpTabWorkTime
, LStab
, ContDoc ContDoc2
, Catalogs CatRepEm
where 
((
    pPerexod        == Perexod.NRec
//AND Perexod.cpodr            == ContDoc.NRec

AND pLsTab          == LStab.NRec
AND LStab.TPERSON            == ContDoc2.Person
AND 92                       == ContDoc2.TYPEOPER
))
;
//*****************************************************************************
.Procedure MyLog(w:string);
begin
 //if length(logfilenm)>0
 if logfilekl
   LogStrToFile(logfilenm,w);
 end.

//*****************************************************************************
.Function TryDouble(value : string) : double;
begin
  var res : double; res := 0;

  _try
  {
     res := double(value);
  }
  _except else
  {
  }

  result := res;
end.

//*****************************************************************************
.Function GetNormDate(dt : string) : string;
begin
  result := dt;

  if (Trim(dt) = '')
    Exit;

  var firstSymbol : string; firstSymbol := SubStr(dt, 1, 1);
  if (InStr(firstSymbol, 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯABCDEFGHIKLMNOPQRSTVXYZ') <> 0)
    exit;

  var pref : string; pref := '';
  var n    : word;

  n := InStr('.', dt);
  if (n <> 0)
    {
      pref := '.';
    }
  else
    {
      n := InStr(',', dt);
      if (n <> 0)
        {
          pref := ',';
        }
      else
        {
          n := InStr(' ', dt);
          if (n <> 0)
            {
              pref := ' ';
            }
          else
            if (InStr(':', dt) = 0 AND InStr('-', dt) = 0)
              {
                result := dt + ':00';
                Exit;
              }
        }
    }

  if (n > 0)
    {
      var h : string; h := SubStr(dt, 1, n -1);
      var m : string; m := SubStr(dt, n + 1, length(dt) - n);

      result := h + ':' +  string(word(60*(double('0.' + m))));
    }
end.

.procedure gp_add_day(incValue: string; var rsltValue: string);//string[20]
begin
  incValue:=trim(incValue);
  rsltValue:=trim(rsltValue);
  rsltValue := replace(rsltValue, 'X', '') + replace(incValue, 'X', '');
  rsltValue := if(rsltValue = '', 'X', rsltValue);
  if logfilekl MyLog('gp_add_day rsltValue='+rsltValue+', incValue='+incValue);
end.

.procedure gp_add_hour(incValue: string; var rsltValue: string);//string[20]
begin
  incValue:=trim(incValue);
  rsltValue:=trim(rsltValue);
  var _boTime:boolean;
  _boTime:=false;
  if rsltValue = '' then rsltValue := '0';
  if incValue = '' then incValue := '0';

  //rsltValue := if(Str2Time(rsltValue) + Double(incValue) = 0, '', Time2StrDef(Str2Time(rsltValue) + Double(incValue)));
  /*
  if pos(':',incValue)>0 or pos(':',rsltValue)>0
  { _boTime:=true;
    var _Time1:time;
    _Time1:=ZeroTime;
    _Time1:=Str2Time(rsltValue,'HH:MM');
    var _Time2:time;
  }
  else
  */
    rsltValue := if(Str2Time(rsltValue) + Double(incValue) = 0, '', Time2StrDef(Str2Time(rsltValue) + Double(incValue)));
  if logfilekl MyLog('gp_add_hour rsltValue='+rsltValue+', incValue='+incValue);

end.

.procedure gp_add_dbl(incValue: string; var rsltValue: string);
begin
  rsltValue := String(double(rsltValue) + double(incValue), 0, 0);
end.
!------------------------------
! заполнить маску дней
! вызвать после заполнения основного буфера
!------------------------------
.Procedure Fill_tmpTabWorkTime;
begin
  var jj:word;
  For(jj:=1;jj<=31;jj++)
  { //if not s.tmpTabWorkTime.MaskDay[jj]
    // s.tmpTabWorkTime.MaskDay[jj]:=WT_IsDayEnabled(jj);
    if not s.tmpTabWorkTime.MaskDay[jj]
    { if jj<16
      { if s.tmpTabWorkTime.day1[jj]<>'X'
          s.tmpTabWorkTime.MaskDay[jj]:=true;
      }
      else
      { if s.tmpTabWorkTime.day2[jj-15]<>'X'
          s.tmpTabWorkTime.MaskDay[jj]:=true;
      }
    }
  }
 end.
!-----------------------------------------
! Начало формирования
!-----------------------------------------
.begin
   sPodr  := Podr;
   dBeg   := StrToDate(PeriodBeginning, 'DD/MM/YYYY');
   dEnd   := StrToDate(PeriodEnding   , 'DD/MM/YYYY');
   s.delete all tmpTabWorkTime;
   logfilenm:='';
   logfilekl:=false;
   boSeparateTablePrinting:=false;
   if not ReadMyDsk(boSeparateTablePrinting, 'PrintT13_SeparateTablePrinting', false)
      boSeparateTablePrinting := true;

   if pr_CurUserAdmin
   { logfilenm:=GetStringParameter('Files','OutputFilesDirectory',0)
           +'\!TabWorkTime_fill.log' ;
      if fileexist(logfilenm) deletefile(logfilenm);
     logfilekl:=true;
   }
  MyLog('SeparateTablePrinting='+string(boSeparateTablePrinting)) ;
  pCex_Old:=0; pAppoint_Old:=0;
  pCex_New:=0; pAppoint_New:=0;

end.
!-----------------------------------
.if SHOW_ANALITIK
.else
.end   
.{CheckEnter 
.{
.{
.begin
  //определяем, создаем новую строку или объединяем данные с предыдущей
  var isNewRec: boolean;
  isNewRec := true;
  set  s.pLsTab  :=comp(LSTABNREC);
  set  s.pPerexod:=comp(PEREXODNREC);
  if logfilekl
  { MyLog('0001 LSTABNREC= '+string(LSTABNREC,0,0)
    +', PEREXODNREC='+string(PEREXODNREC,0,0)
    +', isNewRec='+string(isNewRec)
    +', TabN='+TabN
    +', FIO='+Fio
    );
    MyLog('s.pLsTab='+String(s.pLsTab,0,0)
    +', s.pPerexod='+String(s.pPerexod,0,0)
    );
  }

  if s.getfirst LsTab=tsok
  { pCex_New:=s.LsTab.cex;
    pAppoint_New:=s.LsTab.cAppoint;
    if logfilekl
      Mylog('Нашли табель '+string(pCex_New,0,0)+','+string(boSeparateTablePrinting));
  }
  else
  { if logfilekl
      Mylog('Не Нашли табель '+string(pCex_New,0,0)+','+string(boSeparateTablePrinting));


  }
  if s.pPerexod<>0
  { if s.getfirst perexod = tsOk
      { pCex_New:=s.perexod.cexp;
        pAppoint_New:=s.perexod.cAppoint;
        if logfilekl
          Mylog('Нашли переход '+string(pCex_New,0,0));
      }
      else
      { if logfilekl
          Mylog('Не Нашли переход '+string(pCex_New,0,0)+','+string(boSeparateTablePrinting));


      }
  }

  if boSeparateTablePrinting
  { if logfilekl
       Mylog('SeparateTablePrinting '+string(pCex_New,0,0));

    if prevLstabNrec = comp(LSTABNREC)
      if pCex_Old=pCex_New
      and pAppoint_Old=pAppoint_New
      { isNewRec:=false;
      }
  } //if SeparateTablePrinting

  if logfilekl
  { MyLog('002 LSTABNREC= '+string(LSTABNREC,0,0)
    +', PEREXODNREC='+string(PEREXODNREC,0,0)
    +', isNewRec='+string(isNewRec)
    +', TabN='+TabN
    +', FIO='+Fio
    );
    MyLog('Cex='+String(pCex_Old,0,0)+'/'+String(pCex_New,0,0)
    +', Appoint='+String(pAppoint_Old,0,0)+'/'+String(pAppoint_New,0,0)

    );
  }

  pCex_Old    := pCex_New;
  pAppoint_Old:= pAppoint_New;

  prevLstabNrec := comp(LSTABNREC);

  if (getFirst FastFirstRow CatRepEm where ((pAppoint_New == CatRepEm.nRec)) = tsOk) then
  {
    DolgOldEm := CatRepEm.Name;
  }

  //message(if(IsNewRec,'true','false') + ' ' + prevLstabNrec + ' ' + lstabnrec + '|' + s.perexod.nrec + '|' + s.contdoc.nrec);
  if isNewRec
  {
    s.ClearBuffer(s.tntmpTabWorkTime);
    set s.tmpTabWorkTime.npp         := Npp;
    set s.tmpTabWorkTime.sFIOFull    := FIO;
    set s.tmpTabWorkTime.sFIO        := LogicSubStr(GetSurnameWithInitials(FIO), 22, false, false);
    if (boSeparateTablePrinting) then
    {

      set s.tmpTabWorkTime.sDol1       := LogicSubStr(DolgOldEm, 22, true, true);
      set s.tmpTabWorkTime.sDol2       := LogicSubStr(DolgOldEm, 22, true, true);
      set s.tmpTabWorkTime.sDol3       := LogicSubStr(DolgOldEm, 22, true, true);
    }
    else
    {
      set s.tmpTabWorkTime.sDol1       := LogicSubStr(Dolg, 22, true, true);
      set s.tmpTabWorkTime.sDol2       := LogicSubStr(Dolg, 22, true, true);
      set s.tmpTabWorkTime.sDol3       := LogicSubStr(Dolg, 22, true, true);
    }
    set s.tmpTabWorkTime.sTabNum     := TabN;
    set s.tmpTabWorkTime.LSTABNREC   := comp(LSTABNREC);
    set s.tmpTabWorkTime.cPEREXOD    := comp(PEREXODNREC);
    set s.tmpTabWorkTime.day1[1]     := D1;
    set s.tmpTabWorkTime.day1[2]     := D2;
    set s.tmpTabWorkTime.day1[3]     := D3;
    set s.tmpTabWorkTime.day1[4]     := D4;
    set s.tmpTabWorkTime.day1[5]     := D5;
    set s.tmpTabWorkTime.day1[6]     := D6;
    set s.tmpTabWorkTime.day1[7]     := D7;
    set s.tmpTabWorkTime.day1[8]     := D8;
    set s.tmpTabWorkTime.day1[9]     := D9;
    set s.tmpTabWorkTime.day1[10]    := D10;
    set s.tmpTabWorkTime.day1[11]    := D11;
    set s.tmpTabWorkTime.day1[12]    := D12;
    set s.tmpTabWorkTime.day1[13]    := D13;
    set s.tmpTabWorkTime.day1[14]    := D14;
    set s.tmpTabWorkTime.day1[15]    := D15;
    set s.tmpTabWorkTime.day1[16]    := 'X';
    set s.tmpTabWorkTime.day1[17]    := Time2StrDef(Double(Ch1));
    set s.tmpTabWorkTime.day1[18]    := Time2StrDef(Double(Ch2));
    set s.tmpTabWorkTime.day1[19]    := Time2StrDef(Double(Ch3));
    set s.tmpTabWorkTime.day1[20]    := Time2StrDef(Double(Ch4));
    set s.tmpTabWorkTime.day1[21]    := Time2StrDef(Double(Ch5));
    set s.tmpTabWorkTime.day1[22]    := Time2StrDef(Double(Ch6));
    set s.tmpTabWorkTime.day1[23]    := Time2StrDef(Double(Ch7));
    set s.tmpTabWorkTime.day1[24]    := Time2StrDef(Double(Ch8));
    set s.tmpTabWorkTime.day1[25]    := Time2StrDef(Double(Ch9));
    set s.tmpTabWorkTime.day1[26]    := Time2StrDef(Double(Ch10));
    set s.tmpTabWorkTime.day1[27]    := Time2StrDef(Double(Ch11));
    set s.tmpTabWorkTime.day1[28]    := Time2StrDef(Double(Ch12));
    set s.tmpTabWorkTime.day1[29]    := Time2StrDef(Double(Ch13));
    set s.tmpTabWorkTime.day1[30]    := Time2StrDef(Double(Ch14));
    set s.tmpTabWorkTime.day1[31]    := Time2StrDef(Double(Ch15));
    set s.tmpTabWorkTime.day1[32]    := 'X';
    set s.tmpTabWorkTime.day2[1]     := D16;
    set s.tmpTabWorkTime.day2[2]     := D17;
    set s.tmpTabWorkTime.day2[3]     := D18;
    set s.tmpTabWorkTime.day2[4]     := D19;
    set s.tmpTabWorkTime.day2[5]     := D20;
    set s.tmpTabWorkTime.day2[6]     := D21;
    set s.tmpTabWorkTime.day2[7]     := D22;
    set s.tmpTabWorkTime.day2[8]     := D23;
    set s.tmpTabWorkTime.day2[9]     := D24;
    set s.tmpTabWorkTime.day2[10]    := D25;
    set s.tmpTabWorkTime.day2[11]    := D26;
    set s.tmpTabWorkTime.day2[12]    := D27;
    set s.tmpTabWorkTime.day2[13]    := D28;
    set s.tmpTabWorkTime.day2[14]    := D29;
    set s.tmpTabWorkTime.day2[15]    := D30;
    set s.tmpTabWorkTime.day2[16]    := D31;
    set s.tmpTabWorkTime.day2[17]    := Time2StrDef(Double(Ch16));
    set s.tmpTabWorkTime.day2[18]    := Time2StrDef(Double(Ch17));
    set s.tmpTabWorkTime.day2[19]    := Time2StrDef(Double(Ch18));
    set s.tmpTabWorkTime.day2[20]    := Time2StrDef(Double(Ch19));
    set s.tmpTabWorkTime.day2[21]    := Time2StrDef(Double(Ch20));
    set s.tmpTabWorkTime.day2[22]    := Time2StrDef(Double(Ch21));
    set s.tmpTabWorkTime.day2[23]    := Time2StrDef(Double(Ch22));
    set s.tmpTabWorkTime.day2[24]    := Time2StrDef(Double(Ch23));
    set s.tmpTabWorkTime.day2[25]    := Time2StrDef(Double(Ch24));
    set s.tmpTabWorkTime.day2[26]    := Time2StrDef(Double(Ch25));
    set s.tmpTabWorkTime.day2[27]    := Time2StrDef(Double(Ch26));
    set s.tmpTabWorkTime.day2[28]    := Time2StrDef(Double(Ch27));
    set s.tmpTabWorkTime.day2[29]    := Time2StrDef(Double(Ch28));
    set s.tmpTabWorkTime.day2[30]    := Time2StrDef(Double(Ch29));
    set s.tmpTabWorkTime.day2[31]    := Time2StrDef(Double(Ch30));
    set s.tmpTabWorkTime.day2[32]    := Time2StrDef(Double(Ch31));

    set s.tmpTabWorkTime.otrbHM[1]     := перв_пол_дни;
    set s.tmpTabWorkTime.otrbHM[3]     := Time2StrDef(Double(перв_пол_часы));
    set s.tmpTabWorkTime.otrbHM[5]     := втор_пол_дни;
    set s.tmpTabWorkTime.otrbHM[7]     := Time2StrDef(Double(втор_пол_часы));
    set s.tmpTabWorkTime.otrbM[2]      := месяц_дни;
    set s.tmpTabWorkTime.otrbM[5]      := Time2StrDef(Double(месяц_часы));

    set s.tmpTabWorkTime.kodVidOpl[1]  := VidOpl1;
    set s.tmpTabWorkTime.kodVidOpl[3]  := VidOpl3;
    set s.tmpTabWorkTime.kodVidOpl[5]  := VidOpl5;
    set s.tmpTabWorkTime.kodVidOpl[9]  := VidOpl2;
    set s.tmpTabWorkTime.kodVidOpl[11] := VidOpl4;
    set s.tmpTabWorkTime.kodVidOpl[13] := VidOpl6;
    set s.tmpTabWorkTime.korSchet[1]   := Коррсчет1;
    set s.tmpTabWorkTime.korSchet[3]   := Коррсчет3;
    set s.tmpTabWorkTime.korSchet[5]   := Коррсчет5;
    set s.tmpTabWorkTime.korSchet[9]   := Коррсчет2;
    set s.tmpTabWorkTime.korSchet[11]  := Коррсчет4;
    set s.tmpTabWorkTime.korSchet[13]  := Коррсчет6;
    set s.tmpTabWorkTime.daysHours[1]  := Day1;
    set s.tmpTabWorkTime.daysHours[2]  := Time2StrDef(Double(Chas1));
    set s.tmpTabWorkTime.daysHours[5]  := Day3;
    set s.tmpTabWorkTime.daysHours[6]  := Time2StrDef(Double(Chas3));
    set s.tmpTabWorkTime.daysHours[9]  := Day5;
    set s.tmpTabWorkTime.daysHours[10] := Time2StrDef(Double(Chas5));
    set s.tmpTabWorkTime.daysHours[17] := Day2;
    set s.tmpTabWorkTime.daysHours[18] := Time2StrDef(Double(Chas2));
    set s.tmpTabWorkTime.daysHours[21] := Day4;
    set s.tmpTabWorkTime.daysHours[22] := Time2StrDef(Double(Chas4));
    set s.tmpTabWorkTime.daysHours[25] := Day6;
    set s.tmpTabWorkTime.daysHours[26] := Time2StrDef(Double(Chas6));
    Fill_tmpTabWorkTime;
    var i, ld : word;
    ld := Last_Day(dBeg);

    _LOOP ContDoc2     
      for(i := 1; i <= ld; inc(i))
        if (s.tmpTabWorkTime.wOJ[i] = 0) 
          if (ContDoc2.Dat1 <= date(i, Month(dBeg), Year(dBeg)) AND 
              ContDoc2.Dat2 >= date(i, Month(dBeg), Year(dBeg))
              )
            set s.tmpTabWorkTime.wOJ[i] := 1;

    if (s.Insert current tmpTabWorkTime = tsOk) {}
  }
  else
  {
    gp_add_day(D1,  s.tmpTabWorkTime.day1[1] );
    gp_add_day(D2,  s.tmpTabWorkTime.day1[2] );
    gp_add_day(D3,  s.tmpTabWorkTime.day1[3] );
    gp_add_day(D4,  s.tmpTabWorkTime.day1[4] );
    gp_add_day(D5,  s.tmpTabWorkTime.day1[5] );
    gp_add_day(D6,  s.tmpTabWorkTime.day1[6] );
    gp_add_day(D7,  s.tmpTabWorkTime.day1[7] );
    gp_add_day(D8,  s.tmpTabWorkTime.day1[8] );
    gp_add_day(D9,  s.tmpTabWorkTime.day1[9] );
    gp_add_day(D10, s.tmpTabWorkTime.day1[10]);
    gp_add_day(D11, s.tmpTabWorkTime.day1[11]);
    gp_add_day(D12, s.tmpTabWorkTime.day1[12]);
    gp_add_day(D13, s.tmpTabWorkTime.day1[13]);
    gp_add_day(D14, s.tmpTabWorkTime.day1[14]);
    gp_add_day(D15, s.tmpTabWorkTime.day1[15]);

    gp_add_hour(Ch1,  s.tmpTabWorkTime.day1[17]);
    gp_add_hour(Ch2,  s.tmpTabWorkTime.day1[18]);
    gp_add_hour(Ch3,  s.tmpTabWorkTime.day1[19]);
    gp_add_hour(Ch4,  s.tmpTabWorkTime.day1[20]);
    gp_add_hour(Ch5,  s.tmpTabWorkTime.day1[21]);
    gp_add_hour(Ch6,  s.tmpTabWorkTime.day1[22]);
    gp_add_hour(Ch7,  s.tmpTabWorkTime.day1[23]);
    gp_add_hour(Ch8,  s.tmpTabWorkTime.day1[24]);
    gp_add_hour(Ch9,  s.tmpTabWorkTime.day1[25]);
    gp_add_hour(Ch10, s.tmpTabWorkTime.day1[26]);
    gp_add_hour(Ch11, s.tmpTabWorkTime.day1[27]);
    gp_add_hour(Ch12, s.tmpTabWorkTime.day1[28]);
    gp_add_hour(Ch13, s.tmpTabWorkTime.day1[29]);
    gp_add_hour(Ch14, s.tmpTabWorkTime.day1[30]);
    gp_add_hour(Ch15, s.tmpTabWorkTime.day1[31]);

    gp_add_day(D16, s.tmpTabWorkTime.day2[1]);
    gp_add_day(D17, s.tmpTabWorkTime.day2[2]);
    gp_add_day(D18, s.tmpTabWorkTime.day2[3]);
    gp_add_day(D19, s.tmpTabWorkTime.day2[4]);
    gp_add_day(D20, s.tmpTabWorkTime.day2[5]);
    gp_add_day(D21, s.tmpTabWorkTime.day2[6]);
    gp_add_day(D22, s.tmpTabWorkTime.day2[7]);
    gp_add_day(D23, s.tmpTabWorkTime.day2[8]);
    gp_add_day(D24, s.tmpTabWorkTime.day2[9]);
    gp_add_day(D25, s.tmpTabWorkTime.day2[10]);
    gp_add_day(D26, s.tmpTabWorkTime.day2[11]);
    gp_add_day(D27, s.tmpTabWorkTime.day2[12]);
    gp_add_day(D28, s.tmpTabWorkTime.day2[13]);
    gp_add_day(D29, s.tmpTabWorkTime.day2[14]);
    gp_add_day(D30, s.tmpTabWorkTime.day2[15]);
    gp_add_day(D31, s.tmpTabWorkTime.day2[16]);

    gp_add_hour(Ch16, s.tmpTabWorkTime.day2[17]);
    gp_add_hour(Ch17, s.tmpTabWorkTime.day2[18]);
    gp_add_hour(Ch18, s.tmpTabWorkTime.day2[19]);
    gp_add_hour(Ch19, s.tmpTabWorkTime.day2[20]);
    gp_add_hour(Ch20, s.tmpTabWorkTime.day2[21]);
    gp_add_hour(Ch21, s.tmpTabWorkTime.day2[22]);
    gp_add_hour(Ch22, s.tmpTabWorkTime.day2[23]);
    gp_add_hour(Ch23, s.tmpTabWorkTime.day2[24]);
    gp_add_hour(Ch24, s.tmpTabWorkTime.day2[25]);
    gp_add_hour(Ch25, s.tmpTabWorkTime.day2[26]);
    gp_add_hour(Ch26, s.tmpTabWorkTime.day2[27]);
    gp_add_hour(Ch27, s.tmpTabWorkTime.day2[28]);
    gp_add_hour(Ch28, s.tmpTabWorkTime.day2[29]);
    gp_add_hour(Ch29, s.tmpTabWorkTime.day2[30]);
    gp_add_hour(Ch30, s.tmpTabWorkTime.day2[31]);
    gp_add_hour(Ch31, s.tmpTabWorkTime.day2[32]);

    gp_add_dbl(перв_пол_дни  , s.tmpTabWorkTime.otrbHM[1]);
    gp_add_hour(перв_пол_часы, s.tmpTabWorkTime.otrbHM[3]);
    gp_add_dbl(втор_пол_дни  , s.tmpTabWorkTime.otrbHM[5]);
    gp_add_hour(втор_пол_часы, s.tmpTabWorkTime.otrbHM[7]);
    gp_add_dbl(месяц_дни     , s.tmpTabWorkTime.otrbM[2]);
    gp_add_hour(месяц_часы   , s.tmpTabWorkTime.otrbM[5]);

    gp_add_dbl(Day1  , daysHours[1] );
    gp_add_hour(Chas1, daysHours[2] );
    gp_add_dbl(Day3  , daysHours[5] );
    gp_add_hour(Chas3, daysHours[6] );
    gp_add_dbl(Day5  , daysHours[9] );
    gp_add_hour(Chas5, daysHours[10]);
    gp_add_dbl(Day2  , daysHours[17]);
    gp_add_hour(Chas2, daysHours[18]);
    gp_add_dbl(Day4  , daysHours[21]);
    gp_add_hour(Chas4, daysHours[22]);
    gp_add_dbl(Day6  , daysHours[25]);
    gp_add_hour(Chas6, daysHours[26]);
    Fill_tmpTabWorkTime;
    if (s.update current tmpTabWorkTime = tsOk) {}
  }

/*
   set s.tmpTabWorkTime.kodMiss[1]    := KodNejavki1;
   set s.tmpTabWorkTime.kodMiss[3]    := KodNejavki3;
   set s.tmpTabWorkTime.kodMiss[5]    := KodNejavki5;
   set s.tmpTabWorkTime.kodMiss[7]    := KodNejavki7;
   set s.tmpTabWorkTime.dayMiss[1]    := DayNejavok1 + if (trim(ChasNejavok1) = '', '', '/') + GetNormDate(ChasNejavok1);
   set s.tmpTabWorkTime.dayMiss[3]    := DayNejavok3 + if (trim(ChasNejavok3) = '', '', '/') + GetNormDate(ChasNejavok3);
   set s.tmpTabWorkTime.dayMiss[5]    := DayNejavok5 + if (trim(ChasNejavok5) = '', '', '/') + GetNormDate(ChasNejavok5);
   set s.tmpTabWorkTime.dayMiss[7]    := DayNejavok7 + if (trim(ChasNejavok7) = '', '', '/') + GetNormDate(ChasNejavok7);
   set s.tmpTabWorkTime.kodMiss[9]    := KodNejavki2;
   set s.tmpTabWorkTime.kodMiss[11]   := KodNejavki4;
   set s.tmpTabWorkTime.kodMiss[13]   := KodNejavki6;
   set s.tmpTabWorkTime.kodMiss[15]   := KodNejavki8;
   set s.tmpTabWorkTime.dayMiss[9]    := DayNejavok2 + if (trim(ChasNejavok2) = '', '', '/') + GetNormDate(ChasNejavok2);
   set s.tmpTabWorkTime.dayMiss[11]   := DayNejavok4 + if (trim(ChasNejavok4) = '', '', '/') + GetNormDate(ChasNejavok4);
   set s.tmpTabWorkTime.dayMiss[13]   := DayNejavok6 + if (trim(ChasNejavok6) = '', '', '/') + GetNormDate(ChasNejavok6);
   set s.tmpTabWorkTime.dayMiss[15]   := DayNejavok8 + if (trim(ChasNejavok8) = '', '', '/') + GetNormDate(ChasNejavok8);    
*/

end.
.}
.}
.{CheckEnter FIRSTPAGE
.}
.}
.begin
   RunInterface(iEMTabWorkTime, sPodr, dBeg, dEnd,0);
   s.delete all tmpTabWorkTime;
end.
.endform
