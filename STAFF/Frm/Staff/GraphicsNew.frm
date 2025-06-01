/*************************************************************************************************\
* Наименование: Учет рабочего времени - графики - печать - печать помеченных графиков             *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание: iEMReportGraphicsNew                                                                *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#384          |09/11/18|Кириллов Э.П.            |НПО Энергомаш                    *
\*************************************************************************************************/
//******************************************************************************
//                                                       ][a|coN (haicon@tut.by)
// Галактика 9.10 - ЭнергоМаш
// График работы сотрудников
//******************************************************************************
.linkform 'GP_TabelAllNew_EM' prototype is 'TabelAll'
.nameinlist'ЭМ."График работы сотрудников" ЭМ'
.F 'NUL'
.var
  marker: longint
.endvar
.create view s
var
  sNaim:string;
as select
 ExClassName.Name, tmpReportGraphics.*
from
  tmpReportGraphics
, KlRejim
, ExClassName
, ExClassVal
, ExClassSeg
where
((  
    coKlRejim               == ExClassName.wTable
AND 'Номер смены'           == ExClassName.Name
and sNaim                   ==  KlRejim.NRejim (NoIndex)
and ExClassName.ClassCode   == ExClassVal.ClassCode
AND coKlRejim               == ExClassVal.wTable
AND KlRejim.NRec            == ExClassVal.cRec
and ExClassVal.cClassSeg    == ExClassSeg.NRec
))
;
.begin
  s.delete all tmpReportGraphics;
end.
.{ TabelAll_AllGraphics checkenter
.{ TabelAll_OneGraphics checkenter
.begin
  var sn : string; sn := '';
  s.sNaim:=Naim;
  if (s.GetFirst ExClassName = tsOk)
    if (s.GetFirst KlRejim /*where ((Naim == KlRejim.NRejim (NoIndex)))*/ = tsOk)
      if (s.GetFirst ExClassVal /*where ((s.ExClassName.ClassCode == ExClassVal.ClassCode
                                    AND coKlRejim               == ExClassVal.wTable
                                    AND KlRejim.NRec            == ExClassVal.cRec))*/ = tsOk)
        if (s.GetFirst ExClassSeg /*where ((s.ExClassVal.cClassSeg == ExClassSeg.NRec))*/ = tsOk)
          sn := s.ExClassSeg.Value;

  s.insert into tmpReportGraphics set tmpReportGraphics.wYear := word(Year), tmpReportGraphics.Name := Naim, tmpReportGraphics.sSmena := sn;
end.
.{ TabelAll_day checkenter
.}  // конец цикла day		
.} // TabelAll_OneGraphics checkenter
.}// TabelAll_AllGraphics
.begin
   RunInterface(iEMReportGraphicsNew, word(Year));
   s.delete all tmpReportGraphics;
end.
.endform
