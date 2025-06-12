
table struct tempOborCFG
(
  Nrec    : comp,
  nRep    : word,
  REPNAME : string,
  Descr   : string,
  dRep    : date
)
;

table struct tempView
(
  sString : array [1..40] of string,
  dDouble : array [1..40] of double,
  cComp   : array [1..40] of comp,
  wWord   : array [1..40] of LongInt,
  dDate   : array [1..40] of Date,
  tTime   : array [1..40] of time
)
;

table struct tableCaption
(
  fldName : array [1..120] of string,
  fldOrd  : array [1..120] of word,
  fldType : array [1..120] of word
)
;
const
   cgReport_30000 = 30000;
// По просьбе трудящихся, которым не хватает фильтров из Аналитик
// заводим свои константы для новых фильтров
  cgiDescr = 100000;
  cgiPickNotKAU = 100000;
end;
