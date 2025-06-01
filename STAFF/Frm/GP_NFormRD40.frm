/*************************************************************************************************\
* Наименование: Приказ (распоряжение) об отзыве сотрудника из отпуска                             *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#167          |10/10/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NFormRD40' Prototype is 'VozOtpForm'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ (распоряжение) об отзыве сотрудника из отпуска'
.f 'NUL'
.var
  pContDoc_Fnd:comp;
.endvar

.create view v
var
 pPersons:comp;
 pAppoint:comp;
 pTitleDoc:comp;
 iTypeOper:word;
from
  ContDoc,PartDoc
where
((
    iTypeOper==ContDoc.TypeOper
 and pPersons==ContDoc.Person
 and 0 << ContDoc.SeqNmb
 and pAppoint==ContDoc.cStr(noindex)
 and ContDoc.cPart/==PartDoc.NRec
 and pTitleDoc==PartDoc.cDoc(noindex)
 and ContDoc.ObjNrec  ==  OtpOtz.Nrec
))
  ;
.function FindContDoc(iTypeOper:Integer):comp;
var _DT:date;
begin
 Result:=0;
 v.pPersons:=PersNrec;
 v.pAppoint:=AppointNrec;
 v.pTitleDoc:=TitleDocNrec;
 v.iTypeOper:=iTypeOper;
 _DT:=StrToDate(Replace(Дата_Начала_Отзыва,'.','/'),'DD/MM/YYYY');
 if _DT=ZeroDaTE  //StrToDate('"19" марта 2018','"DD" mon YYYY');
 { _DT:=StrToDate(Дата_Начала_Отзыва,'"DD" mon YYYY');

 }
 v._Loop ContDoc
 { if Result=0  Result:=v.ContDoc.Nrec;
   if v.getfirst OTPOTZ=tsok
    if v.OTPOTZ.DATAN = _DT
    { Result:=v.ContDoc.Nrec;
      break;
    }
 }
end.
.{ PushReasonCycle CheckEnter
.}
.begin
  pContDoc_Fnd:=FindContDoc(40);
  if (pContDoc_Fnd=0) {
    message('Не найдена запись таблицы ContDoc с TypeOper=40'+
            ', Person='+string(PersNrec,0,0)+' и cStr='+string(AppointNrec,0,0),error);
    exit;
  }
  var Rep:GP_RepLaborContract;
  Rep.Run(v.ContDoc.TypeOper,1,pContDoc_Fnd,PersNrec,AppointNrec);
end.
.endform
