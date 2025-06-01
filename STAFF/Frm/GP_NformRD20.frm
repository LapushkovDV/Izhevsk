/*************************************************************************************************\
* Наименование: Приказ о единовременном премировании                                              *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#356          |18/07/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NformRD20' Prototype is 'NformRD20'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ о единовременном премировании'
.f 'NUL'
.create view v
from
  ContDoc,PartDoc;
.if NAttrP
.end
.if NAttrV
.end
.if NAttr
.end
.function FindContDoc(iTypeOper:Integer):boolean;
begin
  result:=v.getfirst fastfirstrow ContDoc
          where((
            iTypeOper==ContDoc.TypeOper and PersNrec==ContDoc.Person and
            0 << ContDoc.SeqNmb and AppointNrec==ContDoc.cStr(noindex) and
            v.ContDoc.cPart/==PartDoc.NRec and TitleDocNrec==PartDoc.cDoc(noindex)
          ))=tsok;
end.
.begin
  if (not FindContDoc(20))
    if (not FindContDoc(21)) {
      message('Не найдена запись таблицы ContDoc с TypeOper=(20 или 21)'+
              ', Person='+string(PersNrec,0,0)+' и cStr='+string(AppointNrec,0,0),error);
      exit;
    }

  var Rep:GP_RepLaborContract;
  Rep.Run(v.ContDoc.TypeOper,1,v.ContDoc.NRec,PersNrec,AppointNrec);
end.
.endform
