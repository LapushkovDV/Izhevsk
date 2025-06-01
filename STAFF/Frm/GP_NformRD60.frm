/*************************************************************************************************\
* Наименование: Приказ (распоряжение) об изменении режима работ                                   *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#172          |28/08/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NformRD60' Prototype is 'NformRD60'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ (распоряжение) об изменении режима работ'
.f 'NUL'
.create view v
from
  ContDoc,PartDoc;
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
  if (not FindContDoc(60)) {
    message('Не найдена запись таблицы ContDoc с TypeOper=60'+
            ', Person='+string(PersNrec,0,0)+' и cStr='+string(AppointNrec,0,0),error);
    exit;
  }

  var Rep:GP_RepLaborContract;
  Rep.Run(v.ContDoc.TypeOper,1,v.ContDoc.NRec,PersNrec,AppointNrec);
end.
.endform
