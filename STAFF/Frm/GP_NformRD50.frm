/*************************************************************************************************\
* Наименование: Приказ (распоряжение) об изменении оклада                                         *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#346          |30/09/17|Кузьмин П.Ю.             |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_NformRD50' Prototype is 'NformRD50'
.declare
#include GP_RepLaborContract.vih
.enddeclare
.NameInList 'ЭМ.Приказ (распоряжение) об изменении оклада'
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
  if (not FindContDoc(50)) {
    message('Не найдена запись таблицы ContDoc с TypeOper=50'+
            ', Person='+string(PersNrec,0,0)+' и cStr='+string(AppointNrec,0,0),error);
    exit;
  }

  var Rep:GP_RepLaborContract;
  Rep.Run(v.ContDoc.TypeOper,1,v.ContDoc.NRec,PersNrec,AppointNrec);
end.
.endform
