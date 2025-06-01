/*************************************************************************************************\
* Наименование:  Разработка отчетной формы                                                        *
*                РПД № 40 "Отзыв из отпуска"                                                      *
* Контур/Модуль: Управление персоналом                                                            *
* Примечание:    Управление персоналом->Документы->Все приказы по персоналу->                     *
*                   Окно редактирования приказа =Приказ=->                                        *
*                   Контекстное меню <Печать индивидуальной формы приказа>                        *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#167          |20/06/17|Валента А.В.             |НПО Энергомаш                    *
\*************************************************************************************************/
.linkform 'GP_WithdrawVacation' prototype is 'VozOtpForm'
.nameinlist'ЭМ."Приказ об отзыве сотрудника из отпуска" ЭМ'
.F 'NUL'
.create view vWV
as select
 *
from
  tmpWithdrawVac
;
.{ PushReasonCycle CheckEnter
.}
.begin

  vWV.delete all tmpWithdrawVac;
  vWV.ClearBuffer(vWV.tntmpWithdrawVac);
  vWV.tmpWithdrawVac.NameVacation := Наименование_отпуска;
  vWV.tmpWithdrawVac.FIO          := FIO;
  vWV.tmpWithdrawVac.Dol          := должность;
  vWV.tmpWithdrawVac.Podr         := подразделение;
  vWV.tmpWithdrawVac.DateWithdraw := StrToDate(Дата_Отзыва, '"DD" Mon YYYY');
  vWV.tmpWithdrawVac.dBeg         := StrToDate(Дата_Начала_Отзыва, '"DD" Mon YYYY');
  vWV.tmpWithdrawVac.dEnd         := StrToDate(Дата_Окончания_Отзыва, '"DD" Mon YYYY');
  vWV.tmpWithdrawVac.NumDoc       := Номер_Приказа;
  vWV.tmpWithdrawVac.DateDoc      := StrToDate(Дата_Приказа, 'DD/MM/YYYY');
  vWV.tmpWithdrawVac.cTitleDoc    := TitleDocNrec;
  vWV.tmpWithdrawVac.cPers        := PersNrec;
  vWV.tmpWithdrawVac.cApp         := AppointNrec;
  if (vWV.insert current tmpWithdrawVac = tsOk)
    {} 

  RunInterface(iEMWithdrawVacationReport, Номер_Приказа, Replace(Дата_Приказа, '/', '.'), TitleDocNrec, PersNrec);
end.
.endform
