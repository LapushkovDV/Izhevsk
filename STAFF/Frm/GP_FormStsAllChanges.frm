/*************************************************************************************************\
* Наименование: Приложение к приказу. Изменение в штатном расписании ЭМ                           *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#1159          |06/10/17|Евдокимов М.А.          |НПО Энергомаш                    *
\*************************************************************************************************/

.LinkForm 'GP_FormStsAllChanges' Prototype is 'FormStsAll'
.declare
#include GP_RepShrChanges.vih
#include GP_RepShrChangesParams.vih
#include GP_UserFilterStoreDSK.vih
.enddeclare
.NameInList 'ЭМ.Приложение к приказу. Изменение в штатном расписании ЭМ'
.Group 'Country' subGroup 'Russia'
.f 'NUL'
.var
  _store : GP_IUserFilterStore; //хранилище настроек отчета
  _rpdFilterMarker: longint; //маркер с кодами отображаемых в отчете видов РПД (из настроек отчета)
  _strPartMarker: longint; //маркер с нреками StrPart, отображаемых в отчете
.endvar
.begin

  //инициализируем хранилище настроек отчета (дск)
  _store := GP_IUserFilterStore(new(GP_UserFilterStoreDSK,InitDSK('GP_RepShrChangesParams_')));

  //интерфейс настроек отчета
  var params: GP_RepShrChangesParams noauto;
  params := new (GP_RepShrChangesParams, GP_RepShrChangesParams(_store));
  //загружаем параметры отчета из хранилища
  params.Load;
  //показываем пользователю интерфейс параметров отчета
  if params.ShowUI != cmDefault
    //если пользователь не нажал продолжить - закрываем отчет
    GP_FormStsAllChanges.fExit;
  else
  {
    //если пользователь нажал продолжить - сохраняем параметры отчета в хранилище
    params.Save;
    //из хранилища загружаем маркер доступных видов рпд
    _rpdFilterMarker := _store.LoadMarker('rpdFilterMarker');
    //т.к. в прототипе нет нрека заголовка приказа - инициализируем маркер с нреками StrPart
    //его мы затем передадим в интерфейс, формирующий excel-отчет
    _strPartMarker := InitMarker('', 8, 10, 100, true);
  }

end.

.{StsAll_Cycle1  CheckEnter

.begin
  //записываем в маркер нрек отображаемого в отчете strpart
  if foundmarker(_rpdFilterMarker, TypeOper)
    InsertMarker(_strPartMarker, STRPARTNREC);
end.


.{StsAll_Cycle2  CheckEnter
.}


.}

.begin
  //сохраняем маркер с нреками strpart в хранилище настроек отчета
  _store.SaveMarker(_strPartMarker, 'strPartMarker');
  //устанавливаем вид отчета - fastreport
  _store.SaveComp(0, 'ReportType');
  //инициализируем интерфейс отчета
  var report: GP_IRepShrChanges(GP_RepShrChanges);
  //запускаем отчет с настройками из хранилища
  report.ShowReport(_store);
  //деинициализируем маркеры
  DoneMarker(_strPartMarker, '');
  DoneMarker(_rpdFilterMarker, '');
end.

.endform
