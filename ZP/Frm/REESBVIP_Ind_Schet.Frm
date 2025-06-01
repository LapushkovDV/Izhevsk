/*************************************************************************************************\
* Наименование: Отчеты - Отчеты по оплате труда - Реестры - ПЕчать                                *
* Контур/Модуль: Заработная плата                                                                 *
* Примечание: Индивидуальный счет                                                                 *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#1724         |19/01/18|Кириллов Э.П.            |НПО Энергомаш                    *
\*************************************************************************************************/
.LinkForm 'REESBVIP_Ind_Schet' Prototype is 'REESBVIP'
.Group 'Реестры по перечислению в сбербанк'
.NameInList 'ЭМ.Индивидуальный счет'
.f "NUL"
.var
  path:String
  pGP_FRFormatterOfReestr: GP_FRFormatterOfReestr;
  NPOEM_OWN_ZP_Rep_PDF :boolean;
.endvar
.fields
  path
.endfields
.{
.{
.{
.}
.}
.if BRPAGE
.else
.end
.begin
  path:= '';
  NPOEM_OWN_ZP_Rep_PDF:=True;
  if ExistTune('NPOEM_OWN.ZP.Rep.PDF')
  {  NPOEM_OWN_ZP_Rep_PDF:=boGetTune('NPOEM_OWN.ZP.Rep.PDF');
  }
  if NPOEM_OWN_ZP_Rep_PDF
  {  pGP_FRFormatterOfReestr.RunPDF(SbbanNrec,'GP Индивидуальный Счет',false);
  }
  else
  { pGP_FRFormatterOfReestr.RunPDF(SbbanNrec,'');
  }
end.
  ^
.}
Формирование файла электронного реестра завершено.
.endform
