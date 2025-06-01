/*************************************************************************************************\
* Наименование: Импорт Каталога Матценностей из Excel в Галактику                                 *
* Контур/Модуль: Обмен бизнес-документами                                                         *
* Примечание:                                                                                     *
*        1) Исправлена ошибка в наименовании статуса ExImLogR.Status                              *
*        2) Выделены в отдельные поля дополнительные данные в протоколе при импорте МЦ из Excel   *
*           обозначение позиции, номер строки, адрес R1C1 ячейки                                  *
* Вид работы  |Номер         |Дата      |Исполнитель              |Проект                         *
* ---------------------------------------------------------------------------------------------   *
* Разработка  |#2515         |18/06/2018|Шунин А.Г.               |НПО Энергомаш                  *
\*************************************************************************************************/

// Печать протокола операции экспорта-импорта операций

.AutoForm 'ExpImpJourn'(hMarker : longint; OnlyBad:word)
.Create view JournLog
var
 lMarker : longint;
 bMarker : longint;
 nIndex  : longint;
 cRec    : Comp;
 sKOD, sEXIMLOGE_NOREFFLD, sR1C1, sExImLogR_Status: string;
 wPos :word;
from
  IEHead,
  ExImLogT,
  ExImLogR,
  ExImLogE,
  UsersDoc
where
((
  ExImLogT.cLog     == IEHead.NRec        and
  IEHead.SubTipDoc  == UsersDoc.TipUsers  and
  ExImLogT.NRec     == ExImLogR.cLogT     and
  ExImLogR.NRec     == ExImLogE.cLogR     and
  OnlyBad           <<= ExImLogR.Status(noindex)
))
;
.fields
  if(OnlyBad = 0, 'всей операции', 'ошибок операции')
  if(IEHead.ImpExp = 2, 'экспорта', 'импорта')
  // тип документа
  if(IEHead.TipDoc = 1, 'Документы-Основания',
    if(IEHead.TipDoc = 2, 'Накладные',
    if(IEHead.TipDoc = 3, 'Счета-Фактуры',
    if(IEHead.TipDoc = 4, 'Платёжные документы',
    if(IEHead.TipDoc = 5, 'Прайс листы',
    if(IEHead.TipDoc = 6, 'Партии',
    if(IEHead.TipDoc = 7, 'Валютные документы',
    if(IEHead.TipDoc = 9, 'Каталог контрагентов',
    if(IEHead.TipDoc = 10, 'Каталог матценностей',
    if(IEHead.TipDoc = 11, 'Каталог вагонов',
    if(IEHead.TipDoc = 12, 'Сертификаты качества',
    if(IEHead.TipDoc = 13, 'Значения показателей качества',
    if(IEHead.TipDoc = 14, 'Лимитно заправочные ведомости',
    if(IEHead.TipDoc = 15, 'Проводки',
    if(IEHead.TipDoc = 16, 'Платежные ведомости',
    if(IEHead.TipDoc = 17, 'Ведомости депонирования',
    if(IEHead.TipDoc = 19, 'Операции ОС',
    if(IEHead.TipDoc = 20, 'Нематериальные активы',
    if(IEHead.TipDoc = 21, 'Путевые листы',
    if(IEHead.TipDoc = 22, 'Типы акцизных марок',
    if(IEHead.TipDoc = 23, 'Акцизные марки',
    if(IEHead.TipDoc = 24, 'Заявки на производство акцизных марок',
    if(IEHead.TipDoc = 32, 'Заявки на расходование средств',
    if(IEHead.TipDoc = 33, 'Документы розничной торговли',
    if(IEHead.TipDoc = 34, 'Инвентаризация',
     'Неизвестный тип данных')))))))))))))))))))))))))
  // подтип документа
  if(IEHead.TipDoc = 1, // ДО
      if(IEHead.SubTipDoc = 101, 'Закупка',
      if(IEHead.SubTipDoc = 102, 'Прием товара на консигнацию',
      if(IEHead.SubTipDoc = 201, 'Продажа',
      if(IEHead.SubTipDoc = 202, 'Отпуск товара на условиях консигнации',
      if(IEHead.SubTipDoc = 501, 'ЛЗК на отпуск в производство',
      if(IEHead.SubTipDoc = 510, 'Заявка на обслуживание',
      if(IEHead.SubTipDoc = 520, 'Давальческая обработка',
      if(IEHead.SubTipDoc = 211, 'ДО на предоплату продаж',
      if(IEHead.SubTipDoc = 111, 'ДО на предоплату закупок',
      'Неизвестный тип данных'))))))))),

    if(IEHead.TipDoc = 2, // Накладные
      if(IEHead.SubTipDoc = 101, 'Накладная на прием товара',
      if(IEHead.SubTipDoc = 201, 'Накладная на отпуск товара',
      if(IEHead.SubTipDoc = 205, 'Товарно-транспортная накладная',
      if(IEHead.SubTipDoc = 111, 'Акт на прием услуги',
      if(IEHead.SubTipDoc = 211, 'Акт на оказание услуги',
      if(IEHead.SubTipDoc = 501, 'Накладная на отпуск в производство',
      if(IEHead.SubTipDoc = 502, 'Накладная на возврат из производства',
      if(IEHead.SubTipDoc = 504, 'Акт на списание из производства',
      if(IEHead.SubTipDoc = 600, 'Накладная на внутреннее перемещение',
      if(IEHead.SubTipDoc = 601, 'Накладная на передачу товаров в ОС',
      if(IEHead.SubTipDoc = 602, 'Накладная на передачу товаров в МБП',
      if(IEHead.SubTipDoc = 603, 'Накладная на передачу товаров в розницу',
      if(IEHead.SubTipDoc = 605, 'Накладная на передачу товаров в НМА',
      if(IEHead.SubTipDoc = 611, 'Акт об излишках',
      if(IEHead.SubTipDoc = 612, 'Акт о недостаче',
      if(IEHead.SubTipDoc = 513, 'Акт на оказание услуг при ремонте',
      if(IEHead.SubTipDoc = 206, 'Рекламационная накладная на возврат продавцу',
      if(IEHead.SubTipDoc = 106, 'Рекламационная накладная от покупателя на возврат',
      if(IEHead.SubTipDoc = 112, 'Акт на регистрацию начисленных налогов',
      if(IEHead.SubTipDoc = 204, 'Акт на списание МЦ',
      if(IEHead.SubTipDoc = 210, 'Акт передачи оборудования в монтаж',
      if(IEHead.SubTipDoc = 229, 'Акт передачи материалов на строительство',
      if(IEHead.SubTipDoc = 629, 'Отчет по форме М-29',
      if(IEHead.SubTipDoc = 521, 'Накладная на отпуск сырья в переработку',
      if(IEHead.SubTipDoc = 522, 'Накладная на возврат готовой продукции из переработки',
      if(IEHead.SubTipDoc = 523, 'Накладная на возврат непереработанного сырья',
      'Неизвестный тип данных')))))))))))))))))))))))))),

    if(IEHead.TipDoc = 3, // Счета-Фактуры
       UsersDoc.Name,

    if(IEHead.TipDoc = 4, // Платежные документы
      if(IEHead.SubTipDoc = 1 , 'Платежное поручение',
      if(IEHead.SubTipDoc = 2 , 'Входящие документы',
      if(IEHead.SubTipDoc = 3 , 'Инкассовое поручение',
      if(IEHead.SubTipDoc = 4 , 'Заявление на аккредитив',
      if(IEHead.SubTipDoc = 5 , 'Заявление об отказе от акцепта',
      if(IEHead.SubTipDoc = 6 , 'Реестр чеков',
      if(IEHead.SubTipDoc = 7 , 'Приходный кассовый ордер',
      if(IEHead.SubTipDoc = 8 , 'Расходный кассовый ордер',
      if(IEHead.SubTipDoc = 9 , 'Авансовый отчет',
      if(IEHead.SubTipDoc = 10, 'Бухгалтерская справка',
      if(IEHead.SubTipDoc = 11, 'Платежное требование-поручение',
      if(IEHead.SubTipDoc = 31, 'Валютное платежное поручение',
      if(IEHead.SubTipDoc = 32, 'Платежное требование',
      if(IEHead.SubTipDoc = 33, 'Входящее валютное',
      if(IEHead.SubTipDoc = 21, 'Исходящее авизо',
      if(IEHead.SubTipDoc = 22, 'Входящее авизо',
      if(IEHead.SubTipDoc = 17, 'Валютный ПКО',
      if(IEHead.SubTipDoc = 18, 'Валютный РКО',
      if(IEHead.SubTipDoc = 37, 'Налоговое платежное поручение',
      if(IEHead.SubTipDoc = 38, 'Исходящее налоговое авизо',
      if(IEHead.SubTipDoc = 39, 'Входящее налоговое авизо',
      'Неизвестный тип данных'))))))))))))))))))))),

    if(IEHead.TipDoc = 5, '',
    if(IEHead.TipDoc = 6, '',
    if(IEHead.TipDoc = 7,
      if(IEHead.SubTipDoc = 34, 'Поручение на покупку иностранной валюты'    ,
      if(IEHead.SubTipDoc = 35, 'Поручение на продажу иностранной валюты'    ,
      if(IEHead.SubTipDoc = 36, 'Поручение на конвертацию иностранной валюты',
      'Неизвестный тип данных'))),
    if(IEHead.TipDoc = 9, '',
    if(IEHead.TipDoc = 10, '',
    if(IEHead.TipDoc = 11, '',
    if(IEHead.TipDoc = 12, '',
    if(IEHead.TipDoc = 13, '',
    if(IEHead.TipDoc = 14, '',
    if(IEHead.TipDoc = 15, '',
    if(IEHead.TipDoc = 16, '',
    if(IEHead.TipDoc = 17, '',
    if(IEHead.TipDoc = 19, // Операции ОС
      if( IEHead.SubTipDoc = 1, 'Операция ОС поступление',
      if( IEHead.SubTipDoc = 2, 'Операция ОС перемещение',
      if( IEHead.SubTipDoc = 3, 'Операция ОС изменение стоимости',
      if( IEHead.SubTipDoc = 4, 'Операция ОС выбытие',
      if( IEHead.SubTipDoc = 5, 'Операция ОС амортизация',
      if( IEHead.SubTipDoc = 6, 'Операция ОС переоценка',
      if( IEHead.SubTipDoc = 7, 'Операция ОС изменение вида(группы)',
      'Неизвестный тип данных'))))))),
    if(IEHead.TipDoc = 20, // Нематериальные активы
      if( IEHead.SubTipDoc = 1, 'НМА поступление',
      if( IEHead.SubTipDoc = 2, 'НМА перемещение',
      if( IEHead.SubTipDoc = 3, 'НМА изменение стоимости',
      if( IEHead.SubTipDoc = 4, 'НМА выбытие',
      if( IEHead.SubTipDoc = 5, 'НМА амортизация',
      if( IEHead.SubTipDoc = 7, 'НМА изменение вида/нормы амортизации',
      'Неизвестный тип данных')))))),
    if(IEHead.TipDoc = 21,'',
    if(IEHead.TipDoc = 22,'',
    if(IEHead.TipDoc = 23,
      if(IEHead.SubTipDoc = 1, 'Приход',
      if(IEHead.SubTipDoc = 2, 'Расход',
      'Неизвестный тип данных')),
    if(IEHead.TipDoc = 24,'',
    if(IEHead.TipDoc = 32,'',
    if(IEHead.TipDoc = 34,'',
    if(IEHead.TipDoc = 33,
      If ( IEHead.SubTipDoc = cgDoc_0917, 'Акт на списание',
      If ( IEHead.SubTipDoc = cgDoc_0908, 'Накладная на внутреннее перемещение',
      If ( IEHead.SubTipDoc = cgDoc_0919, 'Накладная на реализацию','Неизвестный тип данных'))),
    'Неизвестный тип данных')))))))))))))))))))))))))

  IEHead.Name
  ExImLogT.fDate
  ExImLogT.fTime
  if(ExImLogT.Status = 1, 'С ошибками',
    if(ExImLogT.Status = 2, 'Успешно',
    if(ExImLogT.Status = 3, 'Прервана пользователем', '')))
  ExImLogR.Value

  sExImLogR_Status


// список ошибок -- .table 'JournLog.ExImLogE'
  sKOD
  sEXIMLOGE_NOREFFLD
  sR1C1
  ExImLogE.Value
  if(ExImLogE.Status = 1, 'Ссылка не определена', if (ExImLogE.Status = 2 , 'Детализация : ', ''))
  ExImLogE.rStr

// дублирование в списке ошибок
  sExImLogR_Status
  ExImLogR.Value
.endfields

.begin
  if( GetMarkerCount(hMarker) < 1 ) ExpImpJourn.fBreak;
  else nIndex := 0;
end.

.{table 'JournLog.ExImLogT'
.begin
  if( GetMarker(hMarker, nIndex, cRec) )
  {
    if( JournLog.GetFirst ExImLogT where((cRec == ExImLogT.NRec)) <> tsOK ){}
    nIndex := nIndex+1;
  }
  else ExpImpJourn.fBreak;



end.
  Протокол ^ ^ данных типа "^ ^" по настройке ^
  Операция выполнена ^ ^ ^
  ──────────────────────────────────────────────────────────────────────────────────────────────
.{table 'JournLog.ExImLogR'
.begin
 sExImLogR_Status :=  if ( ExImLogR.Status = ieNotDefined,     'Не определен',
                      if ( ExImLogR.Status = ieImportedInTemp, 'Прочитан',
                      if ( ExImLogR.Status = ieExported,       'Экспортирован',
                      if ( ExImLogR.Status = ieImported,       'Импортирован',
                      if ( ExImLogR.Status = ieSkip,           'Уже импортировался',
                      if ( ExImLogR.Status = ieAbsentFields,   'Не заполнены обязательные поля','')))))) ;

sKOD := '';  sR1C1 := '';

wPos := Instr('Обозн.:',EXIMLOGR.VALUE);
if wPos > 0
{
  sKOD := trim(SubStr(EXIMLOGR.VALUE, wPos+7, Instr(' ',EXIMLOGR.VALUE)-7) );
}

end.
  Значение:@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  Статус: ^
.{table 'JournLog.ExImLogE'
.begin
sEXIMLOGE_NOREFFLD := EXIMLOGE.NOREFFLD;
wPos := Instr(' R',sEXIMLOGE_NOREFFLD);
if wPos > 0
{
   sR1C1 := SubStr(sEXIMLOGE_NOREFFLD, wPos+1, 255);
   sEXIMLOGE_NOREFFLD:= SubStr(sEXIMLOGE_NOREFFLD, 1, wPos-1);
}

end.
Ошибка: ^ ^ ^ ^ ^ ^ ^ ^
.}
.}
  ──────────────────────────────────────────────────────────────────────────────────────────────
.}
.endform
