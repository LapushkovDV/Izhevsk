/*************************************************************************************************\
* Наименование:  СКУД Настройка группировок отчетов                                               *
* Контур/Модуль: Кадры                                                                            *
* Примечание:                                                                                     *
*                                                                                                 *
* Вид работы  |Номер         |Дата    |Исполнитель              |Проект                           *
* ----------------------------------------------------------------------------------------------- *
* Разработка  |#175          |24/10/17|Кириллов Э.П.            |НПО Энергомаш                    *
\*************************************************************************************************/
const
 //СКУД
 cg_s_Len_a_gr = 21;
 //=====================
 cg_s_code_Department     = 1;
 cg_s_name_Department     = 'Подразделение';
 cg_s_code_Org            = 2;
 cg_s_name_Org            = 'Организация';
 cg_s_code_Pers           = 3;
 cg_s_name_Pers           = 'Сотрудник';
 cg_s_code_Date           = 4;
 cg_s_name_Date           = 'Дата';
 cg_s_code_Year           = 5;
 cg_s_name_Year           = 'Год';
 cg_s_code_quarter        = 6;
 cg_s_name_quarter        = 'квартал';
 cg_s_code_month          = 7;
 cg_s_name_month          = 'Месяц';
 cg_s_code_day            = 8;
 cg_s_name_day            = 'День';
 cg_s_code_Post           = 9;
 cg_s_name_Post           = 'Должность';
 cg_s_code_Rejim          = 10;
 cg_s_name_Rejim          = 'Режим';
 cg_s_code_Type           = 11;
 cg_s_name_Type           = 'сотрудники/организации';
 cg_s_code_Grup           = 12;
 cg_s_name_Grup           = 'СКУД Группы сотрудников ';
 cg_s_code_SEX            = 13;
 cg_s_name_SEX            = 'Пол';
 cg_s_code_Kateg          = 14;
 cg_s_name_Kateg          = 'СКУД Категории сотрудников';
 cg_s_code_MisConDuct     = 15;
 cg_s_name_MisConDuct     = 'нарушения дисциплины';
 cg_s_code_MisConDuctEx   = 16;
 cg_s_name_MisConDuctEx   = 'наличие нарушений';
 cg_s_code_OverWork       = 17;
 cg_s_name_OverWork       = 'наличие сверхурочных';
 cg_s_code_OverWork2      = 18;
 cg_s_name_OverWork2      = 'наличие переработки';
 cg_s_code_Holiday        = 19;
 cg_s_name_Holiday        = 'наличие работы в выходной';
 cg_s_code_cex            = 20;
 cg_s_name_cex            = 'цех';
 cg_s_code_abs            = 21;
 cg_s_name_abs            = 'код усл.обозначения';
end;
//===============================
//==== Для выбора работников на DSLQ
table struct tmp_report
(  nrec    :comp
  ,fio     :string[50]
  ,cDoc    :comp
  ,cDoc2   :comp
)
with index
( tmp_report01 = nrec
);
//====================
table struct tmp_report_gr
( nrec     :comp
 ,pr       :Integer
 ,sortname :string // для сортировки
 ,NumRec   :Longint //Кол-во записей в группировке
 ,cRec     :comp
 ,LevCode  :Word
 ,LevTitle :String[200]
 ,LevName  :String[200]
 ,LevKod   :String[50]
 ,dmd : array[1..100] of double
 ,dms : array[1..100] of string
 ,wTable   :word //код таблицы
 ,cRecord  :comp //ссылка на запись
)
with index
( tmp_report_gr01  = nrec (AutoInc, Unique),
  tmp_report_gr02  = pr+sortname,
  tmp_report_gr03  = cRec+sortname,
  tmp_report_gr04  = pr+cRec+sortname,
  tmp_report_gr05  = cRec+dms[1]
);

table struct tmp_report_sp
(  nrec     :comp
 , pr       :Integer
 , sortname :string // для сортировки
 , cRec     :comp   // ссылка на заголовок
 , cRecord  :comp   // ссылка на спецификацию
 , cRecord2 :comp   // ссылка на спецификацию
 , cPers    :comp
 , Tabn :string[40]
 , Fio  :string[50]
 , Period :string[10]
 , dmd : array[1..100] of double
 , dms : array[1..100] of string
)
with index
( tmp_report_sp01  = nrec (AutoInc, Unique),
  tmp_report_sp02  = pr+sortname+Period,
  tmp_report_sp03  = cRec+sortname+Period,
  tmp_report_sp04  = pr+cRec+sortname+Period,
  tmp_report_sp05  = cRec+dms[1]+Period
);
//============
table struct tmp_memprint
(  col :word
 , ctmp_report_sp :comp
 , row :longint
 , subrow :longint
 , val    :string
) with index
(   tmp_memprint01 = col+row+subrow
  , tmp_memprint02 = ctmp_report_sp
);
//Declare _MacroGroupVarScud_
var a_GroupRepScud : array [1..cg_s_Len_a_gr] of word;
var sReportGRName  : array [1..cg_s_Len_a_gr] of string;
var sReportGRCode  : array [1..cg_s_Len_a_gr] of string;
var sReportGRNrec  : array [1..cg_s_Len_a_gr] of comp;
var sReportGRTBL   : array [1..cg_s_Len_a_gr] of word;
var dReportDay     : date;
var boMisConDuct_Gr: boolean;
var sMisConDuct    : string;
//3,4,5,6://3 'Не выявленные неявки',  4 'Потери рабочего времени',5 'Сверхурочные', 6 'Выходные'
var sKOD_ABS: string;// код неявки
var isFilterRun    : boolean=false;
var DT_Event       : date;
//end
//====================================
#Declare _MacroGroupInitScud_
  var ii:word;
  For(ii:=1;ii<=cg_s_Len_a_gr;ii++)
  { sReportGRName[ii]:='';
    sReportGRCode[ii]:='';
    sReportGRNrec[ii]:=-999;
    sReportGRTbl [ii]:=ii;
    if True //sReportGRNrec[ii]<>0
    { case ii of  //здесь нужно вычислить
      cg_s_code_Department:// 1 'Подразделение';
      { if isValidAll(tnKatPodr)
        { sReportGRName[ii]:=KatPodr.Name;
          sReportGRCode[ii]:=KatPodr.Kod;
          sReportGRNrec[ii]:=KatPodr.Nrec;
        }
      }
      cg_s_code_Org       :// 2 'Организация';
      { if GP_SCUD_PERS.wType=0
        { sReportGRName[ii]:='  собств.организация' ;
          sReportGRCode[ii]:='  ';
        }
        else
        if isValidAll(tnKatOrg)
        { sReportGRName[ii]:=KatOrg.Name;
          sReportGRCode[ii]:=KatOrg.ShortName;
          sReportGRNrec[ii]:=KatOrg.Nrec;
        }
        else
        { sReportGRName[ii]:=' Не задана организация';
          sReportGRCode[ii]:=' Н/з орг.';
        }
      }
      cg_s_code_Pers      :// 3 'Сотрудник';
      { sReportGRName[ii]:=GP_SCUD_PERS.Fio;
        sReportGRCode[ii]:=GP_SCUD_PERS.Code;
        sReportGRNrec[ii]:=GP_SCUD_PERS.Nrec;
      }
      cg_s_code_Date      :// 4 'Дата';
      { sReportGRName[ii]:=DateToStr(DT_Event,'YYYY-MM-DD');
        sReportGRCode[ii]:=DateToStr(DT_Event,'DD.MM.YYYY');
        sReportGRNrec[ii]:=pTSC_Event_Day.Nrec;
      }
      cg_s_code_Year      :// 5 'Год';
      { sReportGRName[ii]:=DateToStr(DT_Event,'YYYY');
        sReportGRCode[ii]:=sReportGRName[ii];
        sReportGRNrec[ii]:=pTSC_Event_Day.Nrec;
      }
      cg_s_code_quarter   :// 6 'квартал';
      { sReportGRName[ii]:=string((month(DT_Event)+2) div 3)+' кв';
        sReportGRCode[ii]:=sReportGRName[ii];
        sReportGRNrec[ii]:=pTSC_Event_Day.Nrec;
      }
      cg_s_code_month     :// 7 'Месяц';
      { sReportGRName[ii]:=DateToStr(DT_Event,'MM');
        sReportGRCode[ii]:=sReportGRName[ii];
        sReportGRNrec[ii]:=pTSC_Event_Day.Nrec;
      }
      cg_s_code_day       :// 8 'День';
      { sReportGRName[ii]:=DateToStr(DT_Event,'DD');
        sReportGRCode[ii]:=sReportGRName[ii];
        sReportGRNrec[ii]:=pTSC_Event_Day.Nrec;
      }
      cg_s_code_Post      :// 9 'Должность';
      { var _pPost:comp=0;
        sReportGRName[ii]:='-';
        sReportGRCode[ii]:='-';
        if iSValidAll(tnAppointments)
        { _pPost:=Appointments.Post;
          sReportGRName[ii]:=iCatFunc.GetCataLogsName(_pPost);
          sReportGRCode[ii]:=iCatFunc.GetCataLogsCode(_pPost);
          sReportGRNrec[ii]:=_pPost;
        }
      }
      cg_s_code_Rejim     ://10 'Режим';
      { if iSValidAll(tnKlrejim)
        { sReportGRName[ii]:=KlRejim.NRejim;
          sReportGRCode[ii]:=string(KlRejim.Rejim,5);
          sReportGRNrec[ii]:=KlRejim.Rejim;
        }

      }
      cg_s_code_Type      ://11 'сотрудники/организации';
      { if GP_SCUD_PERS.wType=0
        { sReportGRName[ii]:='Сотрудники.предприятия';
          sReportGRCode[ii]:='01' ;
        }
        else
        { sReportGRName[ii]:='сторонние организации';
          sReportGRCode[ii]:='02';
        }
        sReportGRNrec[ii]:=GP_SCUD_PERS.wType;
      }
      cg_s_code_Grup      ://12 'СКУД Группы сотрудников ';
      { if IsValidAll(tnGP_SCUD_GRUP)
        { sReportGRName[ii]:=GP_SCUD_GRUP.Name;
          sReportGRCode[ii]:=GP_SCUD_GRUP.Code;
          sReportGRNrec[ii]:=GP_SCUD_GRUP.Nrec;

        }
      }
      cg_s_code_SEX       ://13 'Пол';
      { if GP_SCUD_PERS.wType=0
        { sReportGRName[ii]:=Persons.Sex;
          sReportGRCode[ii]:=sReportGRName[ii];
        }
      }
      cg_s_code_Kateg     ://14 'СКУД Категории сотрудников';
      { if IsValidAll(tnGP_SCUD_GRUP)
        { sReportGRName[ii]:=GP_SCUD_GRUP.Name;
          sReportGRCode[ii]:=GP_SCUD_GRUP.Code;

        }
      }
      cg_s_code_MisConDuct://15 'нарушения дисциплины';
      { if wTipRep=0
        { sReportGRName[ii]:=pTSC_Event_Day.MisConDuct;
          if sReportGRName[ii]='' sReportGRName[ii]:=' нет';
        }
        else
        { sReportGRName[ii]:=sMisConDuct;
        }
        sReportGRCode[ii]:=sReportGRName[ii];
      }
      cg_s_code_MisConDuctEx://16 'наличие нарушений';
      { if boMisConDuct_exist
        { sReportGRName[ii]:='Есть';
        }
        else
        { sReportGRName[ii]:='Нет';
        }
        sReportGRCode[ii]:=sReportGRName[ii];
      }
      cg_s_code_OverWork  ://17 'наличие сверхурочных';
      { if boOverWork_exist
        { sReportGRName[ii]:='Есть';
        }
        else
        { sReportGRName[ii]:='Нет';
        }
        sReportGRCode[ii]:=sReportGRName[ii];
      }
      cg_s_code_OverWork2  ://18 'наличие переработки';
      { if boOverWork2_exist
        { sReportGRName[ii]:='Есть';
        }
        else
        { sReportGRName[ii]:='Нет';
        }
        sReportGRCode[ii]:=sReportGRName[ii];
      }
      cg_s_code_Holiday  ://19 'выходные';
      { if boOverWork_exist
        { sReportGRName[ii]:='Есть';
        }
        else
        { sReportGRName[ii]:='Нет';
        }
        sReportGRCode[ii]:=sReportGRName[ii];
      }
      cg_s_code_cex:// 20 'Цех';
      { if isValidAll(tnKatPodr)
        { GetKatPodrUp(Katpodr.nrec);
          sReportGRName[ii]:=KatPodr_Flt.Name;
          sReportGRCode[ii]:=KatPodr_Flt.Kod;
          sReportGRNrec[ii]:=KatPodr_Flt.Nrec;
        }
      }
      cg_s_code_abs:// 21 'код усл.обозначения';
      {  sReportGRName[ii]:=sKOD_ABS;// код неявки
         sReportGRCode[ii]:=sKOD_ABS;
         sReportGRNrec[ii]:=1;

      }
      end;

    } //if sReportGRNrec[ii]<>0
  }//For(ii:=1;ii<=cg_s_Len_a_gr;ii++)
#end
#Declare _MacroSortShapScud_
var iiiGR     : Integer;
var cRec_tbl  : Comp;
var sSort_tbl : String;

  sSort_tbl:='';
  cRec_tbl :=0;

  //Заполняем шапку tmp_report_gr
  for(iiiGR:=0; iiiGR<=word(cg_s_Len_a_gr); iiiGR++)
  {

    if (iiiGR=0)
    {
       sSort_tbl :='Всего';
    }
    else
    {  if a_GroupRepScud[iiiGR]>0 and a_GroupRepScud[iiiGR]<=cg_s_Len_a_gr
       {  sSort_tbl:=if(wSort=0
               ,sReportGRName[a_GroupRepScud[iiiGR]]
               ,sReportGRCode[a_GroupRepScud[iiiGR]]
             )
       } else continue;
    }

    if GetFirst fastfirstrow tmp_report_gr where (( //iiiGR       == tmp_report_gr.pr       and
                                                   sSort_tbl == tmp_report_gr.sortname and
                                                   cRec_tbl  == tmp_report_gr.cRec
                                                 )) <> tsOk
    { ClearBuffer(#tmp_report_gr);
      tmp_report_gr.pr      :=iiiGR;
      tmp_report_gr.sortname:=sSort_tbl;
      tmp_report_gr.cRec    :=cRec_tbl;

      if (iiiGR=0)
      {
         tmp_report_gr.LevCode := iiiGR;
         tmp_report_gr.LevName := sSort_tbl;
      }
      else
      { if a_GroupRepScud[iiiGR]>0 and a_GroupRepScud[iiiGR]<=cg_s_Len_a_gr
        { tmp_report_gr.LevCode := a_GroupRepScud[iiiGR];
          tmp_report_gr.LevName := sReportGRName[a_GroupRepScud[iiiGR]];
          tmp_report_gr.LevKod  := sReportGRCode[a_GroupRepScud[iiiGR]];
          tmp_report_gr.wTable  := a_GroupRepScud[iiiGR];
          tmp_report_gr.cRecord := sReportGRNrec[a_GroupRepScud[iiiGR]];
          tmp_report_gr.LevTitle:= GetLevelName(a_GroupRepScud[iiiGR]);
        }
      } //if
      Insert current tmp_report_gr;
      if logfilekl MyLogWrite(false,'>I_shap>'+String(iiiGR)+'|'+tmp_report_gr.sortname);
    } //if
    cRec_tbl  :=tmp_report_gr.Nrec; //!!!
  } //for
#end
//=======================================
//Цвет для уровня
//=======================================
Function GetColorLevel(icg:Integer):Longint;
{
/*
0FF0000h; //синий
000FF00h; //зеленый
00000FFh; //красный
*/
    case icg of
    cg_s_code_Department: Result:=0FFFFAAh; //Бледно синий
    // :=0DDDDDDh; //Бледно серый
    // :=0CCFFCCh; //Бледно зеленый
    //:=0AAFFFFh; //Бледно желтый
    //:=0FFAAFFh; //Бледно розовый
    //:=0FFDDDDh; //Бледно фиолетовый
    else  GetColorLevel:=0DDDDDDh; //Бледно серый
    end;
} //Function GetColorLevel
//=======================================
// Установка массива Группировки
//=======================================
Procedure MySetGroupMasScud(tmpStr:String);
var iii   :word;
{  boMisConDuct_Gr:=false;
   for(iii:=1;iii<=word(cg_s_Len_a_gr);iii:=iii+1)
   { a_GroupRepScud[iii]:=0;
   }
   for(iii:=1;iii<=word(cg_s_Len_a_gr);iii:=iii+1)
   { if(tmpStr<>'')
     { if(InStr('/',tmpStr)>0)
       { a_GroupRepScud[iii]:=Word(SubStr(tmpStr, 1, InStr('/',tmpStr)-1));
         tmpStr:=SubStr(tmpStr, InStr('/',tmpStr)+1, Length(tmpStr));
       }
       else
       { a_GroupRepScud[iii]:=word(tmpStr);
         tmpStr:='';
       }
       case a_GroupRepScud[iii] of
       cg_s_code_MisConDuct: boMisConDuct_Gr:=true;
       end;
     }
     else break;
   } //for
//   message('MySetGroupMasScud:'+a_GroupRepScud[1]+'/'+a_GroupRepScud[2]+'/'+a_GroupRepScud[3]+'/'+a_GroupRepScud[4]+'/'+a_GroupRepScud[5]);
} //Procedure MySetGroupMasScud
//=======================================
// Проверить выбрана Группировка
//=======================================
Function CheckGroupMasScud(wNoGr:word):Word;
var iii   :word;
{
   CheckGroupMasScud:=0;
   for(iii:=1;iii<=word(cg_s_Len_a_gr);iii:=iii+1)
   { if (a_GroupRepScud[iii]=wNoGr)
     { CheckGroupMasScud:=iii;
       exit;
     }
   }
} //Procedure CheckGroupMasScud

//=======================================
//Наименование уровня
//=======================================
Function GetLevelName(icg:word):string;
{ Result:='';
  case icg of
  cg_s_code_Department  :Result:=cg_s_name_Department  ;//'Подразделение';
  cg_s_code_Org         :Result:=cg_s_name_Org         ;//'Организация';
  cg_s_code_Pers        :Result:=cg_s_name_Pers        ;//'Сотрудник';
  cg_s_code_Date        :Result:=cg_s_name_Date        ;//'Дата';
  cg_s_code_Year        :Result:=cg_s_name_Year        ;//'Год';
  cg_s_code_quarter     :Result:=cg_s_name_quarter     ;//'квартал';
  cg_s_code_month       :Result:=cg_s_name_month       ;//'Месяц';
  cg_s_code_day         :Result:=cg_s_name_day         ;//'День';
  cg_s_code_Post        :Result:=cg_s_name_Post        ;//'Должность';
  cg_s_code_Rejim       :Result:=cg_s_name_Rejim       ;//'Режим';
  cg_s_code_Type        :Result:=cg_s_name_Type        ;//'сотрудники/организации';
  cg_s_code_Grup        :Result:=cg_s_name_Grup        ;//'СКУД Группы сотрудников ';
  cg_s_code_SEX         :Result:=cg_s_name_SEX         ;//'Пол';
  cg_s_code_Kateg       :Result:=cg_s_name_Kateg       ;//'СКУД Категории сотрудников';
  cg_s_code_MisConDuct  :Result:=cg_s_name_MisConDuct  ;//'нарушения дисциплины';
  cg_s_code_MisConDuctEx:Result:=cg_s_name_MisConDuctEx;//'наличие нарушений';
  cg_s_code_OverWork    :Result:=cg_s_name_OverWork    ;//'наличие сверхурочных';
  cg_s_code_OverWork2   :Result:=cg_s_name_OverWork2   ;//'наличие переработки';
  cg_s_code_Holiday     :Result:=cg_s_name_Holiday     ;//'наличие выходных';
  cg_s_code_cex         :Result:=cg_s_name_cex         ;//'цех';
  end;
}//Function GetLevelName(icg):string;
//=======================================
//Имя для уровня
//=======================================
Function GetGroupNameLevelScud(icg:word;sName:String):String;
{  Result:=if(sName='',UpCase(GetLevelName(icg) )+': нет',sName);
} //Function GetGroupNameLevelScud
