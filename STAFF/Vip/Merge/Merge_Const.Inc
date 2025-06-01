 const
   cgProff             = -22; // Профессии
   cgFormEducation     = -15; // Форма обучения 1
   cgLevelEducation    = -9;  // Вид обучения 1

   cgRelative          = -16; // Вид(Степень) родства 1
   cgNatureOfWork      = -12; // характер работы 2
   cgDocuments         = -3 ; // Виды документов 1
#Declare UpdateTableOneFieldWithOutIndex(Tabl,Field,TablString)
if Not(NextVisual) ObBreak;
StartNewVisual(vtRotateVisual,vfTimer,'Обработка таблицы '+#TablString,RecordsIntable(tn#Tabl));
 update #Tabl where ((SpJoi.cRec == #Tabl.#Field(noindex))) set #Tabl.#Field := JoiHead.cRec;
StopVisual('',0);
#End

#Declare UpdateTableOneFieldWithOutIndexSQL(Tabl,Field,TablString)
if Not(NextVisual) ObBreak;
StartNewVisual(vtRotateVisual,vfTimer,'Обработка таблицы '+#TablString,RecordsIntable(tn#Tabl));
SQL update #Tabl set #Tabl.#Field = :(JoiHead.cRec) where :(SpJoi.cRec) = #Tabl.#Field ;
StopVisual('',0);
#End

#Declare UpdateTableOneFieldWithIndex(Tabl,Field,TablString)
if Not(NextVisual) ObBreak;
StartNewVisual(vtRotateVisual,vfTimer,'Обработка таблицы '+#TablString,RecordsIntable(tn#Tabl));
 update #Tabl where ((SpJoi.cRec == #Tabl.#Field )) set #Tabl.#Field := JoiHead.cRec;
StopVisual('',0);
#End
   
