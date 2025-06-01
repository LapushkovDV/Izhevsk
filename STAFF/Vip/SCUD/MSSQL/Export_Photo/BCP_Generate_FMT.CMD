rem на первый вопрос - I
rem на второй 0
rem остальные Enter Кроме вопроса создания файла
bcp "select t.portret from erpdb_old.dbo.PhotoForGalFromPerco t where ID_STAFF = 1533769 " queryout "c:\Galaxy\ZxTime\Export_Photo\SQL\Fhoto\1533769.jpg" -T  -S dep968-galsrv\erp