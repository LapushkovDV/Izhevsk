SELECT
 SCUD_EV.Id_reg NPP
,CAST(SCUD_EV.Id_reg as nvarchar) sNPP
,T$GP_SCUD_PERS.F$CODE STRTABN, SCUD_EV.last_name+' '+SCUD_EV.name+' '+SCUD_EV.fathership NAME
,T$GP_SCUD_PERS.F$Nrec CGP_SCUD_PERS
,dbo.ToAtlDate(SCUD_EV.date_ev) ENTERDATE
,dbo.ToAtlTime(SCUD_EV.time_ev) ENTERTIME
,SCUD_EV.date_ev
,SCUD_EV.time_ev
,SCUD_EV.Indetifier CARD,T$GP_SCUD_PERS.F$ckatpodr CKATPODR
, 0 WEVENTS_TYPE
,case WHEN (SCUD_EV.areas_id =125937 or SCUD_EV.areas_id =1318780) THEN 1 else 0 end WRESOURCE
,T$GP_STATUS.F$NREC CSTATUS
,T$GP_SCUD_CONTROLER.F$NREC CCONTROLER
,T$GP_SCUD_AREAS.F$NREC CAREA
FROM OPENQUERY([SCUD],  'select re.time_ev as time_ev 
 ,min(dpt.display_name) dcode
 ,re.id_reg  Id_reg, 
 max(CASE when 1318629=ct.id_configs_tree then 1318770 else ct.id_configs_tree end) AS id_configs_tree
 ,min(re.type_pass) type_pass, min(re.inner_number_ev) inner_number_ev
 ,sc.Identifier as Indetifier,re.date_ev, re.areas_id 
 --,re_o2.time_ev time_ev2
 ,st.last_name,st.first_name name, st.middle_name fathership, st.id_staff
 from REG_EVENTS re  inner join staff st on st.id_staff = re.staff_id
 --inner join AREAS_TREE ar  on (ar.ID_AREAS_TREE=re.AREAS_ID)
 left outer join AREAS_TREE ar  on (ar.ID_AREAS_TREE=re.AREAS_ID)
 left outer join STAFF_REF sr on sr.staff_id=st.id_staff
 left outer join STAFF_CARDS sc on sc.staff_id=st.id_staff
 and sc.DATE_END>CURRENT_DATE-10
 left outer join SUBDIV_REF dpt on dpt.id_ref=sr.subdiv_id
 left outer join CONFIGS_TREE ct on ct.id_configs_tree = re.configs_tree_id_controller
 Left outer join CONFIGS_TREE_LINKS cl on cl.configs_tree_id_child =ct.id_configs_tree 
 left outer join CONFIGS_TREE_IP ip on ip.configs_tree_id = cl.configs_tree_id_parent 
 where 
 --re.type_pass=1  and 
 --  (re.areas_id =125937 or re.areas_id =1318780)   and 
  re.staff_id=10411060
  -- re.id_reg=40149999
  --sc.Identifier=9651185
  and re.date_ev>=CURRENT_DATE-50
  --and SubString(ct.display_name from 1 for 2)<>''''КБ'''' -- не будем брать входы с КБ
  group by re.date_ev,re.areas_id ,re.time_ev -- было re_o2.time_ev
,st.last_name ,st.first_name, st.middle_name, st.id_staff
,sc.Identifier
,re.id_reg
 ') scud_ev
 inner join T$GP_TYPEDOCS on T$GP_TYPEDOCS.F$WTYPE =1009
 inner join  T$GP_STATUS on T$GP_STATUS.F$CTYPEDOC =T$GP_TYPEDOCS.F$NREC
       and T$GP_STATUS.F$NAME='СКУД'
 Inner JOIN T$GP_SCUD_PERS ON T$GP_SCUD_PERS.F$EXT_KEY = cast(scud_ev.ID_STAFF AS nVarChar)
 inner JOIN T$GP_SCUD_AREAS ON T$GP_SCUD_AREAS.F$ID_AREAS_TREE=scud_ev.areas_id
 inner join T$GP_SCUD_CONTROLER ON T$GP_SCUD_CONTROLER.F$EXT_KEY = cast(scud_ev.id_configs_tree AS nVarChar)
     --and SubString(T$GP_SCUD_CONTROLER.F$name, 1, 2)<>''КБ'' -- не будем брать входы с КБ
	 and T$GP_SCUD_CONTROLER.F$wType<>1
 LEFT OUTER JOIN T$GP_SCUD_EVENTS ON T$GP_SCUD_EVENTS.F$CGP_SCUD_PERS =T$GP_SCUD_PERS.F$NREC
    AND T$GP_SCUD_EVENTS.F$ENTERDATE=dbo.ToAtlDate(SCUD_EV.date_ev)--F$ENTERDATE            --і„ в  ўе®¤                        іDate
    AND T$GP_SCUD_EVENTS.F$ENTERTIME =dbo.ToAtlTime(SCUD_EV.time_ev)--F$ENTERTIME            --іFhtvd ўле®¤                      іTime
  --where T$GP_SCUD_EVENTS.F$NREC is NULL
