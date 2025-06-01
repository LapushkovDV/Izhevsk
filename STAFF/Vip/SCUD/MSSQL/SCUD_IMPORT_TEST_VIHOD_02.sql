/*
insert into V$GP_SCUD_EVENTS (
 F$NPP,F$sNPP,F$STRTABN,F$NAME
,F$CGP_SCUD_PERS,F$ENTERDATE,F$ENTERTIME
,F$CARD,F$CKATPODR,F$WEVENTS_TYPE,F$WRESOURCE
,F$CSTATUS,F$CCONTROLER,F$CAREA
)
*/
SELECT
 SCUD_EV.Id_reg NPP
,CAST(SCUD_EV.Id_reg as nvarchar) sNPP
,T$GP_SCUD_PERS.F$CODE STRTABN, SCUD_EV.last_name+' '+SCUD_EV.name+' '+SCUD_EV.fathership NAME
,T$GP_SCUD_PERS.F$Nrec CGP_SCUD_PERS,dbo.ToAtlDate(SCUD_EV.date_ev) ENTERDATE,dbo.ToAtlTime(SCUD_EV.time_ev) ENTERTIME
,SCUD_EV.Indetifier CARD,T$GP_SCUD_PERS.F$ckatpodr CKATPODR, 0 WEVENTS_TYPE
,case WHEN (SCUD_EV.areas_id =125937  or (SCUD_EV.areas_id =1318780 and SCUD_EV.id_configs_tree=1318770)) THEN 1 else 0 end WRESOURCE
,T$GP_STATUS.F$NREC CSTATUS
,T$GP_SCUD_CONTROLER.F$NREC CCONTROLER
,T$GP_SCUD_AREAS.F$NREC CAREA
-----------------------
, SCUD_EV.date_ev
, SCUD_EV.time_ev
, SCUD_EV.areas_id
, SCUD_EV.ip_addr
, SCUD_EV.type_pass
, Scud_ev.ct_name
---------------------------------
FROM OPENQUERY([SCUD],  '
 select distinct  
  --re.time_ev AS TIME_EV
  re.time_ev AS TIME_EV
 ,max(dpt.display_name) dcode,max(re.id_reg) AS ID_REG
 ,max(CASE when 1318629=ct.id_configs_tree then 1318774 else ct.id_configs_tree end) AS id_configs_tree
 --,max(re.type_pass) AS type_pass
 ,re.type_pass
 ,max(re.inner_number_ev) AS inner_number_ev
 ,max(sc.Identifier) as Indetifier,dpt.display_name dcode
 ,re.id_reg AS ID_REG
 ,CASE when 1318629=ct.id_configs_tree then 1318774 else ct.id_configs_tree end AS id_configs_tree
 --,max(re.type_pass) AS type_pass
 ,re.type_pass
 ,re.inner_number_ev AS inner_number_ev
 ,sc.Identifier as Indetifier
 ,re.date_ev,re.areas_id ,re.time_ev time_ev2
 ,st.last_name,st.first_name name, st.middle_name fathership, st.id_staff
 ---
,max(ip.ip_addr) ip_addr
--,max(cl.name)
 ---
 from REG_EVENTS re
inner join staff st on st.id_staff=re.staff_id
left outer join AREAS_TREE ar on (ar.ID_AREAS_TREE=re.AREAS_ID)
left outer join STAFF_REF sr on sr.staff_id=st.id_staff
left outer join STAFF_CARDS sc on sc.staff_id=st.id_staff
and sc.DATE_END>CURRENT_DATE-10
left outer join SUBDIV_REF dpt on dpt.id_ref=sr.subdiv_id
left outer join CONFIGS_TREE ct on ct.id_configs_tree=re.configs_tree_id_controller
Left outer join CONFIGS_TREE_LINKS cl on cl.configs_tree_id_child =ct.id_configs_tree 
left outer join CONFIGS_TREE_IP ip on ip.configs_tree_id = cl.configs_tree_id_parent 
left outer join  REG_EVENTS re_o2 on re_o2.staff_id=re.staff_id and re_o2.type_pass=1
and (re_o2.areas_id=125937 or re.areas_id =1318780) and (re_o2.areas_id<>re.areas_id)
and re_o2.date_ev=re.date_ev and re_o2.time_ev>=re.time_ev
where re.date_ev>=CURRENT_DATE-10  
 and 38223348 = re.Id_reg
--and re.type_pass=1 
--and (re.areas_id =1 )
----
and re.staff_id=1804067 -- À‡ÔÛ¯ÍÓ‚
--and re.staff_id=1563343 --  ÛÍÎÂ‚‡ ŒÎ¸„‡ »‚‡ÌÓ‚Ì‡, 12/12/2017 08:12
/*
and ( (ip.ip_addr is Null)
 or ( ip.ip_addr <>''10.18.1.139'' --
  and ip.ip_addr <>''10.18.1.152''
  and ip.ip_addr <>''10.18.1.153''
  
 )
 )
*/
 /*
and ((re_o2.id_reg is null) or re_o2.time_ev =
 (select min(re_o1.time_ev) as time_ev from REG_EVENTS re_o1
  where re_o1.type_pass=1 and (re_o1.areas_id=1 or re_o1.areas_id=125937 or re_o1.areas_id=1318780 )
  and re_o1.staff_id=re.staff_id and (re_o1.areas_id<>re.areas_id)
  and re_o1.date_ev=re.date_ev and re_o1.time_ev>=re.time_ev
  group by re_o1.staff_id,re_o1.date_ev --multiply,re_o1.areas_id
 ))*/ 
 group by re.date_ev,re.areas_id 
 --,re_o2.time_ev
 ,re.time_ev
 ,st.last_name,st.first_name , st.middle_name, st.id_staff
 ,re.type_pass
 ,re.Id_reg
 ') scud_ev
 inner join T$GP_TYPEDOCS on T$GP_TYPEDOCS.F$WTYPE =1009
 inner join  T$GP_STATUS on T$GP_STATUS.F$CTYPEDOC =T$GP_TYPEDOCS.F$NREC
       and T$GP_STATUS.F$NAME='— ”ƒ'
 left outer JOIN T$GP_SCUD_PERS ON T$GP_SCUD_PERS.F$EXT_KEY = cast(scud_ev.ID_STAFF AS nVarChar)
 left outer JOIN T$GP_SCUD_AREAS ON T$GP_SCUD_AREAS.F$ID_AREAS_TREE=scud_ev.areas_id
 left outer join T$GP_SCUD_CONTROLER ON T$GP_SCUD_CONTROLER.F$EXT_KEY = cast(scud_ev.id_configs_tree AS nVarChar)
 /*
 LEFT OUTER JOIN T$GP_SCUD_EVENTS ON T$GP_SCUD_EVENTS.F$CGP_SCUD_PERS =T$GP_SCUD_PERS.F$NREC
    AND T$GP_SCUD_EVENTS.F$ENTERDATE=dbo.ToAtlDate(SCUD_EV.date_ev)--F$ENTERDATE            --≥Ñ†‚† ¢ÂÆ§†                       ≥Date
    AND T$GP_SCUD_EVENTS.F$ENTERTIME =dbo.ToAtlDate(SCUD_EV.time_ev)--F$ENTERTIME            --≥Fhtvd ¢ÎÂÆ§†                     ≥Time
  where T$GP_SCUD_EVENTS.F$NREC is NULL
  */
order by  SCUD_EV.date_ev, SCUD_EV.time_ev
