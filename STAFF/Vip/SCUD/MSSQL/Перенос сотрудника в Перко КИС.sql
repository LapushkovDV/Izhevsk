select s.id, s.DCODE, s.lastname, s.name, s.fathership , s.indate, s.citezenid
from HR_V_ACTIVE_EMPLOYEE_MAIN s where
s.citezenid NOT IN (select NVL(t.citezenid,0) from TC_STUFF t)
and s.dcode not in ('446') 

select
 GEN_ID(GENERAL_GENERATOR,1) staffid 
from  RDB$DATABASE

insert into 
STAFF(ID_STAFF, LAST_NAME, FIRST_NAME, MIDDLE_NAME, TABEL_ID, PORTRET, 
DATE_BEGIN, DATE_DISMISS, VALID,TEMPORARY_ACC, ID_FROM_1C, DELETED, STAFF_STATE,
 LAST_TIMESTAMP, DIN_PAY_SCHEMES_ID, DIN_GRAPHS_ID,DIN_GRANT_MEASURE_SPENDED, 
 DIN_COST_VALUE_SPENDED, PATH_ACTDIR, PATH_ACTDIR_LOGIN, PATH_ACTDIR_DOMAIN )
 VALUES
 (:staffid, :lastname, :name, :fathership, null, null, :indate, null,1,0, NULL, 0,0,:indate,0,0, 0,0,null,null,null) ;
 
 
 
 
select id_ref from SUBDIV_REF r where
r.display_name=:dcode 


 
-- ≈сли нет то 1 - "Ќе определено" 
 
 insert into STAFF_REF(ID_STAFF_REF, STAFF_ID, DATE_ACTION, SUBDIV_ID, APPOINT_ID, GROUP_WT_ID, DOCUMENTS_ID, LAST_TIMESTAMP)
 values(gen_id(gen_staff_ref_id,1), :staffid, :begindate, :deptid,1,1611478,0 ,:begindate)




 
select distinct     
     st.ID_STAFF, st.LAST_NAME, st.FIRST_NAME, st.MIDDLE_NAME
    ,st.deleted, st.date_begin, st.date_dismiss, sd.display_name
    ,st.id_from_1c
from
     STAFF as st  ,
     STAFF_REF as sr ,
     subdiv_ref    sd
where
    st.ID_STAFF = sr.STAFF_ID
    and sd.id_ref= sr.SUBDIV_ID
    and (st.id_from_1c='' or st.id_from_1c is null)
ORDER by
    st.ID_STAFF, sr.date_action
    
create or replace procedure TC_INSERT_STUFF
( xstuffid in number, xdelflag in number, xlastname in varchar, xname in varchar, 
  xfathership in varchar, xdatebegin in date, xdatedismiss in date, xdcode in varchar
 ,xid_from_1c in number
)
 is
 xcitezenid  number;
begin
  
 if (xid_from_1c is null) or (xid_from_1c=0) then
  begin
    select DISTINCT em.citezenid into xcitezenid from HR_V_ACTIVE_EMPLOYEE em where em.lastname=xlastname and
    em.name=xname and em.fathership=xfathership and em.dcode=xdcode;
  EXCEPTION
    when no_data_found then xcitezenid:=null; -- сотрудник не заведен в кадрах
    when too_many_rows then -- полные тески в одном отделе (и такое бывает )
      begin
       xcitezenid:=null;
       -- ' узьмина' x2
       if xstuffid not in (22960, 22967) then
        insert into TC_JOURNAL j values (null, sysdate,'staff with staff_id='||xstuffid||' has more than one citezenid','ERROR'); end if;
      end;
  end;
 ELSE
   BEGIN
   SELECT DISTINCT e.citezenid INTO xcitezenid FROM hr_v_active_employee_main e where e.lastname=xlastname and
    e.name=xname and e.fathership=xfathership;
   EXCEPTION
    when no_data_found then xcitezenid:=null; -- сотрудник не заведен в кадрах
    when too_many_rows then -- полные тески в одном отделе (и такое бывает )  
      insert into TC_JOURNAL j values (null, sysdate,'staff with staff_id='||xstuffid||' has more than one citezenid','ERROR');
   END;  
   --xcitezenid := xid_from_1c;
 end if;

 insert into TC_STUFF st values( AE_IDSEQ.Nextval, xcitezenid, xstuffid, xdelflag,
 xlastname, xname, xfathership, xdatebegin, xdatedismiss, xdcode, null );
 
 -- логируем
 insert into TC_JOURNAL j values (null, sysdate,'staff with staff_id='||xstuffid||' was inserted','INSERT_STAFF');

 EXCEPTION

   when DUP_VAL_ON_INDEX then -- запись уже есть. обновл€ем данные
     begin
      UPDATE  TC_STUFF st set st.deleteflag=xdelflag, st.lastname=xlastname, st.name=xname 
      , st.fathership=xfathership, st.datebegin=xdatebegin, st.datedismiss=xdatedismiss, st.dcode=xdcode
      WHERE st.stuffid=xstuffid;
     end;

end TC_INSERT_STUFF;
