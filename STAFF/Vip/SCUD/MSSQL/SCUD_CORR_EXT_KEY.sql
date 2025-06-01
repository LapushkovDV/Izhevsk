--exec Create#xx$locks
IF OBJECT_ID('tempdb ..#xx$locks')IS NOT NULL drop table #xx$locks
create table #xx$locks (TableNRec binary(8))

UPDATE P SET P.F$PROPUSK_NMB=scud_pass.IDENTIFIER
    ,P.F$PROPUSK_FROM=dbo.ToAtlDate(scud_pass.DATE_BEGIN)
	,P.F$PROPUSK_TO=dbo.ToAtlDate(scud_pass.DATE_END)
	,P.F$EXT_KEY=cast(scud_pass.STAFF_ID AS nVarChar)
	
--select scud_pass.*,P.F$EXT_KEY
FROM OPENQUERY([SCUD],  'select 
           sc_N.ID_CARD, sc_N.STAFF_ID, sc_N.DATE_BEGIN, sc_N.DATE_END, sc_N.VALID, sc_N.VALID_TRANSFER
		 , sc_N.TEMPORARY_ACC, sc.DOCUMENTS_ID, sc.HISTORY_DATE
         , sc_N.PROHIBIT, sc_N.IDENTIFIER, sc_N.TYPE_IDENTIFIER
         , sc_N.IDENTIFIER_TRANSFORMED
         , sc_N.WITHDRAW_TO_STOP_LIST, sc_N.LAST_TIMESTAMP, sc_N.USER_ID
		 , st_N.last_name,st_N.first_name name, st_N.middle_name
         from staff st
         left outer join STAFF_CARDS sc on st.id_staff = sc.staff_id
         inner join staff st_N  on st_N.last_name=st.last_name
          and st_N.first_name=st.first_name
          and st_N.middle_name=st.middle_name
         and st_N.id_staff <> st.id_staff
         inner join STAFF_CARDS sc_N on st_N.id_staff = sc_N.staff_id
         where sc.ID_CARD is Null                  and
         SC_N.DATE_BEGIN in
          (select max(sc2.DATE_BEGIN) from STAFF_CARDS SC2
           where SC2.staff_id = sc_N.staff_id
           )

') scud_pass
inner join T$GP_SCUD_PERS P on P.F$EXT_KEY=cast(scud_pass.staff_id as nvarchar)
       or P.F$FIo = SCUD_PASS.last_name+' '+SCUD_PASS.name+' '+SCUD_pass.middle_name
    

--where scud_pass.last_name='Моисеенков' and scud_pass.name='Дмитрий'
