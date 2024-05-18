/****** Object:  StoredProcedure [dbo].[S$ENUMERATE_SPMNPLAN]    Script Date: 18.05.2024 9:10:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[S$ENUMERATE_SPMNPLAN](@cMmPlam binary(8))

AS
BEGIN


IF OBJECT_ID('tempdb ..#xx$locks')IS NOT NULL drop table #xx$locks
create table #xx$locks (TableNRec binary(8))

update T$SPMNPLAN
set T$SPMNPLAN.F$NUMBER = t.npp
from T$SPMNPLAN
join (
select
        format(ROW_NUMBER() over(order by mc.F$BARKOD),'000000') as npp
 , spplan.F$NUMBER as number
 , spplan.F$NREC as cmnplan
 , mc.F$BARKOD as barkod
from t$mnplan mnplan
join T$SPMNPLAN spplan on spplan.F$CMNPLAN = mnplan.F$NREC
JOIN T$KATMC mc on mc.F$NREC = spplan.F$CIZD
where mnplan.F$NREC = @cMmPlam
) t on t.cmnplan = T$SPMNPLAN.F$NREC

end
