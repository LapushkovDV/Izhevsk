USE [Russia]
GO
/****** Object:  StoredProcedure [dbo].[S$BUFSPPLAM]    Script Date: 09.05.2024 5:20:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[S$BUFSPPLAM](@cMmPlam binary(8), @sGUID varchar(38))

AS
BEGIN
--declare @cMmPlam binary(8)
-- s$bufSpPlam 0x800000000000005B,'DA0478CD-25BD-4E29-865A-B4E612777B62'
-- Loop through each date in the DataPlan table
DECLARE @DataT int;

DECLARE cur CURSOR FOR
SELECT DISTINCT T$SpMnPl.F$ENDDATE as DataT
                        FROM  T$SpMnPl JOIN T$VALSPMNP on (T$VALSPMNP.F$CSPMNPL = T$SpMnPl.F$NREC)
                        where T$SpMnPl.F$cMnPlan = @cMmPlam and T$VALSPMNP.F$KOL > 0;

OPEN cur;
FETCH NEXT FROM cur INTO @DataT;

WHILE @@FETCH_STATUS = 0
BEGIN
                INSERT INTO V$BUFSPPLAM (f$sRas, f$cDoc, f$tPlan, f$vidProd, f$Model, f$TPOTR, f$CPOTR, f$TIZG, f$CIZG, f$cMC, f$KIND, f$PRMAT, f$cOtpEd, f$dt, f$Kol)
                Select bufSpPlam.*
                FROM [dbo].[fn_bufSpPlamDay](@cMmPlam,@DataT,@sGUID) as bufSpPlam;

FETCH NEXT FROM cur INTO @DataT;
END;

CLOSE cur;
DEALLOCATE cur;
END;

