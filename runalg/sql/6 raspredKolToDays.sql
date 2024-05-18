/****** Object:  UserDefinedFunction [dbo].[fn_raspredKolToDays]    Script Date: 18.05.2024 8:57:58 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER     function [dbo].[fn_raspredKolToDays](
                                                                  @dtPlan INT
                                                                , @DELIM INT
                                                                , @cGRAFIK binary(8)
                                                                , @KOL NUMERIC
        )
RETURNS TABLE
AS
  RETURN

WITH WorkDay as(
SELECT ROW_NUMBER() over(order by F$DDATE ASC) as num, F$DDATE as Dt
FROM T$SPGRAF
WHERE  T$SPGRAF.F$CGRAFIK = @cGRAFIK
        and F$DDATE > ([dbo].[ymd2date]([dbo].[dtYEAR]([dbo].[ToAtlDateTime2](@dtPlan,0)),[dbo].[dtMONTH]([dbo].[ToAtlDateTime2](@dtPlan,0)),1))
        and F$DDATE < ([dbo].[ymd2date]([dbo].[dtYEAR]([dbo].[ToAtlDateTime2](@dtPlan,0)),[dbo].[dtMONTH]([dbo].[ToAtlDateTime2](@dtPlan,0)),1) + 24)
),
RoundedValues AS (
    SELECT
        num, CEILING(@KOL / (SELECT COUNT(*)
 FROM WorkDay)) AS Kol, Dt
                FROM WorkDay
),
SumRounded AS (
    SELECT
        SUM(Kol) AS TotalRoundedSum
    FROM
        RoundedValues
),
Adjustment AS (
    SELECT
        CAST((SELECT COUNT(*)
 FROM WorkDay)+(@KOL - TotalRoundedSum) as INT) AS cnt
    FROM
        SumRounded
)

select * from
(
Select  num
                ,case @DELIM when 1 then (case when num <= (SELECT cnt FROM Adjustment) then
                        (CEILING(@KOL / (SELECT COUNT(*)
 FROM WorkDay))) else ROUND(@KOL / (SELECT COUNT(*)
 FROM WorkDay),0,1) end)
                        else (@KOL / (SELECT COUNT(*)
 FROM WorkDay)) end as Kol
                ,Dt
from WorkDay
) t where t.Kol > 0


