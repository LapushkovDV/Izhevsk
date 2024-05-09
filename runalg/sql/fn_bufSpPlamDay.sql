USE [Russia]
GO
/****** Object:  UserDefinedFunction [dbo].[fn_bufSpPlamDay]    Script Date: 09.05.2024 5:19:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER FUNCTION [dbo].[fn_bufSpPlamDay]
        (
          @cMmPlam binary(8)
        , @dtPlan int
                , @sGUID varchar(40)
        )
RETURNS TABLE
AS
  RETURN
WITH SPPLAN as
(
SELECT
T$SpMnPl.F$CIZD,coalesce(T$KATMARSH.F$NREC , MK.F$NREC, NULL) as MKnrec, T$VALSPMNP.F$KOL, T$SpMnPl.F$NREC as SPnrec, T$SpMnPl.F$ENDDATE as dtEndPl
,T$SpMnPl.F$CANVAL1 as vidProd,T$SpMnPl.F$CANVAL2 as Model
FROM
T$MnPlan
JOIN T$SpMnPl on (T$SpMnPl.F$cMnPlan = T$MnPlan.F$NREC
and T$SpMnPl.F$ENDDATE=@dtPlan
)
JOIN T$VALSPMNP on (T$VALSPMNP.F$CSPMNPL = T$SpMnPl.F$NREC)
LEFT JOIN T$KATMARSH on (T$KATMARSH.F$NREC = T$SpMnPl.F$CANVAL3  )
LEFT JOIN T$KATMARSH MK on (MK.F$COBJECT  = T$SpMnPl.F$CIZD
and (MK.F$DTBEG<=@dtPlan OR MK.F$DTBEG = 0 OR MK.F$DTBEG IS NULL) AND (MK.F$DTEND>=@dtPlan OR MK.F$DTEND=0 OR MK.F$DTEND IS NULL))
WHERE
 T$MnPlan.f$Nrec = @cMmPlam
)
,MK_MC as(
SELECT
            N.f$cMaster   AS cMarsh_Sp
         ,  N.f$cResource AS cRes
         ,  N.f$Rasx      AS Qty
         ,  N.f$dNormEd   AS normQty
         ,  N.f$CED       As cEd
         ,  N.f$Nrec      AS nrec
         ,  M.f$cMarsh    AS MKnrec
         ,  M.f$TDEP          AS TIZG
         ,  M.f$CDEP          AS CIZG
    FROM t$Marsh_Sp M
    JOIN t$Normas AS N
    ON
         11005      = N.f$tMaster
    AND  M.f$Nrec   = N.f$cMaster
    AND  4                         = N.f$tResource
)
,SP_MK as(SELECT
t.SPnrec,fMK.kindID as MKnrec,fMK.TDEP,fMK.CDEP,fMK.Norm
FROM SPPLAN AS t
CROSS APPLY [dbo].[fn_razuzlMK](t.MKnrec,t.dtEndPl) AS fMK
UNION
SELECT
t.SPnrec, t.MKnrec,0 as TDEP, NULL as CDEP,t.F$KOL as Norm
FROM SPPLAN AS t
)
,GR_SP as(SELECT
t.vidProd, t.Model,t.dtEndPl,CASE WHEN SP_MK.TDEP>0 THEN SP_MK.TDEP else MK_MC.TIZG END as TPOTR,CASE WHEN SP_MK.TDEP>0 THEN SP_MK.CDEP else MK_MC.CIZG END as CPOTR,MK_MC.cRes as MCnrec,sum(t.F$KOL*SP_MK.Norm*MK_MC.Qty) as KolMK ,MK_MC.cEd, MK_MC.TIZG,MK_MC.CIZG
FROM SPPLAN AS t
left join SP_MK on (t.SPnrec = SP_MK.SPnrec )
left join MK_MC on (SP_MK.MKnrec = MK_MC.MKnrec )
GROUP BY t.vidProd, t.Model,t.dtEndPl,SP_MK.TDEP,SP_MK.CDEP,MK_MC.cRes,MK_MC.cEd, MK_MC.TIZG,MK_MC.CIZG
)
SELECT
 @sGUID as sRas,
 @cMmPlam as cDoc,
 case MC.F$KIND when 1 then 1 when 0 then MC.F$PRMAT+2 else 0 end  as tPlan,
 GR_SP.vidProd,
 GR_SP.Model,
 GR_SP.TPOTR,
 GR_SP.CPOTR,
 GR_SP.TIZG,
 GR_SP.CIZG,
 GR_SP.MCnrec as cMC,
 MC.F$KIND as KIND,
 MC.F$PRMAT as PRMAT,
 GR_SP.CED as cOtpEd,
 GR_SP.dtEndPl as dt,
 GR_SP.KolMK as Kol
FROM GR_SP
left join t$KatMC as MC on (MC.F$NREC = GR_SP.MCnrec)

