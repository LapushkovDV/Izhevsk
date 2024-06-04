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
JOIN T$VALSPMNP on (T$VALSPMNP.F$CSPMNPL = T$SpMnPl.F$NREC and T$VALSPMNP.F$KOL > 0)
LEFT JOIN T$KATMARSH on (T$KATMARSH.F$NREC = T$SpMnPl.F$CANVAL3  )
LEFT JOIN T$KATMARSH MK on (MK.F$COBJECT  = T$SpMnPl.F$CIZD and MK.F$ACTIVE = 1
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
                 ,   K.F$KIND as KIND
                ,   K.F$PRMAT as PRMAT
                ,   M.f$NUM          AS NUM
                ,   M_p.F$NUM  AS NumPotr
         ,  M_p.f$TDEP          AS tPotr
         ,  M_p.f$CDEP          AS cPotr
                ,   M_i.F$NUM   AS NumIzg
         ,  M_i.f$TDEP          AS tIzg21
         ,  M_i.f$CDEP          AS cIzg21
    FROM t$Marsh_Sp M
    JOIN t$Normas AS N ON (  N.f$tMaster = 11005
                             AND M.f$Nrec = N.f$cMaster
                                                 AND N.f$tResource = 4
                                                 AND N.F$TOBJECT = 11007
                                                 and (   (N.F$DTBEG<=@dtPlan OR N.F$DTBEG=0 OR N.F$DTBEG IS NULL)
                                                     AND (N.F$DTEND>=@dtPlan OR N.F$DTEND =0 OR N.F$DTEND IS NULL)
                                                         )
                                                 and coalesce(N.f$Rasx,0) > 0)
    JOIN t$katmc as K ON (N.f$cResource = K.f$Nrec and K.F$KIND < 2)
    left JOIN t$katotped as Ed ON (N.f$CED = Ed.f$Nrec)
    JOIN t$katotped as edA ON (1 = edA.f$PRMC AND K.f$Nrec = edA.f$CMCUSL and K.f$CED = edA.f$CKATED)
        outer apply (select top 1 M_potr.F$NUM, M_potr.F$TDEP,M_potr.F$CDEP from t$Marsh_Sp as M_potr where (M.F$CMARSH = M_potr.F$CMARSH and M.F$NUM < M_potr.F$NUM) order by M_potr.F$NUM) M_p
--?? пробить изготовителя в МК для быстродействия ??        left JOIN T$KATMARSH
        outer apply (select top 1 M_izg.F$NUM, M_izg.F$TDEP,M_izg.F$CDEP from t$Marsh_Sp as M_izg where (N.F$CDOC = M_izg.F$CMARSH ) order by M_izg.F$NUM DESC) M_i
)
,SP_MK as(SELECT
t.SPnrec,fMK.kindID as MKnrec,fMK.TDEP,fMK.CDEP,fMK.Norm
FROM SPPLAN AS t
CROSS APPLY [dbo].[fn_razuzlMK](t.MKnrec,t.dtEndPl) AS fMK
UNION ALL
SELECT
t.SPnrec, t.MKnrec,0 as TDEP, NULL as CDEP, 1 as Norm
FROM SPPLAN AS t
)
,N_SP as(
SELECT
t.vidProd, t.Model,t.dtEndPl,
   MK_MC.KIND , MK_MC.PRMAT,
--CASE WHEN SP_MK.TDEP>0 THEN SP_MK.TDEP else MK_MC.TIZG END as TPOTR,CASE WHEN SP_MK.TDEP>0 THEN SP_MK.CDEP else MK_MC.CIZG END as CPOTR
CASE (MK_MC.KIND+MK_MC.PRMAT) WHEN 0 THEN MK_MC.TIZG WHEN 2 THEN (CASE WHEN SP_MK.TDEP>0 THEN SP_MK.TDEP else MK_MC.TIZG END) ELSE MK_MC.tPotr END as TPOTR,
CASE (MK_MC.KIND+MK_MC.PRMAT) WHEN 0 THEN MK_MC.CIZG WHEN 2 THEN (CASE WHEN SP_MK.TDEP>0 THEN SP_MK.CDEP else MK_MC.CIZG END) ELSE MK_MC.cPotr END as CPOTR
,MK_MC.cRes as MCnrec,(t.F$KOL*SP_MK.Norm*MK_MC.Qty) as KolMK ,MK_MC.cEd,
CASE (MK_MC.KIND+MK_MC.PRMAT) WHEN 2 THEN (CASE WHEN MK_MC.tIzg21>0 THEN MK_MC.tIzg21 else MK_MC.TIZG END) ELSE MK_MC.TIZG  END as TIZG,
CASE (MK_MC.KIND+MK_MC.PRMAT) WHEN 2 THEN (CASE WHEN MK_MC.tIzg21>0 THEN MK_MC.cIzg21 else MK_MC.CIZG END) ELSE MK_MC.CIZG END as CIZG
FROM SPPLAN AS t
left join SP_MK on (t.SPnrec = SP_MK.SPnrec )
left join MK_MC on (SP_MK.MKnrec = MK_MC.MKnrec )

)
,GR_SP as(
SELECT
N_SP.vidProd, N_SP.Model,N_SP.dtEndPl,
N_SP.KIND , N_SP.PRMAT,
N_SP.TPOTR, N_SP.CPOTR
,N_SP.MCnrec,sum(N_SP.KolMK) as KolMK,N_SP.cEd, N_SP.TIZG,N_SP.CIZG
FROM N_SP
GROUP BY N_SP.vidProd, N_SP.Model,N_SP.dtEndPl,N_SP.KIND , N_SP.PRMAT,
N_SP.TPOTR, N_SP.CPOTR
,N_SP.MCnrec,N_SP.cEd, N_SP.TIZG,N_SP.CIZG
)
SELECT
 @sGUID as sRas,
 @cMmPlam as cDoc,
 case GR_SP.KIND when 1 then 1 when 0 then GR_SP.PRMAT+2 else 0 end  as tPlan,
 GR_SP.vidProd,
 GR_SP.Model,
 GR_SP.TPOTR,
 GR_SP.CPOTR,
 GR_SP.TIZG,
 GR_SP.CIZG,
 GR_SP.MCnrec as cMC,
 GR_SP.KIND as KIND,
 GR_SP.PRMAT as PRMAT,
 GR_SP.CED as cOtpEd,
 GR_SP.dtEndPl as dt,
 GR_SP.KolMK as Kol
FROM GR_SP
