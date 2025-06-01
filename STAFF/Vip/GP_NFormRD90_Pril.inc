#component "GP_Staff"

#include Query.vih

table struct mtRD90Pr (
  DocNmb     : string[20],
  DocDate    : date,
  DepNrec    : comp,
  DepCode    : string[100],
  DepName    : string,
  cPerson    : comp,
  FIO        : string[60],
  PostName   : string,
  VidOpl     : string[10],
  WorkDate   : date,
  AllMemoNrec: comp
) with index (
  i01 = DepCode + WorkDate + FIO
);

table struct mtRD90PrGr (
  DepCode    : string[100],
  DepName    : string,
  WorkDate   : date
) with index (
  i01 = DepCode + WorkDate (Unique)
);

sql query qGP_NFormRD90_Pril =
    select
      TitleDoc.DocNmb,
      TitleDoc.DocDate,
      case when Length(Coalesce(CatDep.Code , '')) = 3 then CatDep.Nrec
           when Length(Coalesce(CatDep1.Code, '')) = 3 then CatDep1.Nrec
           when Length(Coalesce(CatDep2.Code, '')) = 3 then CatDep2.Nrec
           when Length(Coalesce(CatDep3.Code, '')) = 3 then CatDep3.Nrec
           when Length(Coalesce(CatDep4.Code, '')) = 3 then CatDep4.Nrec
           when Length(Coalesce(CatDep5.Code, '')) = 3 then CatDep5.Nrec
           when Length(Coalesce(CatDep6.Code, '')) = 3 then CatDep6.Nrec
           when Length(Coalesce(CatDep7.Code, '')) = 3 then CatDep7.Nrec
           when Length(Coalesce(CatDep8.Code, '')) = 3 then CatDep8.Nrec
           else Coalesce(CatDep9.Nrec, #comp(0))
      end as DepNrec,
      SubString(CatDep.Code, 1, 3) as DepCode,
      case when Length(Coalesce(CatDep.Code,  '')) = 3 then CatDep.Name
           when Length(Coalesce(CatDep1.Code, '')) = 3 then CatDep1.Name
           when Length(Coalesce(CatDep2.Code, '')) = 3 then CatDep2.Name
           when Length(Coalesce(CatDep3.Code, '')) = 3 then CatDep3.Name
           when Length(Coalesce(CatDep4.Code, '')) = 3 then CatDep4.Name
           when Length(Coalesce(CatDep5.Code, '')) = 3 then CatDep5.Name
           when Length(Coalesce(CatDep6.Code, '')) = 3 then CatDep6.Name
           when Length(Coalesce(CatDep7.Code, '')) = 3 then CatDep7.Name
           when Length(Coalesce(CatDep8.Code, '')) = 3 then CatDep8.Name
           else Coalesce(CatDep9.Name, '')
      end as DepName,
      Persons.Nrec as cPerson,
      Persons.FIO as FIO,
      Coalesce(CatPost.Name, '') as PostName,
      Coalesce(UOWrkTabel.Num, '') as VidOpl,
      Coalesce(SalScale.LastDate, 0) as WorkDate,
      Coalesce(AllMemo190.Nrec, #comp(0)) as AllMemoNrec
    from
      TitleDoc
    join PartDoc
      on PartDoc.cDoc = TitleDoc.Nrec
    join ContDoc
      on ContDoc.cPart = PartDoc.Nrec
    left join Catalogs CatDep
      on CatDep.Nrec = ContDoc.cDopRef
    left join Catalogs CatDep1
      on CatDep1.Nrec = CatDep.cParent
    left join Catalogs CatDep2
      on CatDep2.Nrec = CatDep1.cParent
    left join Catalogs CatDep3
      on CatDep3.Nrec = CatDep2.cParent
    left join Catalogs CatDep4
      on CatDep4.Nrec = CatDep3.cParent
    left join Catalogs CatDep5
      on CatDep5.Nrec = CatDep4.cParent
    left join Catalogs CatDep6
      on CatDep6.Nrec = CatDep5.cParent
    left join Catalogs CatDep7
      on CatDep7.Nrec = CatDep6.cParent
    left join Catalogs CatDep8
      on CatDep8.Nrec = CatDep7.cParent
    left join Catalogs CatDep9
      on CatDep9.Nrec = CatDep8.cParent
    join Persons
      on Persons.Nrec = ContDoc.Person
    left join Appointments App
      on App.Nrec = ContDoc.cStr
    left join Catalogs CatPost
      on CatPost.Nrec = App.Post
    join SalScale
      on SalScale.wTblCode = 0 and SalScale.cUnit = ContDoc.Nrec
    left join UOWrkTabel
      on UOWrkTabel.Nrec = Comp(SalScale.LastUser)
    left join (select
                 SalScale.Nrec as cSalScale190,
                 comp(AttrVal.vString) as cSalScale90: comp
               from AttrNam
               join AttrVal on AttrVal.cAttrNam = AttrNam.Nrec
               join SalScale on SalScale.Nrec = AttrVal.cRec
               where
                 AttrNam.wTable = 25054 and
                 AttrNam.Name = 'LinkedSalScaleNrec'
              ) tLink on tLink.cSalScale90 = SalScale.Nrec
    left join SalScale SalScale190
      on SalScale190.Nrec = tLink.cSalScale190
    left join AllMemo AllMemo190
      on AllMemo190.wTable = :coSalScale and AllMemo190.cRec = SalScale190.Nrec and AllMemo190.Npp = 0
    where TitleDoc.Nrec = :titleDocNrec;
