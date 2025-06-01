const
   // DateInterimSpan constants
   DIS_NONE     = 0;
   DIS_YEAR     = 1;
   DIS_HALF     = 2;
   DIS_QUART    = 3;
   DIS_MONTH    = 4;
   DIS_WEEK     = 5;
   DIS_DAY      = 6;
   DIS_TIME     = 7;

   // DateInterimRepresent constants
   DIR_NONE     = 0;
   DIR_RELATIVE = 1;
   DIR_FIXED    = 2;
   DIR_UNLIMIT  = 3;

   // DateInterimType constants
   DIT_NONE     = 0;
   DIT_INTERVAL = 1;
   DIT_PERIOD   = 2;

   // DateInterimFields constants
   DIF_BEG      =    1;
   DIF_END      =    2;

   DIF_NONE     =    0;
   DIF_UNLIMIT  =    1;
   DIF_RELATIVE =    2;
   DIF_YEAR     =    4;
   DIF_HALF     =    8;
   DIF_QUART    =   16;
   DIF_MONTH    =   32;
   DIF_WEEK     =   64;
   DIF_DAY      =  128;
   DIF_TIME     =  256;

   DIF_FIXED    =  DIF_YEAR or DIF_HALF or DIF_QUART or DIF_MONTH or DIF_WEEK or DIF_DAY;


   // DateInterimError constants
   dieOk             =  0;
   dieInvalidInterim =  1;
   dieInvalidDate    =  2;
   dieIncorrectRep   =  3;
   dieIncorrectType  =  4;
   dieIncorrectSpan  =  5;
   dieBadDSKName     =  6;
   dieBadDSKVersion  =  7;
   dieBadDSKStruct   =  8;
   dieSaveError      =  9;
   dieEmptyDate      = 10;

   cdfNone           =  0;
   cdfFirst          =  1;
   cdfLast           =  2;

end;
