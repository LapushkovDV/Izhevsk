#doc
  ƒополнительное соглашение к трудовому договору
#end
.form NForm38Word
.hide
.fields
NDoc         
DateSost     
              
sBottom

FIO_RP
TabN
Dolgn_RP
Podr_RP

wChoice1
NOtpus1 
dBegin1 
dEnd1   
wDur1   
dRpBegin1
dRpEnd1 

wChoice2
NOtpus2 
dBegin2 
dEnd2   
wDur2   
dRpBegin2
dRpEnd2 

sPostUtvLico
sFIOUtvLico
sFIOOtvetstv
PrimOtvetstv
.endfields
^//NDoc         
^//DateSost     
              
^//sBottom

^//FIO_RP
^//TabN
^//Dolgn_RP
^//Podr_RP

.{ NForm38Word_CYCLE1 checkenter
^//wChoice1
^//NOtpus1 
^//dBegin1 
^//dEnd1   
^//wDur1   
^//dRpBegin1
^//dRpEnd1 
.}

.{ NForm38Word_CYCLE2 checkenter
^//wChoice2
^//NOtpus2 
^//dBegin2 
^//dEnd2   
^//wDur2   
^//dRpBegin2
^//dRpEnd2 
.}

^//sPostUtvLico
^//sFIOUtvLico
^//sFIOOtvetstv
^//PrimOtvetstv
.endform
