#doc
  ƒополнительное соглашение к трудовому договору
#end
.form NForm35aWord
.hide
.fields
NN
Post
Dept
FIO
TabN

PostLead
DeptLead
FIOLead
TabNLead

FromDate
ToDate

NN1

FIOLead2

NN2

RukDolg
RukFIO

sPostUtvLico
sFIOUtvLico
sFIOOtvetstv
PrimOtvetstv
.endfields
.{ NForm35aWord_CYCLE checkenter

^//NN
^//Post
^//Dept
^//FIO
^//TabN

^//PostLead
^//DeptLead
^//FIOLead
^//TabNLead

^//FromDate
^//ToDate
.}

^//NN1

.{ NForm35aWord_CYCLE2 checkenter
^//FIOLead2
.}

^//NN2

^//RukDolg
^//RukFIO

^//sPostUtvLico
^//sFIOUtvLico
^//sFIOOtvetstv
^//PrimOtvetstv
.endform
