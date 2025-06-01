table struct GP_tZarHelZ
(
	_year     :	word,
	_mes      : word,
	_chasf    : string[20],
	_chasGr   : string[20],
	_SredSumma: string[20],
	_SredKoef : string[20],
  _SredPremi: string[20]
)
with index (
	i01 = _year + _mes
			)			
;
