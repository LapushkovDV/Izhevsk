go
drop FUNCTION [dbo].[TO_SQLDATE]
go
drop FUNCTION [dbo].[TO_SQLDATE_S]
go
drop FUNCTION [dbo].[TO_SQLTime]
go
create FUNCTION  TO_SQLDATE(@lastdate int )
returns date
as
begin 
 declare @datedisplay char(23) 
 declare @date datetime
 set @date = convert(datetime, 
 convert(char(4),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),1,2)))+'-'+ 
 rtrim(convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),3,1))))+'-'+ 
 rtrim(convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),4,1))))
  ,120) 
 --set @date= rtrim(convert(char(2),DATEPART(DD,@datedisplay)))+'-'+ 
 -- rtrim(convert(char(2),DATEPART(MM,@datedisplay)))+'-'+  
 -- convert(char(4),DATEPART(YYYY,@datedisplay)) 
   RETURN(@date)
END
go
--======================================
create FUNCTION  [dbo].[TO_SQLDATE_S](@lastdate int )
returns char(23) 
as
begin 
 declare @datedisplay char(23) 
 declare @date date
 set @datedisplay =  convert(char(4),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),1,2)))+'-'+ 
 rtrim(convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),3,1))))+'-'+ 
 rtrim(convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lastdate),4,1)))) 
   RETURN(@datedisplay)
END
go
--======================================
create Function TO_SQLTIME(@lasttime int )
returns time
as
begin
declare @timedisplay char(23) 
declare @time time
set @timedisplay = convert(datetime, 
convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lasttime),1,1)))+':'+ 
convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lasttime),2,1)))+':'+ 
convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lasttime),3,1)))+':'+ 
convert(char(2),convert(int,SUBSTRING(CONVERT(BINARY(4),@lasttime),4,1))) ,114) 

set @time= rtrim(convert(char(2),DATEPART(HH,@timedisplay)))+':'+ 
rtrim(convert(char(2),DATEPART(MI,@timedisplay)))+':'+ 
rtrim(convert(char(2),DATEPART(SS,@timedisplay)))+':'+ 
rtrim(convert(char(2),DATEPART(MS,@timedisplay))) 
return @time
end