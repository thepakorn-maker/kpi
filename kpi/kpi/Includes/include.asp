<% 
Function IIf(i,j,k) 
If i Then 
IIf = j 
Else 
IIf = k 
end if
End Function

function sys_date(date1)   
'ถ้า date1=0 แปลงวันที่ปัจจุบันของระบบให้เป็นรูปแบบ วัน/เดือน/ปี ค.ศ.
'ถ้า date1=วันที่ ก็จะแปลงวันที่ให้เป็นรูปแบบ วัน/เดือน/ปี ค.ศ.
if not(isnull(date1)) then
if date1="0" then
  date1=date
elseif not(isdate(date1)) then
  sys_date=""
  exit function
end if
dayTemp=datepart("d",date1)
'if len(dayTemp)<2 then
'  dayTemp="0" & dayTemp
'end if
yearTemp=datepart("yyyy",date1)
if yearTemp<2000 then
  yearTemp=yearTemp+543
end if
monthTemp=datepart("m",date1)
'if len(monthTemp)<2 then
'  monthTemp="0" & monthTemp
'end if

sys_date=dayTemp & "/" & monthTemp & "/" & yearTemp
end if
end function

function ora_date(dateTemp) 'แปลง DD/MM/YYYY เป็น DD-MON-YYYY
if Trim(dateTemp)="" then
 ora_date=""
 exit function
end if
myparam=Split(trim(dateTemp),"/")
if ubound(myparam)<2 or instr("0123456789",left(trim(dateTemp),1))<=0 then
 ora_date=""
 exit function
end if

dayTemp=clng(myparam(0))
yearTemp=clng(myparam(2))
if len(yearTemp)=2 then
  yearTemp=yearTemp+2000
end if
if yearTemp>2500 then
  yearTemp=yearTemp-543
elseif yearTemp>datepart("yyyy",date)+5 then
  yearTemp=yearTemp-43
end if

select case clng(myparam(1))
case 1
monthTemp="JAN"
case 2
monthTemp="FEB"
case 3
monthTemp="MAR"
case 4
monthTemp="APR"
case 5
monthTemp="MAY"
case 6
monthTemp="JUN"
case 7
monthTemp="JUL"
case 8
monthTemp="AUG"
case 9
monthTemp="SEP"
case 10
monthTemp="OCT"
case 11
monthTemp="NOV"
case 12
monthTemp="DEC"
end select

ora_date=dayTemp & "-" & monthTemp & "-" & yearTemp
end function


FUNCTION ISO_DATE(arg)
if isdate(arg) then
  'a=iif(1=1,1,0)
  ISO_DATE=iif(clng(datepart("yyyy",arg))>1900,datepart("yyyy",arg),datepart("yyyy",arg)+543) & "-" & right("0" & datepart("m",arg),2) & "-" & right("0" & datepart("d",arg),2)
elseif arg<>"" then
  myparam=Split(arg,"/")
  if ubound(myparam)=2 then
    ISO_DATE=myparam(2) & "-" & right("0" & myparam(1),2) & "-" & right("0" & myparam(0),2)
  else  
    ISO_DATE=""
  end if
else
  ISO_DATE=""
end if

END FUNCTION
%>