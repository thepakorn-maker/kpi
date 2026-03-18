<%
' ============================================================
' functions.asp — Shared helper functions
' Included by every page via: <!-- #include file="functions.asp" -->
' All functions here are available to every ASP page and
' to header.asp — do NOT redefine them anywhere else.
' ============================================================


' -- String helpers -------------------------------------------

' safeStr: HTML-encodes a value for safe output in page content
Function safeStr(s)
    If IsNull(s) Then safeStr = "" Else safeStr = Server.HTMLEncode(CStr(s))
End Function

' iif: inline if — returns trueVal when condition is true, falseVal otherwise
' Same as IIF() in other languages. Keeps templates concise.
Function iif(condition, trueVal, falseVal)
    If condition Then iif = trueVal Else iif = falseVal
End Function

' fmtDate: formats any date value as YYYY-MM-DD string
' Returns empty string for null/empty input
Function fmtDate(d)
    If IsNull(d) Or d = "" Then
        fmtDate = ""
    Else
        fmtDate = datepart("yyyy",d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
        'fmtDate=d
    End If
End Function


' -- KPI / Tier helpers ---------------------------------------
' These match the prototype getTier() logic exactly.

' getTier: returns tier name string from a KPI percentage
Function getTier(pct)
    Dim p : p = CDbl(pct)
    If p >= 95 Then
        getTier = "Platinum"
    ElseIf p >= 85 Then
        getTier = "Gold"
    ElseIf p >= 70 Then
        getTier = "Silver"
    ElseIf p >= 50 Then
        getTier = "Bronze"
    Else
        getTier = "Needs Work"
    End If
End Function

' getTierIcon: returns the emoji/HTML entity for a tier
Function getTierIcon(pct)
    Dim p : p = CDbl(pct)
    If p >= 95 Then
        getTierIcon = "&#x1F4A0;"  ' diamond  — Platinum
    ElseIf p >= 85 Then
        getTierIcon = "&#x1F947;"  ' gold medal
    ElseIf p >= 70 Then
        getTierIcon = "&#x1F948;"  ' silver medal
    ElseIf p >= 50 Then
        getTierIcon = "&#x1F949;"  ' bronze medal
    Else
        getTierIcon = "&#x26A0;"   ' warning sign — Needs Work
    End If
End Function

' getTierColor: returns the CSS hex color for a tier (used for text and badges)
Function getTierColor(pct)
    Dim p : p = CDbl(pct)
    If p >= 95 Then
        getTierColor = "#0ea5e9"   ' sky blue  — Platinum
    ElseIf p >= 85 Then
        getTierColor = "#f59e0b"   ' amber     — Gold
    ElseIf p >= 70 Then
        getTierColor = "#6b7280"   ' gray      — Silver
    ElseIf p >= 50 Then
        getTierColor = "#b45309"   ' brown     — Bronze
    Else
        getTierColor = "#dc2626"   ' red       — Needs Work
    End If
End Function


' -- Job display helpers --------------------------------------

' getJobTypeName: human-readable name for a job type ID
Function getJobTypeName(typeID)
    Select Case CStr(typeID)
        Case "-2"  : getJobTypeName = "Owner Emergency"
        Case "-1"  : getJobTypeName = "Owner Critical"
        Case "0"   : getJobTypeName = "Owner/Admin"
        Case "1"   : getJobTypeName = "Urgent/No Delay"
        Case "2"   : getJobTypeName = "Urgent"
        Case "3"   : getJobTypeName = "Important"
        Case "4.1" : getJobTypeName = "Plan-Gen"
        Case "4.2" : getJobTypeName = "Plan-Soln"
        Case "4.3" : getJobTypeName = "Plan-Fashion"
        Case "5"   : getJobTypeName = "Mr. Problem"
        Case Else  : getJobTypeName = "Type " & typeID
    End Select
End Function

' getUrgencyBadge: returns a colored HTML badge for urgent/emergency jobs
' Returns empty string for normal urgency (no badge shown)
Function getUrgencyBadge(urgency)
    Select Case LCase(CStr(urgency))
        Case "emergency"
            getUrgencyBadge = "<span style=""background:#dc2626;color:white;padding:2px 8px;" & _
                              "border-radius:4px;font-size:11px;margin-right:5px;font-weight:700"">" & _
                              "&#x1F534; EMERGENCY</span>"
        Case "urgent"
            getUrgencyBadge = "<span style=""background:#f59e0b;color:white;padding:2px 8px;" & _
                              "border-radius:4px;font-size:11px;margin-right:5px;font-weight:700"">" & _
                              "&#x26A0; URGENT</span>"
        Case Else
            getUrgencyBadge = ""  ' Normal urgency — no badge
    End Select
End Function


' -- Date helpers ---------------------------------------------

' nextMonday: returns the date of the next Monday
' Used for Type-5 (Mr. Problem) job auto-assignment
Function nextMonday()
    Dim d  : d  = Date()
    Dim wd : wd = Weekday(d, vbMonday)  ' 1 = Mon, 7 = Sun
    If wd = 1 Then
        nextMonday = d + 7   ' Today is Monday — use NEXT Monday
    Else
        nextMonday = d + (8 - wd)
    End If
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
