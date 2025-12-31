<%@ Language=VBScript %>

<!--#include file="db/connect.asp" -->
<!--#include file="Includes/include.asp" -->

<%

If request.cookies("UserID") = "" Then Response.Redirect "login.asp"

Dim title, typeID, assignedTo, assignedBy, targetDate, isPrivate
title = Trim(Request.Form("title"))
typeID = Request.Form("typeID")
assignedTo = CLng(Request.Form("assignedTo"))
assignedBy = CLng(request.cookies("UserID"))
targetDate = (iso_date(Request.Form("targetDate")))
'response.write targetDate
'response.end
isPrivate = IIf(Request.Form("isPrivate") = "1", 1, 0)

' ถ้า Type 5 ==> Force วันจันทร์ถัดไป
If typeID = "5" and false Then
    Dim todayDate, dayOfWeek, daysToMonday
    todayDate = Date()
    dayOfWeek = Weekday(todayDate)
    daysToMonday = IIf(dayOfWeek = 2, 7, 9 - dayOfWeek) ' จันทร์ = 2
    targetDate = DateAdd("d", daysToMonday, todayDate)
End If

Dim sqlInsert
sqlInsert = "INSERT INTO Jobs (Title, TypeID, AssignedTo, AssignedBy, TargetDate, IsPrivate) " & _
            "VALUES (?, ?, ?, ?, ?, ?)"

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = sqlInsert
cmd.Parameters.Append cmd.CreateParameter("@title", 200, 1, 500, title)
cmd.Parameters.Append cmd.CreateParameter("@typeID", 200, 1, 10, typeID)
cmd.Parameters.Append cmd.CreateParameter("@assignedTo", 3, 1, , assignedTo)
cmd.Parameters.Append cmd.CreateParameter("@assignedBy", 3, 1, , assignedBy)
cmd.Parameters.Append cmd.CreateParameter("@targetDate", 135, 1, , targetDate)
cmd.Parameters.Append cmd.CreateParameter("@isPrivate", 11, 1, , isPrivate)

cmd.Execute

Response.Redirect "dashboard.asp"
%>