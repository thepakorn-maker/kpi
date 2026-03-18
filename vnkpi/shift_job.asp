<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
' ============================================================
' shift_job.asp — POST action: shift a job's target date
' Flow: validate inputs -> verify ownership -> call sp_ShiftJobDate
' Redirects back to myjobs.asp on any error.
' ============================================================

Dim jobID   : jobID   = CLng(Request.Form("JobID"))
Dim newDate : newDate = Trim(Request.Form("NewDate"))
Dim reason  : reason  = Trim(Request.Form("Reason"))

' Basic input check
If jobID = 0 Or newDate = "" Then
    Response.Redirect "myjobs.asp"
    Response.End
End If

' Parse new date — redirect if malformed
Dim ndDate
On Error Resume Next
ndDate = CDate(newDate)
If Err.Number <> 0 Then
    On Error GoTo 0
    Response.Redirect "myjobs.asp"
    Response.End
End If
On Error GoTo 0

' New date cannot be in the past
If ndDate < Date() Then
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode("Cannot shift to a past date.")
    Response.End
End If

' Fetch the job — user must be either the assignee or the giver, and status must be Pending
Dim rs
Set rs = objConn.Execute( _
    "SELECT j.Title, j.TargetDate, j.ShiftsUsed, jt.ShiftsAllowed AS MaxShifts " & _
    "FROM "  & DB_SCHEMA & "Jobs j " & _
    "JOIN "  & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
    "WHERE j.JobID = " & jobID & _
    "  AND j.Status = 'Pending'" & _
    "  AND (j.AssignedTo = " & curUserID & " OR j.AssignedBy = " & curUserID & ")")

If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode("Job not found or already submitted.")
    Response.End
End If

Dim jobTitle   : jobTitle   = CStr(rs("Title"))
Dim currTarget : currTarget = CDate(rs("TargetDate"))
Dim shiftsUsed : shiftsUsed = CLng(rs("ShiftsUsed"))
Dim maxShifts  : maxShifts  = CLng(rs("MaxShifts"))
rs.Close : Set rs = Nothing

' Cannot shift a job that is already overdue
If currTarget < Date() Then
    objConn.Close : Set objConn = Nothing
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode("Cannot shift an overdue job.")
    Response.End
End If

' New date must differ from current target
If ndDate = currTarget Then
    objConn.Close : Set objConn = Nothing
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode("New date must be different from the current target.")
    Response.End
End If

' Check shifts remaining — managers bypass this limit
If maxShifts <> -1 Then
    If (maxShifts - shiftsUsed) <= 0 And Not curIsManager Then
        objConn.Close : Set objConn = Nothing
        Response.Redirect "myjobs.asp?err=" & Server.URLEncode("No shifts remaining for this job.")
        Response.End
    End If
End If

' -- Call sp_ShiftJobDate -------------------------------------
' The SP updates TargetDate, increments ShiftsUsed, logs the reason
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = objConn
    .CommandType      = 4       ' adCmdStoredProc
    .CommandText      = DB_SCHEMA & "sp_ShiftJobDate"
    .Parameters.Append .CreateParameter("@JobID",          3,   1,   , jobID)    ' adInteger
    .Parameters.Append .CreateParameter("@RequestedBy",    3,   1,   , curUserID)
    .Parameters.Append .CreateParameter("@NewTargetDate",  135, 1,   , ndDate)   ' adDBDate
    .Parameters.Append .CreateParameter("@Reason",         200, 1, 500, reason)  ' adVarChar, max 500
End With

On Error Resume Next
cmd.Execute
Dim spError : spError = Err.Description
On Error GoTo 0
Set cmd = Nothing

If spError <> "" Then
    objConn.Close : Set objConn = Nothing
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode(spError)
    Response.End
End If

objConn.Close : Set objConn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="3;url=myjobs.asp">
<title>Date Shifted</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f3f4f6;display:flex;align-items:center;justify-content:center;min-height:100vh}
.box{background:#fff;border-radius:12px;padding:40px;max-width:480px;width:90%;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1)}
.alert-success{background:#d1fae5;color:#065f46;border-left:4px solid #10b981;padding:15px;border-radius:8px;text-align:left;margin:20px 0}
</style>
</head>
<body>
<div class="box">
    <div style="font-size:64px;margin-bottom:16px">&#x1F4C5;</div>
    <h2 style="font-size:22px;margin-bottom:8px">Date Shifted!</h2>
    <p style="color:#6b7280;margin-bottom:20px"><strong><%= Server.HTMLEncode(jobTitle) %></strong></p>
    <div class="alert-success">
        &#x2705; Successfully shifted!<br><br>
        Old target: <strong><%= fmtDate(currTarget) %></strong><br>
        New target: <strong><%= fmtDate(ndDate) %></strong>
        <% If reason <> "" Then %>
        <br>Reason logged: <em><%= Server.HTMLEncode(reason) %></em>
        <% End If %>
    </div>
    <p style="color:#9ca3af;font-size:13px">Redirecting to My Jobs in 3 seconds&hellip;</p>
    <a href="myjobs.asp" style="display:inline-block;margin-top:12px;color:#3b82f6">Go now &rarr;</a>
</div>
</body>
</html>
