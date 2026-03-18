<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
' ============================================================
' submit_job.asp — POST action: employee submits a job
' Flow: validate ownership → call sp_SubmitJob → show result
' On success: Status becomes 'Submitted', SubmitDate = today
' ============================================================

Dim jobID : jobID = CLng(Request.Form("JobID"))

If jobID = 0 Then
    Response.Redirect "myjobs.asp"
    Response.End
End If

' Verify this job belongs to the current user and is still Pending.
' Also fetch title and penalty rate for the result message.
Dim rs
Set rs = objConn.Execute( _
    "SELECT j.Title, j.TargetDate, jt.PenaltyPerDay " & _
    "FROM "   & DB_SCHEMA & "Jobs j " & _
    "JOIN "   & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
    "WHERE j.JobID = " & jobID & _
    "  AND j.AssignedTo = " & curUserID & _
    "  AND j.Status = 'Pending'")

If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "myjobs.asp?err=" & Server.URLEncode("Job not found or already submitted.")
    Response.End
End If

Dim jobTitle   : jobTitle   = CStr(rs("Title"))
Dim targetDate : targetDate = CDate(rs("TargetDate"))
Dim penPerDay  : penPerDay  = CLng(rs("PenaltyPerDay"))
rs.Close : Set rs = Nothing

' -- Call sp_SubmitJob ----------------------------------------
' The SP sets Status = 'Submitted' and SubmitDate = GETDATE()
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = objConn
    .CommandType      = 4       ' adCmdStoredProc
    .CommandText      = DB_SCHEMA & "sp_SubmitJob"
    .Parameters.Append .CreateParameter("@JobID",       3, 1, , jobID)      ' adInteger, adParamInput
    .Parameters.Append .CreateParameter("@SubmittedBy", 3, 1, , curUserID)
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

' -- Build result message -------------------------------------
' Compare today against target to show early / on-time / late
Dim daysLate : daysLate = DateDiff("d", targetDate, Date())
Dim resultClass, resultMsg

If daysLate < 0 Then
    resultClass = "alert-success"
    resultMsg   = "&#x1F389; EARLY FINISH! " & Abs(daysLate) & " days early! Great work!"
ElseIf daysLate = 0 Then
    resultClass = "alert-success"
    resultMsg   = "&#x2705; ON-TIME! Excellent!"
Else
    Dim penaltyTotal : penaltyTotal = daysLate * penPerDay
    resultClass = "alert-warning"
    resultMsg   = "&#x26A0; LATE by " & daysLate & " days<br>" & _
                  "Penalty: -" & penaltyTotal & " pts " & _
                  "(-" & penPerDay & " pts/day &times; " & daysLate & " days)"
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="3;url=myjobs.asp">
<title>Job Submitted</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f3f4f6;display:flex;align-items:center;justify-content:center;min-height:100vh}
.box{background:#fff;border-radius:12px;padding:40px;max-width:480px;width:90%;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1)}
.alert{padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid;text-align:left}
.alert-success{background:#d1fae5;color:#065f46;border-left-color:#10b981}
.alert-warning{background:#fef3c7;color:#92400e;border-left-color:#f59e0b}
</style>
</head>
<body>
<div class="box">
    <div style="font-size:64px;margin-bottom:16px">&#x1F4E4;</div>
    <h2 style="font-size:22px;margin-bottom:8px">Job Submitted!</h2>
    <p style="color:#6b7280;margin-bottom:20px"><strong><%= Server.HTMLEncode(jobTitle) %></strong></p>
    <div class="alert <%= resultClass %>">
        <%= resultMsg %><br><br>
        Submitted: <strong><%= fmtDate(Date()) %></strong><br>
        Target was: <strong><%= fmtDate(targetDate) %></strong><br><br>
        Waiting for job giver to approve.
    </div>
    <p style="color:#9ca3af;font-size:13px">Redirecting to My Jobs in 3 seconds&hellip;</p>
    <a href="myjobs.asp" style="display:inline-block;margin-top:12px;color:#3b82f6">Go now &rarr;</a>
</div>
</body>
</html>
