<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
' ============================================================
' assign_job_action.asp — POST handler for assign_job.asp
' Flow: validate inputs → INSERT vn.Jobs → redirect to myjobs.asp
'
' POST fields:
'   AssignTo   = UserID or "SELF"
'   JobType    = TypeID string e.g. "1", "4.1", "5"
'   Urgency    = "normal" | "urgent" | "emergency"
'   Title      = job title text
'   TargetDate = YYYY-MM-DD (not sent for Type 5 — auto Monday)
'   IsPrivate  = "1" if checked (only valid when AssignTo=SELF)
'
' Special rules:
'   Type 5 (Mr. Problem): AssignedTo = Mr. Problem user, date = next Monday
'   SELF: AssignedTo = curRealUserID, IsPrivate forced to 1
' ============================================================

' -- Read and sanitise form inputs ----------------------------
Dim assignTo  : assignTo  = Trim(Request.Form("AssignTo"))
Dim jobType   : jobType   = Trim(Request.Form("JobType"))
Dim urgency   : urgency   = Trim(LCase(Request.Form("Urgency")))
Dim title     : title     = Trim(Request.Form("Title"))
Dim targetDt  : targetDt  = Trim(Request.Form("TargetDate"))
Dim isPrivate : isPrivate = (Request.Form("IsPrivate") = "1")

' -- Basic validation -----------------------------------------
If jobType = "" Or title = "" Then
    Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Job type and title are required.")
    Response.End
End If

If urgency <> "urgent" And urgency <> "emergency" Then urgency = "normal"

' -- Resolve AssignedTo UserID ---------------------------------
Dim assignedToID : assignedToID = 0

If jobType = "5" Then
    ' Type 5: always assign to Mr. Problem user
    Dim rsMrP
    Set rsMrP = objConn.Execute( _
        "SELECT UserID FROM " & DB_SCHEMA & "Users " & _
        "WHERE LOWER(Username) = 'mrproblem' AND IsActive = 1")
    If rsMrP.EOF Then
        rsMrP.Close : Set rsMrP = Nothing
        objConn.Close : Set objConn = Nothing
        Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Mr. Problem user not found. Please check the Users table.")
        Response.End
    End If
    assignedToID = CLng(rsMrP("UserID"))
    rsMrP.Close : Set rsMrP = Nothing

    ' Auto-set target to next Monday
    targetDt  = fmtDate(nextMonday())
    isPrivate = False

ElseIf UCase(assignTo) = "SELF" Then
    ' Assigning to self — always private
    assignedToID = curRealUserID
    isPrivate    = True

Else
    ' Assigning to another user
    assignedToID = CLng(assignTo)

    ' Verify target user exists and is active
    Dim rsTarget
    Set rsTarget = objConn.Execute( _
        "SELECT UserID FROM " & DB_SCHEMA & "Users " & _
        "WHERE UserID = " & assignedToID & " AND IsActive = 1")
    If rsTarget.EOF Then
        rsTarget.Close : Set rsTarget = Nothing
        objConn.Close : Set objConn = Nothing
        Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Selected user not found or inactive.")
        Response.End
    End If
    rsTarget.Close : Set rsTarget = Nothing

    ' Private only allowed for self-assign
    If isPrivate Then isPrivate = False
End If

' -- Validate target date --------------------------------------
If jobType <> "5" Then
    If targetDt = "" Then
        objConn.Close : Set objConn = Nothing
        Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Please select a target date.")
        Response.End
    End If

    Dim dtCheck
    On Error Resume Next
    dtCheck = CDate(targetDt)
    If Err.Number <> 0 Then
        On Error GoTo 0
        objConn.Close : Set objConn = Nothing
        Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Invalid target date format.")
        Response.End
    End If
    On Error GoTo 0

    If dtCheck < Date() Then
        objConn.Close : Set objConn = Nothing
        Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Target date cannot be in the past.")
        Response.End
    End If
End If

' -- INSERT into vn.Jobs --------------------------------------
' OriginalTargetDate is set to TargetDate on creation
Dim sql
sql = "INSERT INTO " & DB_SCHEMA & "Jobs " & _
      "    (Title, TypeID, AssignedTo, AssignedBy, TargetDate, OriginalTargetDate, " & _
      "     Status, IsPrivate, Urgency, ShiftsUsed, CreatedAt) " & _
      "VALUES (" & _
      "    '" & Replace(title, "'", "''") & "', " & _
      "    '" & Replace(jobType, "'", "''") & "', " & _
      "    " & assignedToID & ", " & _
      "    " & curRealUserID & ", " & _
      "    '" & targetDt & "', " & _
      "    '" & targetDt & "', " & _
      "    'Pending', " & _
      "    " & iif(isPrivate, "1", "0") & ", " & _
      "    '" & urgency & "', " & _
      "    0, " & _
      "    GETDATE())"

On Error Resume Next
objConn.Execute sql
Dim insertErr : insertErr = Err.Description
On Error GoTo 0

If insertErr <> "" Then
    objConn.Close : Set objConn = Nothing
    Response.Redirect "assign_job.asp?err=" & Server.URLEncode("Failed to create job: " & insertErr)
    Response.End
End If

objConn.Close : Set objConn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="2;url=myjobs.asp">
<title>Job Assigned</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f3f4f6;display:flex;align-items:center;justify-content:center;min-height:100vh}
.box{background:#fff;border-radius:12px;padding:40px;max-width:480px;width:90%;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1)}
.alert-success{background:#d1fae5;color:#065f46;border-left:4px solid #10b981;padding:15px;border-radius:8px;text-align:left;margin:20px 0}
</style>
</head>
<body>
<div class="box">
    <div style="font-size:64px;margin-bottom:16px">&#x2705;</div>
    <h2 style="font-size:22px;margin-bottom:8px">Job Assigned!</h2>
    <div class="alert-success">
        <strong><%= Server.HTMLEncode(title) %></strong><br><br>
        <% If jobType = "5" Then %>
        &#x1F4CB; Added to meeting agenda for Monday.<br>
        Target: <strong><%= targetDt %></strong>
        <% ElseIf UCase(assignTo) = "SELF" Then %>
        &#x1F512; Assigned to yourself as a private job.<br>
        Target: <strong><%= targetDt %></strong>
        <% Else %>
        &#x1F4E4; Job assigned successfully.<br>
        Target: <strong><%= targetDt %></strong>
        <% End If %>
    </div>
    <p style="color:#9ca3af;font-size:13px">Redirecting to My Jobs in 2 seconds&hellip;</p>
    <a href="myjobs.asp" style="display:inline-block;margin-top:12px;color:#3b82f6">Go now &rarr;</a>
</div>
</body>
</html>
