<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="functions.asp" -->
<%
' ============================================================
' daily_penalty.asp — Scheduled daily trigger
' Called by Windows Task Scheduler (not by browser directly).
' Executes sp_RunDailyKPI which:
'   1. sp_AutoApplyPenalties_v2 — applies late penalties
'   2. sp_RecordKPIHistory      — snapshots KPI scores
'
' Setup: Task Scheduler calls this URL once per day at midnight.
' Example: C:\Windows\System32\cscript.exe //nologo trigger.vbs
' Or use SQL Server Agent to call the SP directly instead.
'
' Security: restrict access to localhost only via IIS IP filter,
' or check for a secret token in querystring.
' ============================================================

' Optional secret token check — set your own token in IIS env
' Dim token : token = Request.QueryString("token")
' If token <> "YOUR_SECRET_TOKEN" Then Response.Status = "403 Forbidden" : Response.End

Response.ContentType = "text/plain"

Dim startTime : startTime = Now()

On Error Resume Next

objConn.CommandTimeout = 120        ' 2 min timeout for large datasets
objConn.Execute "EXEC " & DB_SCHEMA & "sp_RunDailyKPI"
Dim spError : spError = Err.Description
On Error GoTo 0

objConn.Close : Set objConn = Nothing

Dim endTime : endTime = Now()
Dim elapsed : elapsed = DateDiff("s", startTime, endTime)

If spError <> "" Then
    Response.Write "ERROR: " & spError & vbCrLf
    Response.Write "Time: " & Now() & vbCrLf
Else
    Response.Write "OK: sp_RunDailyKPI completed in " & elapsed & "s" & vbCrLf
    Response.Write "Time: " & Now() & vbCrLf
End If
%>
