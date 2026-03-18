<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' resolve_problem_action.asp — POST: mark Type 5 job as Completed
' Called by "Mark Resolved" button in meeting.asp
' Sets Job Status = 'Completed' without applying any KPI change
' ============================================================

Dim jobID : jobID = CLng(Request.Form("JobID"))

If jobID = 0 Then
    Response.Redirect "meeting.asp"
    Response.End
End If

' Verify it is a Type 5 Pending job
Dim rs
Set rs = objConn.Execute( _
    "SELECT JobID FROM " & DB_SCHEMA & "Jobs " & _
    "WHERE JobID = " & jobID & " AND TypeID = '5' AND Status = 'Pending'")

If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "meeting.asp?err=" & Server.URLEncode("Agenda item not found or already resolved.")
    Response.End
End If
rs.Close : Set rs = Nothing

objConn.Execute _
    "UPDATE " & DB_SCHEMA & "Jobs " & _
    "SET Status='Completed', CompletedDate=CAST(GETDATE() AS DATE) " & _
    "WHERE JobID=" & jobID

objConn.Close : Set objConn = Nothing
Response.Redirect "meeting.asp"
Response.End
%>
