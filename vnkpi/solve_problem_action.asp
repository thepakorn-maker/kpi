<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' solve_problem_action.asp — POST: mark problem Resolved + apply KPI
' Called by "Discuss & Apply" button in problems.asp
' Calls sp_AddMeetingKPIAdjustment with ProblemID to apply
' the suggested Points and mark status = 'Resolved'
' ============================================================

Dim problemID : problemID = CLng(Request.Form("ProblemID"))

If problemID = 0 Then
    Response.Redirect "problems.asp"
    Response.End
End If

' Fetch problem details — we need UserID and Points to pass to SP
Dim rs
Set rs = objConn.Execute( _
    "SELECT UserID, Points, Description FROM " & DB_SCHEMA & "ProblemTracking " & _
    "WHERE ProblemID = " & problemID & " AND Status = 'Pending'")

If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "problems.asp?err=" & Server.URLEncode("Problem not found or already resolved.")
    Response.End
End If

Dim targetUserID  : targetUserID  = CLng(rs("UserID"))
Dim adjPoints     : adjPoints     = CLng(rs("Points"))
Dim adjReason     : adjReason     = CStr(rs("Description"))
rs.Close : Set rs = Nothing

' If points = 0, just mark resolved without KPI log entry
If adjPoints = 0 Then
    objConn.Execute "UPDATE " & DB_SCHEMA & "ProblemTracking " & _
        "SET Status='Resolved' WHERE ProblemID=" & problemID
    objConn.Close : Set objConn = Nothing
    Response.Redirect "problems.asp"
    Response.End
End If

' Apply KPI adjustment via SP (also marks problem Resolved)
Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = objConn
    .CommandType      = 4       ' adCmdStoredProc
    .CommandText      = DB_SCHEMA & "sp_AddMeetingKPIAdjustment"
    .Parameters.Append .CreateParameter("@TargetUserID", 3,   1,    , targetUserID)
    .Parameters.Append .CreateParameter("@Points",       3,   1,    , adjPoints)
    .Parameters.Append .CreateParameter("@Reason",       200, 1, 200, adjReason)
    .Parameters.Append .CreateParameter("@AdjustedBy",   3,   1,    , curRealUserID)
    .Parameters.Append .CreateParameter("@ProblemID",    3,   1,    , problemID)
End With

On Error Resume Next
cmd.Execute
Dim spError : spError = Err.Description
On Error GoTo 0
Set cmd = Nothing
objConn.Close : Set objConn = Nothing

If spError <> "" Then
    Response.Redirect "problems.asp?err=" & Server.URLEncode(spError)
    Response.End
End If

Response.Redirect "problems.asp"
Response.End
%>
