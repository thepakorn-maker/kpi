<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' add_problem_action.asp — POST: insert into ProblemTracking
' Calls sp_AddProblem then redirects to problems.asp
' ============================================================

Dim userID      : userID      = 0
Dim category    : category    = Trim(Request.Form("Category"))
Dim description : description = Trim(Request.Form("Description"))
Dim points      : points      = 0

If Request.Form("UserID") <> "" Then userID = CLng(Request.Form("UserID"))
If Request.Form("Points") <> "" Then points = CLng(Request.Form("Points"))

If userID = 0 Or description = "" Then
    Response.Redirect "add_problem.asp?err=" & Server.URLEncode("Person and description are required.")
    Response.End
End If

Dim sql
sql = "EXEC " & DB_SCHEMA & "sp_AddProblem " & _
      "@UserID="       & userID                            & ", " & _
      "@Category='"    & Replace(category,    "'", "''")  & "', " & _
      "@Description='" & Replace(description, "'", "''")  & "', " & _
      "@Points="       & points                            & ", " & _
      "@ReportedBy="   & curRealUserID

On Error Resume Next
objConn.Execute sql
Dim spError : spError = Err.Description
On Error GoTo 0
objConn.Close : Set objConn = Nothing

If spError <> "" Then
    Response.Redirect "add_problem.asp?err=" & Server.URLEncode(spError)
    Response.End
End If

Response.Redirect "problems.asp"
Response.End
%>
