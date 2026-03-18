<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' edit_problem_action.asp — POST: UPDATE existing Pending problem
' Called by edit_problem.asp form submission.
' Only Status='Pending' problems can be edited.
' Does NOT use a stored procedure — direct UPDATE statement.
' ============================================================

Dim problemID   : problemID   = CLng(Request.Form("ProblemID"))
Dim userID      : userID      = 0
Dim category    : category    = Trim(Request.Form("Category"))
Dim description : description = Trim(Request.Form("Description"))
Dim points      : points      = 0

If Request.Form("UserID") <> "" Then userID = CLng(Request.Form("UserID"))
If Request.Form("Points") <> "" Then points = CLng(Request.Form("Points"))

' -- Basic validation -----------------------------------------
If problemID = 0 Then
    Response.Redirect "problems.asp"
    Response.End
End If

If userID = 0 Or description = "" Then
    Response.Redirect "edit_problem.asp?ProblemID=" & problemID & _
        "&err=" & Server.URLEncode("Person and description are required.")
    Response.End
End If

' -- Verify problem exists and is still Pending ---------------
' (prevent editing a Resolved problem via direct POST)
Dim rs
Set rs = objConn.Execute( _
    "SELECT ProblemID, Status FROM " & DB_SCHEMA & "ProblemTracking " & _
    "WHERE ProblemID = " & problemID)

If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "problems.asp?err=" & Server.URLEncode("Problem not found.")
    Response.End
End If

If CStr(rs("Status")) <> "Pending" Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "problems.asp?err=" & Server.URLEncode("Cannot edit a Resolved problem.")
    Response.End
End If

rs.Close : Set rs = Nothing

' -- Validate points is an allowed value ----------------------
Dim allowedPoints(9)
allowedPoints(0) = 0
allowedPoints(1) = -1 : allowedPoints(2) = -2
allowedPoints(3) = -3 : allowedPoints(4) = -5
allowedPoints(5) = 1  : allowedPoints(6) = 2
allowedPoints(7) = 3  : allowedPoints(8) = 5
allowedPoints(9) = 10

Dim validPts : validPts = False
Dim i
For i = 0 To 9
    If points = allowedPoints(i) Then validPts = True : Exit For
Next
If Not validPts Then
    Response.Redirect "edit_problem.asp?ProblemID=" & problemID & _
        "&err=" & Server.URLEncode("Invalid points value.")
    Response.End
End If

' -- Validate category is an allowed value --------------------
Dim allowedCats(5)
allowedCats(0) = "quality"       : allowedCats(1) = "deadline"
allowedCats(2) = "communication" : allowedCats(3) = "training"
allowedCats(4) = "recognition"   : allowedCats(5) = "other"

Dim validCat : validCat = False
For i = 0 To 5
    If LCase(category) = allowedCats(i) Then validCat = True : Exit For
Next
If Not validCat Then
    Response.Redirect "edit_problem.asp?ProblemID=" & problemID & _
        "&err=" & Server.URLEncode("Invalid category.")
    Response.End
End If

' -- UPDATE ProblemTracking -----------------------------------
Dim sql
sql = "UPDATE " & DB_SCHEMA & "ProblemTracking " & _
      "SET UserID      = "  & userID                              & ", " & _
      "    Category    = '" & Replace(category,    "'", "''")     & "', " & _
      "    Description = '" & Replace(description, "'", "''")     & "', " & _
      "    Points      = "  & points                              & " " & _
      "WHERE ProblemID = "  & problemID & " AND Status = 'Pending'"

On Error Resume Next
objConn.Execute sql
Dim spError : spError = Err.Description
On Error GoTo 0

objConn.Close : Set objConn = Nothing

If spError <> "" Then
    Response.Redirect "edit_problem.asp?ProblemID=" & problemID & _
        "&err=" & Server.URLEncode(spError)
    Response.End
End If

Response.Redirect "problems.asp"
Response.End
%>
