<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<%
' ============================================================
' switch_user.asp — Manager-only: switch the "view as" context
' POST fields:
'   TargetUserID = UserID to view as (0 = reset to self)
'   ReturnPage   = page to redirect back to after switching
'
' Sets Session ViewAs* variables that auth_check.asp reads.
' Only managers are allowed to use this.
' ============================================================

' Only managers may switch users
If Not curIsManager Then
    Response.Redirect "dashboard.asp"
    Response.End
End If

Dim targetID   : targetID   = CLng(Request.Form("TargetUserID"))
Dim returnPage : returnPage = Trim(Request.Form("ReturnPage"))

' Validate returnPage to prevent open redirect
' Allow only known pages within 
Dim allowedPages(6)
allowedPages(0) = "dashboard.asp"
allowedPages(1) = "myjobs.asp"
allowedPages(2) = "assign_job.asp"
allowedPages(3) = "approve_list.asp"
allowedPages(4) = "team.asp"
allowedPages(5) = "problems.asp"
allowedPages(6) = "meeting.asp"

Dim validReturn : validReturn = False
Dim i
For i = 0 To 6
    If returnPage = allowedPages(i) Then validReturn = True : Exit For
Next
If Not validReturn Then returnPage = "dashboard.asp"

' -- Reset to self (targetID = 0 or = real manager) ----------
If targetID = 0 Or targetID = curRealUserID Then
    Session("ViewAsUserID")   = ""
    Session("ViewAsUserName") = ""
    Session("ViewAsDept")     = ""
    Response.Redirect "" & returnPage
    Response.End
End If

' -- Switch to target user ------------------------------------
' Fetch target user details from DB — must be active
Dim rs
Set rs = objConn.Execute( _
    "SELECT UserID, Username, ISNULL(Dept, Department) AS Dept FROM " & DB_SCHEMA & "Users " & _
    "WHERE UserID = " & targetID & " AND IsActive = 1")

If rs.EOF Then
    ' User not found or inactive — reset to self silently
    rs.Close : Set rs = Nothing
    Session("ViewAsUserID")   = ""
    Session("ViewAsUserName") = ""
    Session("ViewAsDept")     = ""
    Response.Redirect "" & returnPage
    Response.End
End If

' Store ViewAs in session
Session("ViewAsUserID")   = CInt(rs("UserID"))
Session("ViewAsUserName") = CStr(rs("Username"))
Session("ViewAsDept")     = CStr(rs("Dept"))

rs.Close : Set rs = Nothing
objConn.Close : Set objConn = Nothing

Response.Redirect "" & returnPage
Response.End
%>
