<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<%
' ============================================================
' login_action.asp — Login POST handler
' Reads username/password from form, validates against DB,
' sets Session variables, redirects to dashboard on success.
' ============================================================

Dim username, password
username = Trim(LCase(Request.Form("username")))
password = Trim(Request.Form("password"))

' Reject empty fields before hitting the DB
If username = "" Or password = "" Then
    Response.Redirect "login.asp?err=1"
    Response.End
End If

' Look up user by Username (case-insensitive)
' ISNULL fallbacks handle cases where column_patch.sql has not yet been run
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT UserID, Username, " & _
        "ISNULL(LoginName, LOWER(Username)) AS LoginName, " & _
        "PasswordHash, IsManager, " & _
        "ISNULL(Dept, Department) AS Dept, " & _
        "IsActive " & _
        "FROM " & DB_SCHEMA & "Users " & _
        "WHERE LOWER(Username) = '" & Replace(username, "'", "''") & "'", _
        objConn, 1, 1   ' adOpenKeyset, adLockReadOnly

' Unknown username
If rs.EOF Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "login.asp?err=1"
    Response.End
End If

' Inactive account
If Not CBool(rs("IsActive")) Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "login.asp?err=2"
    Response.End
End If

' Wrong password
' NOTE: PasswordHash currently stores plain text.
'       Replace this comparison if you add hashing later.
If password <> CStr(rs("PasswordHash")) Then
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Redirect "login.asp?err=1"
    Response.End
End If

' ── All checks passed — store user info in Session ──────────
Session("UserID")    = CInt(rs("UserID"))
Session("UserName")  = CStr(rs("Username"))
Session("LoginName") = CStr(rs("LoginName"))
Session("IsManager") = CBool(rs("IsManager"))
Session("Dept")      = CStr(rs("Dept"))

rs.Close : Set rs = Nothing
objConn.Close : Set objConn = Nothing

Response.Redirect "dashboard.asp"
Response.End
%>
