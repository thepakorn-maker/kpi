<%
' ============================================================
' auth_check.asp — Session authentication guard
' Must be included AFTER conn.asp, BEFORE functions.asp
'
' Provides these variables to every page:
'   curRealUserID   — always the actual logged-in user's ID
'   curRealName     — always the actual logged-in user's name
'   curIsManager    — always reflects the REAL logged-in user
'
'   curUserID       — the "active view" user ID
'   curUserName     — the "active view" user name
'   curDept         — the "active view" user department
'
' When a manager uses the user switcher:
'   curUserID/curUserName/curDept = the selected user
'   curIsManager                  = still True (real manager)
'   curRealUserID/curRealName     = the manager themselves
'
' When no ViewAs is set, curUserID = curRealUserID (normal mode)
' ============================================================

' Redirect to login if no session
If Session("UserID") = "" Then
    Response.Redirect "login.asp"
    Response.End
End If

' -- Real logged-in user (never changes with ViewAs) ---------
Dim curRealUserID, curRealName, curIsManager
curRealUserID = CInt(Session("UserID"))
curRealName   = CStr(Session("UserName"))
curIsManager  = CBool(Session("IsManager"))

' -- ViewAs logic — only managers can switch user context ----
Dim curUserID, curUserName, curDept

If curIsManager And Session("ViewAsUserID") <> "" Then
    ' Manager is currently viewing as another user
    curUserID   = CInt(Session("ViewAsUserID"))
    curUserName = CStr(Session("ViewAsUserName"))
    curDept     = CStr(Session("ViewAsDept"))
Else
    ' Normal mode — view as self
    curUserID   = curRealUserID
    curUserName = curRealName
    curDept     = CStr(Session("Dept"))
End If
%>
