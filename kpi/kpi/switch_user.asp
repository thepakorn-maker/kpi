<!-- includes/switch_user.asp -->
<!-- include file="db/connect.asp" -->
<%
' ============ ประมวลผลการสลับ user (ในไฟล์นี้เลย) ============
If Request.Form("switch_user") = "1" And Request.Form("newUserID") <> "" Then
    Dim newID : newID = CLng(Request.Form("newUserID"))
    
    Dim rsNewUser
    Set rsNewUser = conn.Execute("SELECT Name, Department, IsManager, KPIScore FROM Users WHERE UserID = " & newID)
    
    If Not rsNewUser.EOF Then
        Response.Cookies("UserID") = newID
        Response.Cookies("UserID").Expires = DateAdd("d", 30, Now())
        
        Response.Cookies("UserName") = rsNewUser("Name")
        Response.Cookies("UserName").Expires = DateAdd("d", 30, Now())
        
        Response.Cookies("Department") = rsNewUser("Department")
        Response.Cookies("Department").Expires = DateAdd("d", 30, Now())
        
        Response.Cookies("IsManager") = rsNewUser("IsManager")
        Response.Cookies("IsManager").Expires = DateAdd("d", 30, Now())
        
        Response.Cookies("KPIScore") = rsNewUser("KPIScore")
        Response.Cookies("KPIScore").Expires = DateAdd("d", 30, Now())
    End If
    rsNewUser.Close
    Set rsNewUser = Nothing
    
    ' รีเฟรชหน้าปัจจุบัน
    Response.Redirect Request.ServerVariables("SCRIPT_NAME")
End If

' ============ ดึง UserID ปัจจุบันจาก Cookie ============
Dim currentCookieID
If Request.Cookies("UserID") <> "" Then
    currentCookieID = CLng(Request.Cookies("UserID"))
Else
    currentCookieID = 0
End If
%>

<div class="switch-user">
    <form method="post" style="margin:0; display:inline;">
        <input type="hidden" name="switch_user" value="1">
        Switch User : 
        <select name="newUserID" onchange="this.form.submit()" style="border:none; background:transparent; font-weight:600; color:#1e3a8a; font-size:16px; cursor:pointer;">
            <%
            Dim rsAllUsers
            Set rsAllUsers = conn.Execute("SELECT UserID, Name, Department FROM Users ORDER BY Name")
            Do While Not rsAllUsers.EOF
                Dim displayName : displayName = rsAllUsers("Name") & " - " & rsAllUsers("Department")
                %>
                <option value="<%= rsAllUsers("UserID") %>" <%= IIf(CLng(rsAllUsers("UserID")) = currentCookieID, "selected", "") %>>
                    <%= displayName %>
                </option>
                <%
                rsAllUsers.MoveNext
            Loop
            rsAllUsers.Close
            Set rsAllUsers = Nothing
            %>
        </select>
    </form>
</div>