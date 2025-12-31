<%@ Language="VBScript" %>
<%
Option Explicit

Dim conn, rs, sql, u, p, errMsg
errMsg = ""

u = Trim(Request.Form("username"))
p = Trim(Request.Form("password"))

If u <> "" And p <> "" Then
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB;Data Source=.;Initial Catalog=vndb;Integrated Security=SSPI;"

    sql = "SELECT u.UserID, u.UserName, r.RoleName " & _
          "FROM vn.Users u " & _
          "JOIN vn.Roles r ON u.RoleID = r.RoleID " & _
          "WHERE u.LoginName = '" & Replace(u,"'","''") & "' " & _
          "AND u.LoginPassword = '" & Replace(p,"'","''") & "' " & _
          "AND u.IsActive = 1"

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        Response.Cookies("vn_userid")   = rs("UserID")
        Response.Cookies("vn_username") = rs("UserName")
        Response.Cookies("vn_role")     = rs("RoleName")

        Response.Cookies("vn_userid").Expires = Date + 1
        Response.Redirect "dashboard.asp"
    Else
        errMsg = "Invalid username or password"
    End If
End If
%>

<!DOCTYPE html>
<html>
<head>
<title>Login</title>
<style>
body{
    font-family:Segoe UI,Arial;
    background:#f2f4f7;
    height:100vh;
    display:flex;
    align-items:center;
    justify-content:center;
}
.login-box{
    width:360px;
    background:#fff;
    padding:30px;
    border-radius:10px;
    box-shadow:0 10px 25px rgba(0,0,0,.15)
}
.login-box h2{
    text-align:center;
    margin-bottom:20px;
    color:#333
}
.input-group{margin-bottom:15px}
.input-group label{
    display:block;
    font-size:14px;
    color:#555;
    margin-bottom:5px
}
.input-group input{
    width:100%;
    padding:10px;
    border:1px solid #ccc;
    border-radius:6px;
    font-size:15px
}
.btn{
    width:100%;
    padding:12px;
    border:none;
    border-radius:6px;
    background:#0066cc;
    color:#fff;
    font-size:16px;
    cursor:pointer
}
.btn:hover{background:#0052a3}
.err{
    color:#c00;
    text-align:center;
    margin-bottom:10px
}
.footer{
    text-align:center;
    font-size:12px;
    color:#888;
    margin-top:15px
}
</style>
</head>
<body>

<div class="login-box">
    <h2>Vinova System</h2>

    <% If errMsg <> "" Then %>
        <div class="err"><%=errMsg%></div>
    <% End If %>

    <form method="post">
        <div class="input-group">
            <label>Username</label>
            <input type="text" name="username" required>
        </div>

        <div class="input-group">
            <label>Password</label>
            <input type="password" name="password" required>
        </div>

        <input type="submit" value="Login" class="btn">
    </form>

    <div class="footer">Vinova Ventures</div>
</div>

</body>
</html>
