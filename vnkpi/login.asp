<%@ Language="VBScript" %>
<%
' ============================================================
' login.asp — Login form
' If already logged in, redirect to dashboard
' ============================================================
If Session("UserID") <> "" Then
    Response.Redirect "dashboard.asp"
    Response.End
End If

Dim errMsg : errMsg = ""
If Request.QueryString("err") = "1" Then errMsg = "Invalid username or password."
If Request.QueryString("err") = "2" Then errMsg = "Account is inactive. Contact administrator."
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Login — Vinova KPI System</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: Arial, sans-serif;
    background: #f3f4f6;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
}
.login-box {
    background: #fff;
    border-radius: 12px;
    padding: 40px;
    width: 100%;
    max-width: 420px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.08);
}
.logo {
    text-align: center;
    margin-bottom: 28px;
}
.logo h1 {
    font-size: 22px;
    color: #1e3a5f;
    font-weight: 700;
}
.logo p {
    font-size: 13px;
    color: #6b7280;
    margin-top: 4px;
}
.form-group {
    margin-bottom: 18px;
}
label {
    display: block;
    font-size: 14px;
    font-weight: 600;
    color: #374151;
    margin-bottom: 6px;
}
input[type="text"], input[type="password"] {
    width: 100%;
    padding: 10px 14px;
    border: 1px solid #d1d5db;
    border-radius: 8px;
    font-size: 15px;
    outline: none;
    transition: border-color 0.2s;
}
input[type="text"]:focus, input[type="password"]:focus {
    border-color: #3b82f6;
}
.btn-login {
    width: 100%;
    padding: 12px;
    background: #1e3a5f;
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    margin-top: 8px;
    transition: background 0.2s;
}
.btn-login:hover { background: #1e40af; }
.error-msg {
    background: #fee2e2;
    color: #dc2626;
    border: 1px solid #fca5a5;
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 14px;
    margin-bottom: 18px;
}
.version {
    text-align: center;
    font-size: 11px;
    color: #9ca3af;
    margin-top: 20px;
}
</style>
</head>
<body>
<div class="login-box">
    <div class="logo">
        <h1>&#x1F3ED; Vinova KPI System</h1>
        <p>Job &amp; Performance Management</p>
    </div>

    <% If errMsg <> "" Then %>
    <div class="error-msg">&#x26A0; <%= errMsg %></div>
    <% End If %>

    <form method="POST" action="login_action.asp">
        <div class="form-group">
            <label for="username">Username</label>
            <input type="text" id="username" name="username"
                   value="<%= safeStr(Request.Form("username")) %>"
                   autocomplete="username" autofocus 
                   onkeydown="if(event.keyCode==13) { 
                       event.preventDefault(); 
                       if(this.value.trim()=='') { 
                           alert('Please input user code first.'); 
                       } else { 
                           document.getElementById('password').focus(); 
                       } 
                   }"
                   >
        </div>
        <div class="form-group">
            <label for="password">Password</label>
            <input type="password" id="password" name="password"
                   autocomplete="current-password">
        </div>
        <button type="submit" class="btn-login">Sign In</button>
    </form>

    <div class="version">Vinova Production System v5.0</div>
</div>
</body>
</html>
<%
Function safeStr(s)
    If IsNull(s) Then safeStr = "" Else safeStr = Server.HTMLEncode(CStr(s))
End Function
%>
