<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="db/connect.asp" -->
<%
' ถ้ามี Session อยู่แล้ว  ข้ามไป dashboard
If Request.Cookies("UserID") <> "" Then
    Response.Redirect "dashboard.asp"
End If

Dim strError
strError = ""

If Request.Form("action") = "login" Then
    Dim username, password, rs, sql
    
    username = Trim(Request.Form("username"))
    password = Trim(Request.Form("password"))
    
    If username <> "" And password <> "" Then
        ' ในตัวอย่างนี้ใช้ password ธรรมดาเพื่อความง่าย (จริง ๆ ควร hash)
        sql = "SELECT UserID, Name, IsManager, KPIScore FROM Users WHERE Username = ? AND PasswordHash = ?"
        
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.Parameters.Append cmd.CreateParameter("@username", 200, 1, 50, username)
        cmd.Parameters.Append cmd.CreateParameter("@password", 200, 1, 255, password) ' ใช้ hashed password จริง
        
        Set rs = cmd.Execute
        
        If Not rs.EOF Then
            Session("UserID") = rs("UserID")
            Session("UserName") = rs("Name")
            Session("IsManager") = rs("IsManager")
            Session("KPIScore") = rs("KPIScore")
            Response.Cookies("UserID") = rs("UserID")
            Response.Cookies("UserName") = rs("Name")
            Response.Cookies("IsManager") = rs("IsManager")
            Response.Cookies("KPIScore") = rs("KPIScore")

            Response.Redirect "dashboard.asp"
        Else
            strError = "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง"
        End If
        
        rs.Close
        Set rs = Nothing
        Set cmd = Nothing
    Else
        strError = "กรุณากรอกชื่อผู้ใช้และรหัสผ่าน"
    End If
End If
%>

<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="Windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - Project Management Dashboard</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
        
        body {
            font-family: 'Sarabun', sans-serif;
            margin: 0;
            padding: 0;
            background: #f9fafb;
        }
        
        .header {
            background: #1e3a8a;
            color: white;
            padding: 1rem;
            text-align: center;
            font-size: 20px;
            font-weight: 700;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .version {
            background: #22c55e;
            color: white;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            font-size: 14px;
            display: inline-flex;
            align-items: center;
        }
        
        .version::after {
            content: ' ';
            margin-left: 4px;
        }
        
        .kpi {
            font-size: 20px;
            font-weight: 700;
        }
        
        .login-card {
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            max-width: 400px;
            margin: 2rem auto;
            padding: 2rem;
        }
        
        .login-title {
            font-size: 24px;
            font-weight: 700;
            color: #1f2937;
            margin-bottom: 0.5rem;
        }
        
        .login-subtitle {
            color: #6b7280;
            margin-bottom: 1.5rem;
        }
        
        input[type="text"], input[type="password"] {
            width: 100%;
            padding: 12px 16px;
            margin: 8px 0;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            font-size: 16px;
            box-sizing: border-box;
        }
        
        input:focus {
            outline: none;
            border-color: #1e3a8a;
            box-shadow: 0 0 0 3px rgba(30,58,138,0.1);
        }
        
        .btn-login {
            background: #22c55e;
            color: white;
            font-size: 18px;
            font-weight: 600;
            padding: 12px;
            border: none;
            border-radius: 8px;
            width: 100%;
            cursor: pointer;
            margin-top: 1rem;
        }
        
        .btn-login:hover {
            background: #16a34a;
        }
        
        .error {
            color: #dc2626;
            background: #fee2e2;
            padding: 10px;
            border-radius: 8px;
            margin-bottom: 1rem;
            font-weight: 500;
        }
        
        @media (max-width: 640px) {
            .login-card {
                margin: 1rem;
                padding: 1.5rem;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <div>Project Management Dashboard <span class="version">v4.0 FINAL</span></div>
        <div class="kpi">KPI: 100 pts</div>
    </div>
    
    <div class="login-card">
        <div class="login-title">เข้าสู่ระบบ</div>
        <div class="login-subtitle">กรุณากรอกชื่อผู้ใช้และรหัสผ่านเพื่อเข้าถึงแดชบอร์ด</div>
        
        <% If strError <> "" Then %>
            <div class="error"><%= strError %></div>
        <% End If %>
        
        <form method="post">
            <input type="hidden" name="action" value="login">
            <input type="text" name="username" placeholder="ชื่อผู้ใช้" required autofocus>
            <input type="password" name="password" placeholder="รหัสผ่าน" required>
            <button type="submit" class="btn-login">เข้าสู่ระบบ</button>
        </form>
    </div>
    
    <% Call CloseConnection() %>
</body>
</html>