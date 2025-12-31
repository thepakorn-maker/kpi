<!--#include file="../db/connect.asp" -->
<%
Dim currentUserName, currentDepartment, currentKPI
If request.cookies("UserID") <> "" Then
    Dim rsUser
    Set rsUser = conn.Execute("SELECT Name, Department, KPIScore FROM Users WHERE UserID = " & request.cookies("UserID"))
    If Not rsUser.EOF Then
        currentUserName = rsUser("Name")
        currentDepartment = rsUser("Department")
        currentKPI = rsUser("KPIScore")
    End If
    rsUser.Close
Else
    Response.Redirect "login.asp"
End If
%>

<div class="header">
    <div>Project Management Dashboard <span class="version">v4.0 FINAL</span></div>
    <div class="kpi">KPI: <%= currentKPI %> pts</div>
</div>