<%@ Language=VBScript %>

<% If not(request.cookies("IsManager")) Then Response.Redirect "dashboard.asp" %>
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Meeting Mode - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
    <script>
        // ฟังก์ชัน AJAX สำหรับปรับ KPI
        function adjustKPI(userID, change) {
            var xhr = new XMLHttpRequest();
            xhr.open("POST", "process_adjust_kpi.asp", true);
            xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

            var params = "userID=" + userID + "&change=" + change;

            xhr.onreadystatechange = function () {
                if (xhr.readyState == 4 && xhr.status == 200) {
                    var response = JSON.parse(xhr.responseText);
                    if (response.success) {
                        // อัปเดต KPI Score ทันที
                        document.getElementById('kpi-' + userID).textContent = response.newScore;
                    }
                }
            };

            xhr.send(params);
        }
    </script>
</head>
<body>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <div style="padding: 1rem; text-align: center;">
        <h1 style="font-size: 28px; font-weight: 700; margin: 2rem 0; color: #1e3a8a;">
            &#127919; Meeting Mode
        </h1>
        <p style="color: #6b7280; margin-bottom: 3rem;">
            Quick KPI adjustments during team meetings (เปลี่ยนแปลงทันทีโดยไม่ reload หน้า)
        </p>

        <%
        Dim sqlUsers, rsUsers
        sqlUsers = "SELECT UserID, Name, Department, KPIScore FROM Users ORDER BY Name"
        Set rsUsers = conn.Execute(sqlUsers)

        Do While Not rsUsers.EOF
        %>
        <div class="user-card" style="max-width: 700px; margin: 2rem auto; background: #f0f9ff;">
            <div class="user-card-header">
                <div>
                    <h2><%= rsUsers("Name") %></h2>
                    <div class="position"><%= rsUsers("Department") %></div>
                </div>
                <div class="kpi-score" style="font-size: 48px;" id="kpi-<%= rsUsers("UserID") %>">
                    <%= rsUsers("KPIScore") %>
                </div>
            </div>

            <div style="display: flex; justify-content: center; gap: 3rem; margin: 2rem 0;">
                <button onclick="adjustKPI(<%= rsUsers("UserID") %>, 1)" 
                        style="background: #22c55e; color: white; padding: 20px 40px; border: none; border-radius: 12px; font-size: 24px; font-weight: 700; cursor: pointer;">
                    +1 Point<br><small>Raised/Explained Points</small>
                </button>

                <button onclick="adjustKPI(<%= rsUsers("UserID") %>, -1)" 
                        style="background: #dc2626; color: white; padding: 20px 40px; border: none; border-radius: 12px; font-size: 24px; font-weight: 700; cursor: pointer;">
                    -1 Point<br><small>No Update/Missing</small>
                </button>
            </div>
        </div>
        <%
            rsUsers.MoveNext
        Loop
        rsUsers.Close
        %>
    </div>

    <!-- Switch User -->
    <div class="switch-user">
        Switch User (Demo): 
        <strong><%= request.cookies("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% Call CloseConnection() %>
</body>
</html>