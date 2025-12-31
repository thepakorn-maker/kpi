<%@ Language=VBScript %>

<% If not(Session("IsManager")) Then Response.Redirect "dashboard.asp" %>
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Meeting Mode - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
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
            Quick KPI adjustments during team meetings
        </p>

        <%
        Dim sqlUsers, rsUsers
        sqlUsers = "SELECT UserID, Name, Department, KPIScore FROM Users ORDER BY Name"
        Set rsUsers = conn.Execute(sqlUsers)

        Do While Not rsUsers.EOF
        %>
        <div class="user-card" style="max-width: 600px; margin: 2rem auto; background: #f0f9ff;">
            <div class="user-card-header">
                <div>
                    <h2><%= rsUsers("Name") %></h2>
                    <div class="position"><%= rsUsers("Department") %></div>
                </div>
                <div class="kpi-score" style="font-size: 48px;"><%= rsUsers("KPIScore") %></div>
            </div>

            <div style="display: flex; justify-content: center; gap: 2rem; margin: 2rem 0;">
                <form method="post" style="display: inline;">
                    <input type="hidden" name="action" value="plus">
                    <input type="hidden" name="userID" value="<%= rsUsers("UserID") %>">
                    <button type="submit" style="background: #22c55e; color: white; padding: 20px 40px; border: none; border-radius: 12px; font-size: 24px; font-weight: 700; cursor: pointer;">
                        +1 Point<br><small>Raised/Explained Points</small>
                    </button>
                </form>

                <form method="post" style="display: inline;">
                    <input type="hidden" name="action" value="minus">
                    <input type="hidden" name="userID" value="<%= rsUsers("UserID") %>">
                    <button type="submit" style="background: #dc2626; color: white; padding: 20px 40px; border: none; border-radius: 12px; font-size: 24px; font-weight: 700; cursor: pointer;">
                        -1 Point<br><small>No Update/Missing</small>
                    </button>
                </form>
            </div>
        </div>
        <%
            rsUsers.MoveNext
        Loop
        rsUsers.Close
        %>

        <%
        ' Process +1 / -1
        If Request.Form("action") <> "" Then
            Dim adjustUserID : adjustUserID = CLng(Request.Form("userID"))
            Dim changeAmt : changeAmt = IIf(Request.Form("action") = "plus", 1, -1)
            Dim reason : reason = IIf(changeAmt = 1, "Meeting +1", "Meeting -1")

            conn.Execute "UPDATE Users SET KPIScore = KPIScore + " & changeAmt & " WHERE UserID = " & adjustUserID
            conn.Execute "INSERT INTO KPILog (UserID, ChangeAmount, Reason) VALUES (" & adjustUserID & ", " & changeAmt & ", '" & reason & "')"

            Response.Redirect "meetingmode.asp"
        End If
        %>
    </div>

    <!-- Switch User -->
    <div class="switch-user">
        Switch User (Demo): 
        <strong><%= Session("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% Call CloseConnection() %>
</body>
</html>