<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
' Guard — managers only
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
Dim activePage : activePage = "team"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' team.asp — Team Performance Leaderboard (manager only)
' Matches prototype renderTeam() exactly.
' Shows all non-manager staff sorted by KPIPct descending.
' Top 3 rows highlighted amber. Medal icons for rank 1-3.
' Bottom section: Top Performers (85%+) and Needs Support (<70%)
' ============================================================

Dim rs
Set rs = objConn.Execute( _
    "SELECT v.UserID, v.Username, v.KPIPct, v.TotalJobs, " & _
    "       v.OnTimeJobs, v.LateJobs, v.Streak, v.Tier, " & _
    "       ISNULL(u.Dept, u.Department) AS Dept " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct v " & _
    "JOIN " & DB_SCHEMA & "Users u ON v.UserID = u.UserID " & _
    "WHERE v.IsManager = 0 AND v.IsActive = 1 " & _
    "ORDER BY v.KPIPct DESC, v.Username ASC")
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F465; Team Performance Leaderboard</h2>
    <p style="color:#6b7280;margin-bottom:20px">
        &#x1F441; Managers Only &mdash; Full team rankings and performance metrics
    </p>

    <table>
        <tr>
            <th style="width:80px">Rank</th>
            <th>Name</th>
            <th>Department</th>
            <th style="width:120px">On-Time %</th>
            <th>Jobs Completed</th>
            <th>On-Time / Late</th>
            <th>Streak</th>
            <th>Performance Tier</th>
        </tr>
        <%
        Dim rank : rank = 0
        Dim topPerformers, needsSupport
        topPerformers = "" : needsSupport = ""

        Do While Not rs.EOF
            rank = rank + 1
             uPct   = CDbl(rs("KPIPct"))
             uName  = CStr(rs("Username"))
             uDept  = CStr(rs("Dept"))
             uJobs  = CLng(rs("TotalJobs"))
             uOT    = CLng(rs("OnTimeJobs"))
             uLate  = CLng(rs("LateJobs"))
             uStrk  = CLng(rs("Streak"))
             uTier  = getTier(uPct)
             uIcon  = getTierIcon(uPct)

            ' Medal for top 3
            Dim medal : medal = ""
            If rank = 1 Then medal = "&#x1F947;"
            If rank = 2 Then medal = "&#x1F948;"
            If rank = 3 Then medal = "&#x1F949;"

            ' Amber highlight for top 3 — matches prototype
            Dim rowStyle : rowStyle = ""
            If rank <= 3 Then rowStyle = "background:#fef3c7"
        %>
        <tr style="<%= rowStyle %>">
            <td><strong style="font-size:20px"><%= rank %>.</strong> <%= medal %></td>
            <td><strong style="font-size:16px"><%= safeStr(uName) %></strong></td>
            <td><%= safeStr(uDept) %></td>
            <td>
                <div style="font-size:24px;font-weight:700;color:#3b82f6">
                    <%= FormatNumber(uPct, 0) %>%
                </div>
            </td>
            <td><strong><%= uJobs %></strong> jobs</td>
            <td>
                <span style="color:#059669"><%= uOT %> &#x2705;</span> /
                <span style="color:#dc2626"><%= uLate %> &#x26A0;</span>
            </td>
            <td><strong style="font-size:18px"><%= uStrk %></strong> &#x1F525;</td>
            <td>
                <%= uIcon %>
                <span style="background:#fef3c7;color:#92400e;padding:6px 12px;border-radius:20px;font-weight:700;font-size:11px">
                    <%= uTier %>
                </span>
            </td>
        </tr>
        <%
            ' Build lists for bottom section
            If uPct >= 85 Then
                topPerformers = topPerformers & "<li><strong>" & Server.HTMLEncode(uName) & "</strong> - " & FormatNumber(uPct,0) & "%</li>"
            End If
            If uPct < 70 Then
                needsSupport = needsSupport & "<li><strong>" & Server.HTMLEncode(uName) & "</strong> - " & FormatNumber(uPct,0) & "%</li>"
            End If

            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing
        %>
    </table>

    <!-- Top Performers / Needs Support — matches prototype grid -->
    <div style="margin-top:40px;display:grid;grid-template-columns:1fr 1fr;gap:20px">
        <div class="alert alert-success">
            <h3 style="margin-bottom:10px">&#x1F31F; Top Performers (85%+)</h3>
            <ul style="margin-left:20px">
                <% If topPerformers = "" Then %>
                <li style="color:#065f46">No staff at 85%+ yet</li>
                <% Else %>
                <%= topPerformers %>
                <% End If %>
            </ul>
        </div>
        <div class="alert alert-warning">
            <h3 style="margin-bottom:10px">&#x26A0; Needs Support (&lt;70%)</h3>
            <ul style="margin-left:20px">
                <% If needsSupport = "" Then %>
                <li style="color:#92400e">No staff below 70% &mdash; great!</li>
                <% Else %>
                <%= needsSupport %>
                <% End If %>
            </ul>
        </div>
    </div>
</div>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
