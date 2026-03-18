<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
'response.write session.timeout
Dim activePage : activePage = "dashboard"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' dashboard.asp
' Employee view: matches prototype renderDashboard() for developer
' Manager view:  matches prototype renderDashboard() for manager
' ============================================================

' -- Fetch KPI data for the logged-in user -------------------
Dim rs
Set rs = objConn.Execute( _
    "SELECT KPIPct, Tier, TotalJobs, OnTimeJobs, LateJobs, Streak " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct " & _
    "WHERE UserID = " & curUserID)

Dim kpiPct, totalJobs, onTimeJobs, lateJobs, streak
kpiPct     = 100 : totalJobs = 0 : onTimeJobs = 0 : lateJobs = 0 : streak = 0

If Not rs.EOF Then
    kpiPct     = CDbl(rs("KPIPct"))
    totalJobs  = CLng(rs("TotalJobs"))
    onTimeJobs = CLng(rs("OnTimeJobs"))
    lateJobs   = CLng(rs("LateJobs"))
    streak     = CLng(rs("Streak"))
End If
rs.Close : Set rs = Nothing

' -- Count my overdue jobs ------------------------------------
Dim overdueCount
Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " AND Status='Pending' AND CAST(TargetDate AS DATE) < CAST(GETDATE() AS DATE)")
overdueCount = rs(0) : rs.Close : Set rs = Nothing

' -- Tier — recalculated from live kpiPct --------------------
Dim tierName : tierName = getTier(kpiPct)
Dim tierIcon : tierIcon = getTierIcon(kpiPct)

' -- Team average KPI — for comparison bars ------------------
Dim teamAvgPct
Set rs = objConn.Execute("SELECT AVG(KPIPct) FROM " & DB_SCHEMA & "v_UserKPIPct WHERE TotalJobs > 0")
If Not rs.EOF And Not IsNull(rs(0)) Then
    teamAvgPct = CLng(rs(0))
Else
    teamAvgPct = 74  ' fallback if no data yet
End If
rs.Close : Set rs = Nothing

' -- Top performer KPI ----------------------------------------
Dim topPct
Set rs = objConn.Execute("SELECT MAX(KPIPct) FROM " & DB_SCHEMA & "v_UserKPIPct WHERE TotalJobs > 0")
If Not rs.EOF And Not IsNull(rs(0)) Then
    topPct = CLng(rs(0))
Else
    topPct = 90
End If
rs.Close : Set rs = Nothing

' -- On-time / late percentages for stat box subtitles --------
Dim onTimePct, latePct
If totalJobs > 0 Then
    onTimePct = CLng((onTimeJobs / totalJobs) * 100)
    latePct   = CLng((lateJobs   / totalJobs) * 100)
Else
    onTimePct = 0 : latePct = 0
End If

' -- Manager extras -------------------------------------------
Dim teamOverdue, pendingProblems, topPerformerName, topPerformerPct
teamOverdue = 0 : pendingProblems = 0
topPerformerName = "" : topPerformerPct = 0

If curIsManager Then
    Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE Status='Pending' AND CAST(TargetDate AS DATE) < CAST(GETDATE() AS DATE)")
    teamOverdue = rs(0) : rs.Close : Set rs = Nothing

    Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "ProblemTracking WHERE Status='pending'")
    pendingProblems = rs(0) : rs.Close : Set rs = Nothing

    ' Top performer name and score
    Set rs = objConn.Execute( _
        "SELECT TOP 1 u.Username, v.KPIPct " & _
        "FROM " & DB_SCHEMA & "v_UserKPIPct v " & _
        "JOIN " & DB_SCHEMA & "Users u ON v.UserID = u.UserID " & _
        "WHERE v.TotalJobs > 0 ORDER BY v.KPIPct DESC")
    If Not rs.EOF Then
        topPerformerName = CStr(rs("Username"))
        topPerformerPct  = CLng(rs("KPIPct"))
    End If
    rs.Close : Set rs = Nothing
End If
%>


<% If Not curIsManager Then %>
<!-- ========================================================
     EMPLOYEE DASHBOARD — matches prototype developer view
======================================================== -->
<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F4CA; My Performance Dashboard</h2>

    <!-- Floating tier icon -->
    <div class="tier-icon"><%= tierIcon %></div>

    <!-- KPI score — prototype uses #3b82f6 always, not tier color -->
    <div style="text-align:center">
        <div style="font-size:80px;font-weight:700;color:#3b82f6;margin:20px 0">
            <%= FormatNumber(kpiPct, 0) %>%
        </div>
        <!-- Tier badge — prototype uses #fef3c7 / #92400e for ALL tiers -->
        <span style="background:#fef3c7;color:#92400e;padding:10px 20px;border-radius:25px;font-size:16px;font-weight:700">
            <%= tierName %> TIER
        </span>
    </div>

    <!-- Progress bar — height:40px matches prototype -->
    <div class="progress" style="height:40px;margin:30px 0">
        <div class="progress-bar" style="width:<%= FormatNumber(kpiPct, 0) %>%">
            <span style="font-size:16px"><%= FormatNumber(kpiPct, 0) %>% On-Time Rate</span>
        </div>
    </div>

    <!-- Stat boxes — exact prototype colors and labels -->
    <div class="stat-grid">
        <div class="stat-box">
            <div class="stat-label">Total Jobs</div>
            <div class="stat-value"><%= totalJobs %></div>
            <div class="stat-label">Completed</div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#10b981,#059669)">
            <div class="stat-label">On-Time</div>
            <div class="stat-value"><%= onTimeJobs %> &#x2705;</div>
            <div class="stat-label"><%= onTimePct %>%</div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#ef4444,#dc2626)">
            <div class="stat-label">Late</div>
            <div class="stat-value"><%= lateJobs %> &#x26A0;</div>
            <div class="stat-label"><%= latePct %>%</div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#f59e0b,#d97706)">
            <div class="stat-label">Current Streak</div>
            <div class="stat-value"><%= streak %> &#x1F525;</div>
            <div class="stat-label">Jobs on-time</div>
        </div>
    </div>

    <!-- Team Comparison section — matches prototype -->
    <h3 style="margin:30px 0 15px;font-size:20px">&#x1F4C8; Team Comparison</h3>

    <div style="margin:15px 0">
        <div style="display:flex;justify-content:space-between;margin-bottom:5px">
            <span style="font-weight:700">Your Performance</span>
            <strong style="color:#3b82f6;font-size:18px"><%= FormatNumber(kpiPct, 0) %>%</strong>
        </div>
        <div class="progress" style="height:30px">
            <div class="progress-bar" style="width:<%= FormatNumber(kpiPct, 0) %>%"></div>
        </div>
    </div>

    <div style="margin:15px 0">
        <div style="display:flex;justify-content:space-between;margin-bottom:5px">
            <span>Team Average</span>
            <strong><%= teamAvgPct %>%</strong>
        </div>
        <div class="progress" style="height:25px">
            <div class="progress-bar" style="background:#6b7280;width:<%= teamAvgPct %>%"></div>
        </div>
    </div>

    <div style="margin:15px 0">
        <div style="display:flex;justify-content:space-between;margin-bottom:5px">
            <span>Top Performer (Anonymous)</span>
            <strong><%= topPct %>%</strong>
        </div>
        <div class="progress" style="height:25px">
            <div class="progress-bar" style="background:#fbbf24;width:<%= topPct %>%"></div>
        </div>
    </div>

    <!-- Above/below team average alert -->
    <% If kpiPct >= teamAvgPct Then %>
    <div class="alert alert-success">
        <strong>&#x2705; Excellent Performance!</strong><br>
        You are <strong><%= FormatNumber(Abs(kpiPct - teamAvgPct), 0) %>%</strong> above team average.
    </div>
    <% Else %>
    <div class="alert alert-warning">
        <strong>&#x26A0; Focus Needed</strong><br>
        You are <strong><%= FormatNumber(Abs(kpiPct - teamAvgPct), 0) %>%</strong> below team average.
    </div>
    <% End If %>
</div>

<% Else %>
<!-- ========================================================
     MANAGER DASHBOARD — matches prototype manager view
======================================================== -->
<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F454; Manager Dashboard</h2>

    <!-- Stat boxes — exact prototype colors and labels -->
    <div class="stat-grid">
        <div class="stat-box">
            <div class="stat-label">Team Average</div>
            <div class="stat-value"><%= teamAvgPct %>%</div>
            <div class="stat-label">On-Time Rate</div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#10b981,#059669)">
            <div class="stat-label">Top Performer</div>
            <div class="stat-value"><%= topPerformerPct %>%</div>
            <div class="stat-label"><%= safeStr(topPerformerName) %></div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#ef4444,#dc2626)">
            <div class="stat-label">Overdue Jobs</div>
            <div class="stat-value"><%= teamOverdue %></div>
            <div class="stat-label">Needs attention</div>
        </div>
        <div class="stat-box" style="background:linear-gradient(135deg,#f59e0b,#d97706)">
            <div class="stat-label">Pending Problems</div>
            <div class="stat-value"><%= pendingProblems %></div>
            <div class="stat-label">For review</div>
        </div>
    </div>

    <!-- Quick Actions alert — matches prototype -->
    <div class="alert alert-info" style="margin-top:30px">
        <strong>&#x1F4A1; Quick Actions:</strong><br>
        &bull; Review <strong><%= pendingProblems %> pending problems</strong> in Problems tab<br>
        &bull; Check <strong><%= teamOverdue %> overdue jobs</strong> in team workload<br>
        &bull; Adjust KPIs and reward problem solvers in Meeting Mode<br>
        &bull; View full team performance rankings in Team tab
    </div>

    <!-- Recent Management Actions table — columns: Date / Action / Team Member / Points -->
    <h3 style="margin:30px 0 15px;font-size:20px">&#x1F3AF; Recent Management Actions</h3>
    <table>
        <tr><th>Date</th><th>Action</th><th>Team Member</th><th>Points</th></tr>
        <%
        Set rs = objConn.Execute( _
            "SELECT TOP 5 k.ChangedAt, u.Username, k.ChangeAmount, k.Reason " & _
            "FROM " & DB_SCHEMA & "KPILog k " & _
            "JOIN " & DB_SCHEMA & "Users u ON k.UserID = u.UserID " & _
            "WHERE k.Source = 'meeting' " & _
            "ORDER BY k.ChangedAt DESC")
        If rs.EOF Then %>
        <tr><td colspan="4" style="color:#6b7280;text-align:center">No adjustments recorded yet</td></tr>
        <% Else
           Do While Not rs.EOF %>
        <tr style="background:<%= iif(rs("ChangeAmount") > 0, "#d1fae5", "#fee2e2") %>">
            <td><%= sys_date(rs("ChangedAt")) %></td>
            <td style="font-size:13px"><%= safeStr(rs("Reason")) %></td>
            <td><strong><%= safeStr(rs("Username")) %></strong></td>
            <td style="color:<%= iif(rs("ChangeAmount") > 0, "#059669", "#dc2626") %>;font-weight:700;font-size:16px">
                <%= iif(rs("ChangeAmount") > 0, "+", "") %><%= rs("ChangeAmount") %>
            </td>
        </tr>
        <% rs.MoveNext : Loop
           End If
           rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
