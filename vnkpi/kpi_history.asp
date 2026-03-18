<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
Dim activePage : activePage = "team"   ' highlight Team tab
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' kpi_history.asp — KPI score history (manager only)
' Shows KPI score over time from vn.KPIHistory table.
' Manager can select any user via dropdown.
' Chart: bar chart using inline HTML/CSS (no external library).
' Table: date / score / period type below the chart.
' ============================================================

' -- Selected user (defaults to first staff member) -----------
Dim selectedID : selectedID = CLng(Request.QueryString("UserID"))

' -- Fetch all staff for dropdown -----------------------------
Dim rsStaff
Set rsStaff = objConn.Execute( _
    "SELECT v.UserID, v.Username, v.KPIPct, " & _
    "       ISNULL(u.Dept, u.Department) AS Dept " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct v " & _
    "JOIN " & DB_SCHEMA & "Users u ON v.UserID = u.UserID " & _
    "WHERE v.IsActive = 1 AND v.IsManager = 0 " & _
    "AND LOWER(v.Username) <> 'mrproblem' " & _
    "ORDER BY v.Username ASC")

' Default to first user if none selected
If selectedID = 0 And Not rsStaff.EOF Then
    selectedID = CLng(rsStaff("UserID"))
End If

' -- Fetch history for selected user (last 30 records) --------
Dim rsHist, histRows(), histDates(), histScores()
Dim histCount : histCount = 0

If selectedID > 0 Then
    Set rsHist = objConn.Execute( _
        "SELECT TOP 30 RecordDate, KPIScore, PeriodType " & _
        "FROM " & DB_SCHEMA & "KPIHistory " & _
        "WHERE UserID = " & selectedID & " " & _
        "ORDER BY RecordDate DESC")

    ' Load into arrays (reverse order for chart = oldest left)
    ReDim histRows(29)
    Do While Not rsHist.EOF And histCount < 30
        Set histRows(histCount) = Server.CreateObject("Scripting.Dictionary")
        histRows(histCount).Add "Date",   fmtDate(rsHist("RecordDate"))
        histRows(histCount).Add "Score",  CLng(rsHist("KPIScore"))
        histRows(histCount).Add "Period", CStr(rsHist("PeriodType"))
        histCount = histCount + 1
        rsHist.MoveNext
    Loop
    rsHist.Close : Set rsHist = Nothing
End If

' -- Find max score for scaling the bar chart -----------------
Dim maxScore : maxScore = 100
Dim i
For i = 0 To histCount - 1
    If histRows(i)("Score") > maxScore Then maxScore = histRows(i)("Score")
Next

' -- Get selected user's current name -------------------------
Dim selectedName : selectedName = ""
rsStaff.MoveFirst   ' rewind
Do While Not rsStaff.EOF
    If CLng(rsStaff("UserID")) = selectedID Then
        selectedName = CStr(rsStaff("Username"))
        Exit Do
    End If
    rsStaff.MoveNext
Loop
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F4C8; KPI Score History</h2>

    <!-- -- User selector -- -->
    <form method="GET" action="kpi_history.asp" style="margin-bottom:24px">
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
            <label style="font-size:16px;font-weight:600">Team Member:</label>
            <select name="UserID" style="font-size:15px;padding:8px 12px;border:2px solid #e5e7eb;border-radius:8px"
                    onchange="this.form.submit()">
                <%
                rsStaff.MoveFirst
                Do While Not rsStaff.EOF
                    Dim sel : sel = ""
                    If CLng(rsStaff("UserID")) = selectedID Then sel = " selected"
                %>
                <option value="<%= rsStaff("UserID") %>"<%= sel %>>
                    <%= safeStr(rsStaff("Username")) %>
                    (<%= safeStr(rsStaff("Dept")) %>)
                    — <%= FormatNumber(CDbl(rsStaff("KPIPct")),0) %>%
                </option>
                <% rsStaff.MoveNext : Loop %>
                <% rsStaff.Close : Set rsStaff = Nothing %>
            </select>
        </div>
    </form>

    <% If histCount = 0 Then %>
    <div class="alert alert-info">
        No KPI history recorded yet for <strong><%= safeStr(selectedName) %></strong>.<br>
        History is recorded each time <code>sp_RunDailyKPI</code> or
        <code>sp_RunManualKPI</code> runs.
    </div>
    <% Else %>

    <!-- -- Bar chart (oldest → newest, left to right) -- -->
    <h3 style="font-size:18px;margin-bottom:16px">
        &#x1F4CA; Score History — <%= safeStr(selectedName) %>
        (last <%= histCount %> records)
    </h3>

    <div style="overflow-x:auto;margin-bottom:30px">
        <div style="display:flex;align-items:flex-end;gap:4px;height:200px;
                    padding:0 8px 0 8px;border-left:2px solid #e5e7eb;
                    border-bottom:2px solid #e5e7eb;min-width:<%= histCount * 28 %>px">
            <%
            ' Render bars in reverse (oldest first = index histCount-1 down to 0)
            For i = histCount - 1 To 0 Step -1
                Dim sc      : sc      = histRows(i)("Score")
                Dim dt      : dt      = histRows(i)("Date")
                Dim barPct  : barPct  = CLng((sc / maxScore) * 100)
                Dim barColor
                If sc >= 95 Then barColor = "#0ea5e9"       ' Platinum
                ElseIf sc >= 85 Then barColor = "#f59e0b"   ' Gold
                ElseIf sc >= 70 Then barColor = "#6b7280"   ' Silver
                ElseIf sc >= 50 Then barColor = "#b45309"   ' Bronze
                Else barColor = "#dc2626"                   ' Needs Work
                End If
                Dim shortDate : shortDate = Mid(dt, 6, 5)  ' MM-DD
            %>
            <div style="display:flex;flex-direction:column;align-items:center;gap:4px">
                <div style="font-size:10px;color:#374151;font-weight:700"><%= sc %></div>
                <div style="width:20px;background:<%= barColor %>;
                            height:<%= barPct %>%;border-radius:3px 3px 0 0;
                            min-height:4px;cursor:default"
                     title="<%= dt %>: <%= sc %> pts">
                </div>
                <div style="font-size:9px;color:#6b7280;writing-mode:vertical-rl;
                            transform:rotate(180deg);height:36px">
                    <%= shortDate %>
                </div>
            </div>
            <% Next %>
        </div>
    </div>

    <!-- -- Legend -- -->
    <div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:24px;font-size:12px">
        <span><span style="display:inline-block;width:12px;height:12px;background:#0ea5e9;border-radius:2px"></span> Platinum (95%+)</span>
        <span><span style="display:inline-block;width:12px;height:12px;background:#f59e0b;border-radius:2px"></span> Gold (85%+)</span>
        <span><span style="display:inline-block;width:12px;height:12px;background:#6b7280;border-radius:2px"></span> Silver (70%+)</span>
        <span><span style="display:inline-block;width:12px;height:12px;background:#b45309;border-radius:2px"></span> Bronze (50%+)</span>
        <span><span style="display:inline-block;width:12px;height:12px;background:#dc2626;border-radius:2px"></span> Needs Work</span>
    </div>

    <!-- -- History table -- -->
    <table>
        <tr>
            <th>Date</th>
            <th>KPI Score</th>
            <th>Period Type</th>
            <th>Tier</th>
        </tr>
        <%
        For i = 0 To histCount - 1
            Dim hSc   : hSc   = histRows(i)("Score")
            Dim hTier : hTier = getTier(hSc)
            Dim hIcon : hIcon = getTierIcon(hSc)
        %>
        <tr>
            <td><%= histRows(i)("Date") %></td>
            <td style="font-weight:700;font-size:16px;color:#3b82f6"><%= hSc %></td>
            <td style="color:#6b7280;font-size:13px"><%= histRows(i)("Period") %></td>
            <td>
                <%= hIcon %>
                <span style="background:#fef3c7;color:#92400e;padding:3px 10px;
                             border-radius:12px;font-size:11px;font-weight:700">
                    <%= hTier %>
                </span>
            </td>
        </tr>
        <% Next %>
    </table>
    <% End If %>
</div>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
