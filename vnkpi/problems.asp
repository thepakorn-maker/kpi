<%@ Language="VBScript" CODEPAGE="874" LCID="1033" %>
<%
    ' Fixes browser display
    Response.CharSet = "windows-874" 
    
    ' Optional: Keeps dynamic/session data safe, 
    ' though your SQL collation handles the heavy lifting.
    Response.CodePage = 874 
%>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
Dim activePage : activePage = "problems"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' problems.asp — Problem Tracking & Meeting Agenda (manager only)
' Matches prototype renderProblems() exactly.
'
' Shows:
'   - "Report New Problem" button → add_problem.asp
'   - Pending items (Status='Pending') as problem-cards
'     color: problem-urgent (penalty), problem-solved (reward), plain (neutral)
'     "Discuss & Apply" button → solve_problem_action.asp
'   - Resolved items (Status='Resolved') — top 5 most recent
' ============================================================

' -- Category label helper ------------------------------------
Function categoryLabel(cat)
    Select Case LCase(cat)
        Case "quality"       : categoryLabel = "&#x26A0; Quality Issue"
        Case "deadline"      : categoryLabel = "&#x23F0; Deadline Pattern"
        Case "communication" : categoryLabel = "&#x1F4AC; Communication"
        Case "training"      : categoryLabel = "&#x1F4DA; Training Need"
        Case "recognition"   : categoryLabel = "&#x2B50; Recognition"
        Case Else            : categoryLabel = "&#x1F4CC; Other"
    End Select
End Function

' -- Count sections -------------------------------------------
Dim cntPending, cntResolved
Dim rs

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "ProblemTracking WHERE Status='Pending'")
cntPending = rs(0) : rs.Close : Set rs = Nothing

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "ProblemTracking WHERE Status='Resolved'")
cntResolved = rs(0) : rs.Close : Set rs = Nothing
%>

<%
' -- Error / success message from querystring -----------------
Dim pageMsg : pageMsg = ""
If Request.QueryString("err") <> "" Then
    pageMsg = "<div class=""alert alert-danger"" style=""margin-bottom:20px"">&#x26A0; <strong>Error:</strong> " & safeStr(Server.URLDecode(Request.QueryString("err"))) & "</div>"
End If
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F4DD; Problem Tracking &amp; Meeting Agenda</h2>
    <%= pageMsg %>

    <a href="add_problem.asp" class="btn btn-success" style="font-size:16px;padding:15px 30px">
        &#x2795; Report New Problem / Achievement
    </a>

    <div class="alert alert-info" style="margin-top:20px">
        <strong>&#x1F4A1; How it works:</strong><br>
        1. Report problems or achievements during the week<br>
        2. System organizes them by priority for meeting<br>
        3. Discuss in meeting and apply KPI adjustments<br>
        4. Track history of all issues and resolutions
    </div>

    <!-- -- Pending Items -- -->
    <h3 style="margin-top:40px;font-size:24px;margin-bottom:20px">
        &#x1F534; Pending Items for Meeting (<%= cntPending %>)
    </h3>

    <% If cntPending = 0 Then %>
    <div style="text-align:center;padding:40px;background:#f9fafb;border-radius:8px">
        <div style="font-size:60px;margin-bottom:15px">&#x2705;</div>
        <p style="font-size:18px;color:#6b7280">No pending items &mdash; all clear for this week!</p>
    </div>
    <% Else %>
    <%
    Set rs = objConn.Execute( _
        "SELECT p.ProblemID, p.Category, p.Description, p.Points, p.RaisedDate, " & _
        "       u.Username " & _
        "FROM " & DB_SCHEMA & "ProblemTracking p " & _
        "JOIN " & DB_SCHEMA & "Users u ON p.UserID = u.UserID " & _
        "WHERE p.Status = 'Pending' " & _
        "ORDER BY p.RaisedDate ASC")

    Do While Not rs.EOF
        Dim pts     : pts     = CLng(rs("Points"))
        Dim isPen   : isPen   = (pts < 0)
        Dim isRew   : isRew   = (pts > 0)
        Dim cardCls : cardCls = "problem-card"
        If isPen Then cardCls = "problem-card problem-urgent"
        If isRew Then cardCls = "problem-card problem-solved"

        Dim impactColor : impactColor = "#6b7280"
        If isPen Then impactColor = "#dc2626"
        If isRew Then impactColor = "#059669"

        Dim ptsDisp : ptsDisp = CStr(pts)
        If isRew Then ptsDisp = "+" & pts
    %>
    <div class="<%= cardCls %>">
        <div style="display:flex;justify-content:space-between;align-items:start">
            <div style="flex:1">
                <h4 style="margin-bottom:10px;font-size:18px">
                    <%= categoryLabel(rs("Category")) %>
                    &mdash; <strong style="color:#3b82f6"><%= safeStr(rs("Username")) %></strong>
                </h4>
                <p style="margin-bottom:15px;font-size:15px;line-height:1.6">
                    <%= safeStr(rs("Description")) %>
                </p>
                <div style="font-size:13px;color:#6b7280">
                    Reported: <%= sys_date(rs("RaisedDate")) %> |
                    Suggested Impact:
                    <strong style="color:<%= impactColor %>;font-size:16px">
                        <%= ptsDisp %> pts
                    </strong>
                </div>
            </div>
            <div style="margin-left:20px;display:flex;flex-direction:column;gap:8px">
                <form method="POST" action="solve_problem_action.asp" style="display:inline"
                    onsubmit="return confirm('Discuss & Apply this problem? This will update the DB.')">
                    <input type="hidden" name="ProblemID" value="<%= rs("ProblemID") %>">
                    <button type="submit" class="btn btn-success" style="font-size:14px;width:100%">
                        &#x2705; Discuss &amp; Apply
                    </button>
                </form>
                <a href="edit_problem.asp?ProblemID=<%= rs("ProblemID") %>"
                   class="btn btn-secondary"
                   style="font-size:14px;text-align:center;padding:8px 16px">
                    &#x270F;&#xFE0F; Edit
                </a>
            </div>
        </div>
    </div>
    <% rs.MoveNext : Loop %>
    <% rs.Close : Set rs = Nothing %>
    <% End If %>

    <!-- -- Resolved Items (top 5) -- -->
    <% If cntResolved > 0 Then %>
    <h3 style="margin-top:40px;font-size:24px;margin-bottom:20px">
        &#x2705; Resolved Problems (<%= cntResolved %>)
    </h3>
    <%
    Set rs = objConn.Execute( _
        "SELECT TOP 5 p.Category, p.Description, p.Points, p.RaisedDate, " & _
        "       u.Username " & _
        "FROM " & DB_SCHEMA & "ProblemTracking p " & _
        "JOIN " & DB_SCHEMA & "Users u ON p.UserID = u.UserID " & _
        "WHERE p.Status = 'Resolved' " & _
        "ORDER BY p.RaisedDate DESC")

    Do While Not rs.EOF
        Dim rPts    : rPts    = CLng(rs("Points"))
        Dim rDisp   : rDisp   = CStr(rPts) : If rPts > 0 Then rDisp = "+" & rPts
    %>
    <div class="problem-card problem-solved">
        <strong><%= safeStr(rs("Username")) %></strong> &mdash; <%= safeStr(rs("Description")) %>
        <div style="font-size:12px;color:#6b7280;margin-top:5px">
            Resolved: <%= sys_date(rs("RaisedDate")) %> | Applied: <%= rDisp %> pts
        </div>
    </div>
    <% rs.MoveNext : Loop %>
    <% rs.Close : Set rs = Nothing %>
    <% End If %>
</div>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
