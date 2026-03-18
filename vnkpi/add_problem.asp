<%@ Language="VBScript" %>
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
' add_problem.asp — Report Problem / Achievement form (manager only)
' Matches prototype problemModal exactly.
' Submits to add_problem_action.asp
' ============================================================

' Fetch all active non-mrproblem staff for the person dropdown
Dim rs
Set rs = objConn.Execute( _
    "SELECT UserID, Username, Department AS Dept, KPIPct " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct " & _
    "WHERE IsActive = 1 AND IsManager = 0 " & _
    "ORDER BY Username ASC")
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F4DD; Report Problem / Achievement</h2>

    <form method="POST" action="add_problem_action.asp" onsubmit="return validateAddProblem()">

        <div class="form-group">
            <label>Problem Type:</label>
            <select name="Category" style="font-size:15px">
                <option value="quality">&#x26A0; Quality Issue</option>
                <option value="deadline">&#x23F0; Deadline Pattern</option>
                <option value="communication">&#x1F4AC; Communication</option>
                <option value="training">&#x1F4DA; Training Need</option>
                <option value="recognition">&#x2B50; Recognition/Praise</option>
                <option value="other">&#x1F4CC; Other</option>
            </select>
        </div>

        <div class="form-group">
            <label>Related Person: <span style="color:#dc2626">*</span></label>
            <select name="UserID" style="font-size:15px">
                <option value="">-- Select Person --</option>
                <%
                Do While Not rs.EOF
                    Dim uPct : uPct = CDbl(rs("KPIPct"))
                %>
                <option value="<%= rs("UserID") %>">
                    <%= safeStr(rs("Username")) %> &mdash;
                    <%= FormatNumber(uPct,0) %>%
                    (<%= safeStr(rs("Dept")) %>)
                </option>
                <% rs.MoveNext : Loop %>
                <% rs.Close : Set rs = Nothing %>
            </select>
        </div>

        <div class="form-group">
            <label>Description: <span style="color:#dc2626">*</span></label>
            <textarea name="Description" rows="4"
                      placeholder="Describe the problem or achievement in detail..."
                      style="font-size:14px"></textarea>
        </div>

        <div class="form-group">
            <label>Suggested KPI Impact:</label>
            <select name="Points" style="font-size:15px">
                <option value="0">0 pts (No impact &mdash; discussion only)</option>
                <optgroup label="&#x26A0; Penalties">
                    <option value="-1">-1 pt (Minor issue)</option>
                    <option value="-2">-2 pts (Moderate issue)</option>
                    <option value="-3">-3 pts (Serious issue)</option>
                    <option value="-5">-5 pts (Critical issue)</option>
                </optgroup>
                <optgroup label="&#x2B50; Rewards">
                    <option value="1">+1 pt (Good work)</option>
                    <option value="2">+2 pts (Great work)</option>
                    <option value="3">+3 pts (Excellent work)</option>
                    <option value="5">+5 pts (Outstanding &mdash; Problem Solver!)</option>
                    <option value="10">+10 pts (Exceptional achievement)</option>
                </optgroup>
            </select>
        </div>

        <button type="submit" class="btn btn-success" style="font-size:16px;padding:12px 30px">
            &#x1F4DD; Add to Meeting Agenda
        </button>
        <a href="problems.asp" class="btn btn-secondary" style="font-size:16px;padding:12px 24px">
            Cancel
        </a>

    </form>
</div>

<script>
function validateAddProblem() {
    var userID      = document.querySelector('select[name="UserID"]').value;
    var description = document.querySelector('textarea[name="Description"]').value.trim();

    if (!userID) {
        alert('Please select a Related Person.');
        return false;
    }
    if (!description) {
        alert('Please enter a Description.');
        return false;
    }
    return confirm('Add this problem/achievement to the Meeting Agenda. Are you sure?');
}
</script>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
