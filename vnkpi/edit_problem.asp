<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' edit_problem.asp — Edit existing Pending Problem / Achievement
' Manager only. Only Status='Pending' problems can be edited.
' Loads existing data into form, submits to edit_problem_action.asp
' ============================================================
Dim problemID : problemID = CLng(Request.QueryString("ProblemID"))

If problemID = 0 Then
    Response.Redirect "problems.asp"
    Response.End
End If

' -- Fetch existing problem (must be Pending) -----------------
Dim rsProblem
Set rsProblem = objConn.Execute( _
    "SELECT p.ProblemID, p.Category, p.Description, p.Points, " & _
    "       p.UserID, p.Status, u.Username " & _
    "FROM " & DB_SCHEMA & "ProblemTracking p " & _
    "JOIN " & DB_SCHEMA & "Users u ON p.UserID = u.UserID " & _
    "WHERE p.ProblemID = " & problemID)

If rsProblem.EOF Then
    rsProblem.Close : Set rsProblem = Nothing
    objConn.Close   : Set objConn   = Nothing
    Response.Redirect "problems.asp?err=" & Server.URLEncode("Problem not found.")
    Response.End
End If

' -- Block editing Resolved problems --------------------------
If CStr(rsProblem("Status")) = "Resolved" Then
    rsProblem.Close : Set rsProblem = Nothing
    objConn.Close   : Set objConn   = Nothing
    Response.Redirect "problems.asp?err=" & Server.URLEncode("Cannot edit a Resolved problem.")
    Response.End
End If

' -- Store current values for pre-fill ------------------------
Dim curCategory    : curCategory    = CStr(rsProblem("Category"))
Dim curDescription : curDescription = CStr(rsProblem("Description"))
Dim curPoints      : curPoints      = CLng(rsProblem("Points"))
curUserID      = CLng(rsProblem("UserID"))
 curUsername    = CStr(rsProblem("Username"))
rsProblem.Close : Set rsProblem = Nothing

' -- Fetch all active non-manager staff for person dropdown ---
Dim rsStaff
Set rsStaff = objConn.Execute( _
    "SELECT UserID, Username, Department AS Dept, KPIPct " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct " & _
    "WHERE IsActive = 1 AND IsManager = 0 " & _
    "ORDER BY Username ASC")

' -- Error message from querystring ---------------------------
Dim errMsg : errMsg = ""
If Request.QueryString("err") <> "" Then
    errMsg = Server.URLDecode(Request.QueryString("err"))
End If

Dim activePage : activePage = "problems"
%>
<!-- #include file="header.asp" -->

<div class="card">
    <h2 style="font-size:28px;margin-bottom:8px">&#x270F;&#xFE0F; Edit Problem / Achievement</h2>
    <p style="color:#6b7280;margin-bottom:24px;font-size:15px">
        Editing Problem #<strong><%= problemID %></strong>
        &mdash; currently assigned to <strong style="color:#3b82f6"><%= safeStr(curUsername) %></strong>
    </p>

    <% If errMsg <> "" Then %>
    <div class="alert alert-danger" style="margin-bottom:20px">
        &#x26A0; <strong>Error:</strong> <%= safeStr(errMsg) %>
    </div>
    <% End If %>

    <form method="POST" action="edit_problem_action.asp" onsubmit="return validateEditProblem()">
        <input type="hidden" name="ProblemID" value="<%= problemID %>">

        <!-- Problem Type -->
        <div class="form-group">
            <label>Problem Type:</label>
            <select name="Category" style="font-size:15px">
                <option value="quality"      <%= iif(curCategory="quality",      "selected","") %>>&#x26A0; Quality Issue</option>
                <option value="deadline"     <%= iif(curCategory="deadline",     "selected","") %>>&#x23F0; Deadline Pattern</option>
                <option value="communication"<%= iif(curCategory="communication","selected","") %>>&#x1F4AC; Communication</option>
                <option value="training"     <%= iif(curCategory="training",     "selected","") %>>&#x1F4DA; Training Need</option>
                <option value="recognition"  <%= iif(curCategory="recognition",  "selected","") %>>&#x2B50; Recognition/Praise</option>
                <option value="other"        <%= iif(curCategory="other",        "selected","") %>>&#x1F4CC; Other</option>
            </select>
        </div>

        <!-- Related Person -->
        <div class="form-group">
            <label>Related Person: <span style="color:#dc2626">*</span></label>
            <select name="UserID" style="font-size:15px">
                <option value="">-- Select Person --</option>
                <%
                Do While Not rsStaff.EOF
                     uPct     = CDbl(rsStaff("KPIPct"))
                     uID      = CLng(rsStaff("UserID"))
                     uName    = CStr(rsStaff("Username"))
                     uDept    = CStr(rsStaff("Dept"))
                     selected = iif(uID = curUserID, "selected", "")
                %>
                <option value="<%= uID %>" <%= selected %>>
                    <%= safeStr(uName) %> &mdash;
                    <%= FormatNumber(uPct,0) %>%
                    (<%= safeStr(uDept) %>)
                </option>
                <% rsStaff.MoveNext : Loop %>
                <% rsStaff.Close : Set rsStaff = Nothing %>
            </select>
        </div>

        <!-- Description -->
        <div class="form-group">
            <label>Description: <span style="color:#dc2626">*</span></label>
            <textarea name="Description" rows="4"
                      placeholder="Describe the problem or achievement in detail..."
                      style="font-size:14px"><%= safeStr(curDescription) %></textarea>
        </div>

        <!-- Suggested KPI Impact -->
        <div class="form-group">
            <label>Suggested KPI Impact:</label>
            <select name="Points" style="font-size:15px">
                <option value="0"   <%= iif(curPoints=0,   "selected","") %>>0 pts (No impact &mdash; discussion only)</option>
                <optgroup label="&#x26A0; Penalties">
                    <option value="-1"  <%= iif(curPoints=-1,  "selected","") %>>-1 pt (Minor issue)</option>
                    <option value="-2"  <%= iif(curPoints=-2,  "selected","") %>>-2 pts (Moderate issue)</option>
                    <option value="-3"  <%= iif(curPoints=-3,  "selected","") %>>-3 pts (Serious issue)</option>
                    <option value="-5"  <%= iif(curPoints=-5,  "selected","") %>>-5 pts (Critical issue)</option>
                </optgroup>
                <optgroup label="&#x2B50; Rewards">
                    <option value="1"   <%= iif(curPoints=1,   "selected","") %>>+1 pt (Good work)</option>
                    <option value="2"   <%= iif(curPoints=2,   "selected","") %>>+2 pts (Great work)</option>
                    <option value="3"   <%= iif(curPoints=3,   "selected","") %>>+3 pts (Excellent work)</option>
                    <option value="5"   <%= iif(curPoints=5,   "selected","") %>>+5 pts (Outstanding &mdash; Problem Solver!)</option>
                    <option value="10"  <%= iif(curPoints=10,  "selected","") %>>+10 pts (Exceptional achievement)</option>
                </optgroup>
            </select>
        </div>

        <!-- Buttons -->
        <div style="display:flex;gap:12px;margin-top:10px">
            <button type="submit" class="btn btn-success" style="font-size:16px;padding:12px 30px">
                &#x1F4BE; Save Changes
            </button>
            <a href="problems.asp" class="btn btn-secondary" style="font-size:16px;padding:12px 24px">
                Cancel
            </a>
        </div>

    </form>
</div>

<script>
function validateEditProblem() {
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
    return confirm('Save changes to this problem? Are you sure?');
}
</script>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
