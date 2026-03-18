<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
Dim activePage : activePage = "meeting"   ' no dedicated tab — use meeting highlight
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' manual_penalty.asp — Manual KPI recalculation (manager only)
' Manager selects a specific user or all users and triggers
' sp_RunManualKPI to apply any outstanding late penalties
' and snapshot the current KPI scores.
'
' Useful when:
'   - Daily scheduler did not run
'   - Manager wants to force-update before a meeting
'   - Testing the penalty engine
' ============================================================

Dim resultMsg  : resultMsg  = ""
Dim resultClass: resultClass = ""

' -- Handle POST — run the penalty SP -------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim targetID : targetID = CLng(Request.Form("TargetUserID"))

    On Error Resume Next
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = objConn
        .CommandType      = 4       ' adCmdStoredProc
        .CommandText      = DB_SCHEMA & "sp_RunManualKPI"
        .CommandTimeout   = 120
        ' @UserID = NULL means all users
        If targetID > 0 Then
            .Parameters.Append .CreateParameter("@UserID", 3, 1, , targetID)
        Else
            .Parameters.Append .CreateParameter("@UserID", 3, 1, , Null)
        End If
    End With
    cmd.Execute
    Dim spError : spError = Err.Description
    On Error GoTo 0
    Set cmd = Nothing

    If spError <> "" Then
        resultClass = "alert-danger"
        resultMsg   = "&#x274C; Error: " & Server.HTMLEncode(spError)
    Else
        resultClass = "alert-success"
        If targetID > 0 Then
            resultMsg = "&#x2705; Manual KPI recalculation completed for selected user."
        Else
            resultMsg = "&#x2705; Manual KPI recalculation completed for all staff."
        End If
    End If
End If

' -- Fetch all active staff for user selector -----------------
Dim rsStaff
Set rsStaff = objConn.Execute( _
    "SELECT v.UserID, v.Username, v.KPIPct, " & _
    "       ISNULL(u.Dept, u.Department) AS Dept " & _
    "FROM " & DB_SCHEMA & "v_UserKPIPct v " & _
    "JOIN " & DB_SCHEMA & "Users u ON v.UserID = u.UserID " & _
    "WHERE v.IsActive = 1 AND v.IsManager = 0 " & _
    "AND LOWER(v.Username) <> 'mrproblem' " & _
    "ORDER BY v.Username ASC")
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x1F527; Manual KPI Recalculation</h2>

    <div class="alert alert-info">
        <strong>&#x1F4A1; When to use:</strong><br>
        &bull; Daily scheduler did not run<br>
        &bull; Force-update KPI scores before a meeting<br>
        &bull; Apply outstanding late penalties for a specific person<br><br>
        This runs <code>sp_RunManualKPI</code> which applies any pending late
        penalties and snapshots current KPI scores.
    </div>

    <% If resultMsg <> "" Then %>
    <div class="alert <%= resultClass %>" style="font-size:16px">
        <%= resultMsg %>
    </div>
    <% End If %>

    <form method="POST" action="manual_penalty.asp">
        <div class="form-group">
            <label style="font-size:16px">Apply to:</label>
            <select name="TargetUserID" style="font-size:16px">
                <option value="0">&#x1F465; ALL STAFF (recommended)</option>
                <%
                Do While Not rsStaff.EOF
                %>
                <option value="<%= rsStaff("UserID") %>">
                    <%= safeStr(rsStaff("Username")) %> —
                    <%= FormatNumber(CDbl(rsStaff("KPIPct")),0) %>%
                    (<%= safeStr(rsStaff("Dept")) %>)
                </option>
                <% rsStaff.MoveNext : Loop %>
                <% rsStaff.Close : Set rsStaff = Nothing %>
            </select>
        </div>

        <div class="alert alert-warning">
            <strong>&#x26A0; Note:</strong> This will apply late penalties for any
            overdue Pending jobs that have not been penalised today yet.
            Running multiple times on the same day is safe — the SP prevents duplicates.
        </div>

        <button type="submit" class="btn btn-warning"
                style="font-size:16px;padding:14px 32px">
            &#x1F527; Run Manual KPI Update
        </button>
        <a href="dashboard.asp" class="btn btn-secondary"
           style="font-size:16px;padding:14px 24px">
            Cancel
        </a>
    </form>
</div>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
