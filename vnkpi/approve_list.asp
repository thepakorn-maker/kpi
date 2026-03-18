<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
Dim activePage : activePage = "approve"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' approve_list.asp — Approve Submitted Jobs
' Matches prototype renderJobsWIP() exactly.
'
' Shows jobs assigned BY the current user (curRealUserID)
' that are in Status = 'Submitted'.
' Two sections:
'   1. Regular jobs (IsPrivate = 0)
'   2. Private jobs  (IsPrivate = 1)
' Each row shows early/late/on-time with colour coding.
' Approve button opens modal; modal POSTs to approve_job_action.asp
' ============================================================

' -- Count both sections --------------------------------------
Dim cntRegular, cntPrivate
Set rs = objConn.Execute( _
    "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs " & _
    "WHERE AssignedBy = " & curRealUserID & " AND Status = 'Submitted' AND IsPrivate = 0")
cntRegular = rs(0) : rs.Close : Set rs = Nothing

Set rs = objConn.Execute( _
    "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs " & _
    "WHERE AssignedBy = " & curRealUserID & " AND Status = 'Submitted' AND IsPrivate = 1")
cntPrivate = rs(0) : rs.Close : Set rs = Nothing

Dim totalPending : totalPending = cntRegular + cntPrivate

' -- Helper: job type pill -------------------------------------
Function jtPill(typeID, typeName)
    Dim cls : cls = "jt-" & Replace(CStr(typeID), ".", "")
    jtPill = "<span class=""" & cls & """>" & Server.HTMLEncode(CStr(typeName)) & "</span>"
End Function

Dim rs
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">
        &#x1F504; Jobs Submitted &mdash; Awaiting My Approval (<%= totalPending %>)
    </h2>

    <div class="alert alert-info">
        <strong>&#x1F4A1; How it works:</strong>
        When someone submits a job you assigned, it appears here for you to review and approve.<br>
        Jobs can be submitted anytime &mdash; before, on, or after the target date.
        When you approve, any late penalties are automatically applied to their KPI.
    </div>

    <!-- ════════════════════════════════════════════════════
         SECTION 1: Regular Jobs
    ═════════════════════════════════════════════════════ -->
    <% If cntRegular > 0 Then %>
    <h3 style="margin:30px 0 15px">Regular Jobs (<%= cntRegular %>)</h3>
    <table>
        <tr>
            <th>ID</th>
            <th>Job Type</th>
            <th>Title</th>
            <th>Assigned To</th>
            <th>Target Date</th>
            <th>Submit Date</th>
            <th>Result</th>
            <th>Action</th>
        </tr>
        <%
        Set rs = objConn.Execute( _
            "SELECT j.JobID, j.Title, j.TypeID AS JobTypeID, jt.TypeName, " & _
            "       j.TargetDate, j.SubmitDate, jt.PenaltyPerDay, " & _
            "       at.Username AS AssignedToName " & _
            "FROM " & DB_SCHEMA & "Jobs j " & _
            "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
            "JOIN " & DB_SCHEMA & "Users at    ON j.AssignedTo = at.UserID " & _
            "WHERE j.AssignedBy = " & curRealUserID & " " & _
            "  AND j.Status = 'Submitted' " & _
            "  AND j.IsPrivate = 0 " & _
            "ORDER BY j.SubmitDate ASC")

        Do While Not rs.EOF
            Dim targetDt : targetDt = CDate(rs("TargetDate"))
            Dim submitDt : submitDt = CDate(rs("SubmitDate"))
            Dim daysDiff : daysDiff = DateDiff("d", targetDt, submitDt)
            Dim penPD    : penPD    = CLng(rs("PenaltyPerDay"))
            Dim isLate   : isLate   = (daysDiff > 0)
            Dim isEarly  : isEarly  = (daysDiff < 0)
            Dim penalty  : penalty  = 0
            If isLate And penPD < 0 Then penalty = daysDiff * penPD  ' negative value

            ' Row CSS class matches prototype exactly
            Dim rowClass : rowClass = ""
            If isEarly Then rowClass = "early-finish"
            If isLate  Then rowClass = "overdue-row"

            ' Result cell text and color
            Dim resultColor, resultHTML
            If isLate Then
                resultColor = "#dc2626"
                resultHTML  = daysDiff & " days LATE &#x26A0;<br><small>(" & penalty & " pts total)</small>"
            ElseIf isEarly Then
                resultColor = "#059669"
                resultHTML  = Abs(daysDiff) & " days EARLY! &#x1F389;"
            Else
                resultColor = "#059669"
                resultHTML  = "On-Time &#x2705;"
            End If

            ' Pass job data to modal via JS call
            Dim jID     : jID     = CLng(rs("JobID"))
            Dim jTitle  : jTitle  = Replace(CStr(rs("Title")), "'", "\'")
            Dim jToName : jToName = Replace(CStr(rs("AssignedToName")), "'", "\'")
            Dim jType   : jType   = Server.HTMLEncode(CStr(rs("TypeName")))
        %>
        <tr class="<%= rowClass %>">
            <td><strong>#<%= jID %></strong></td>
            <td><%= jtPill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= safeStr(rs("Title")) %></td>
            <td><strong><%= safeStr(rs("AssignedToName")) %></strong></td>
            <td><%= fmtDate(targetDt) %></td>
            <td><%= fmtDate(submitDt) %></td>
            <td style="color:<%= resultColor %>;font-weight:700;font-size:16px">
                <%= resultHTML %>
            </td>
            <td>
                <button class="btn btn-primary"
                    onclick="openAcceptModal(
                        <%= jID %>,
                        '<%= jTitle %>',
                        '<%= jType %>',
                        '<%= jToName %>',
                        '<%= fmtDate(targetDt) %>',
                        '<%= fmtDate(submitDt) %>',
                        <%= daysDiff %>,
                        <%= Abs(penalty) %>,
                        <%= Abs(penPD) %>)">
                    &#x2705; Approve
                </button>
 <button class="btn btn-danger" style="margin-left:5px;"
                    onclick="openRejectModal(<%= jID %>, '<%= jTitle %>')">
                    &#x274C; Reject
                </button>
            </td>
        </tr>
        <% rs.MoveNext : Loop %>
        <% rs.Close : Set rs = Nothing %>
    </table>
    <% End If %>

    <!-- ════════════════════════════════════════════════════
         SECTION 2: Private Jobs
    ═════════════════════════════════════════════════════ -->
    <% If cntPrivate > 0 Then %>
    <h3 style="margin:30px 0 15px">&#x1F512; Private Jobs (<%= cntPrivate %>)</h3>
    <table>
        <tr>
            <th>ID</th>
            <th>Job Type</th>
            <th>Title</th>
            <th>Target</th>
            <th>Submit</th>
            <th>Action</th>
        </tr>
        <%
        Set rs = objConn.Execute( _
            "SELECT j.JobID, j.Title, j.TypeID AS JobTypeID, jt.TypeName, " & _
            "       j.TargetDate, j.SubmitDate, jt.PenaltyPerDay " & _
            "FROM " & DB_SCHEMA & "Jobs j " & _
            "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
            "WHERE j.AssignedBy = " & curRealUserID & " " & _
            "  AND j.Status = 'Submitted' " & _
            "  AND j.IsPrivate = 1 " & _
            "ORDER BY j.SubmitDate ASC")

        Do While Not rs.EOF
            Dim pTargetDt : pTargetDt = CDate(rs("TargetDate"))
            Dim pSubmitDt : pSubmitDt = CDate(rs("SubmitDate"))
            Dim pDaysDiff : pDaysDiff = DateDiff("d", pTargetDt, pSubmitDt)
            Dim pPenPD    : pPenPD    = CLng(rs("PenaltyPerDay"))
            Dim pPenalty  : pPenalty  = 0
            If pDaysDiff > 0 And pPenPD < 0 Then pPenalty = pDaysDiff * pPenPD

            Dim pjID    : pjID    = CLng(rs("JobID"))
            Dim pjTitle : pjTitle = Replace(CStr(rs("Title")), "'", "\'")
            Dim pjType  : pjType  = Server.HTMLEncode(CStr(rs("TypeName")))
        %>
        <tr>
            <td><strong>#<%= pjID %></strong> <span class="private-badge">PRIVATE</span></td>
            <td><%= jtPill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= safeStr(rs("Title")) %></td>
            <td><%= fmtDate(pTargetDt) %></td>
            <td><%= fmtDate(pSubmitDt) %></td>
            <td>
                <button class="btn btn-primary"
                    onclick="openAcceptModal(
                        <%= pjID %>,
                        '<%= pjTitle %>',
                        '<%= pjType %>',
                        'yourself',
                        '<%= fmtDate(pTargetDt) %>',
                        '<%= fmtDate(pSubmitDt) %>',
                        <%= pDaysDiff %>,
                        <%= Abs(pPenalty) %>,
                        <%= Abs(pPenPD) %>)">
                    Accept
                </button>
            </td>
        </tr>
        <% rs.MoveNext : Loop %>
        <% rs.Close : Set rs = Nothing %>
    </table>
    <% End If %>

    <!-- ════════════════════════════════════════════════════
         ALL CLEAR
    ═════════════════════════════════════════════════════ -->
    <% If totalPending = 0 Then %>
    <div style="text-align:center;padding:60px">
        <div style="font-size:60px;margin-bottom:15px">&#x2705;</div>
        <p style="font-size:18px;color:#6b7280">No jobs waiting for acceptance</p>
    </div>
    <% End If %>
</div>

<!-- ════════════════════════════════════════════════════════
     APPROVE MODAL — matches prototype acceptModal exactly
═════════════════════════════════════════════════════════ -->
<div id="acceptModal" class="modal">
    <div class="modal-content">
        <span class="close-btn"
              onclick="document.getElementById('acceptModal').classList.remove('active')">&times;</span>
        <h2 style="margin-bottom:20px">&#x2705; Approve Submitted Work</h2>

        <!-- Job details + result — filled by openAcceptModal() JS -->
        <div id="acceptJobDetails"></div>

        <div class="form-group">
            <label>Approval Date (TODAY ONLY):</label>
            <input type="date" id="acceptDate" readonly style="background:#f3f4f6">
        </div>

        <div class="alert alert-warning">
            <strong>&#x26A0; Important:</strong> Approval date is locked to TODAY.
            This confirms the work is complete and applies any late penalties to their KPI.
        </div>

        <input type="hidden" id="acceptJobID">

        <button class="btn btn-success" onclick="confirmAccept()">&#x2705; Approve &amp; Complete</button>
        <button class="btn btn-secondary"
                onclick="document.getElementById('acceptModal').classList.remove('active')">Cancel</button>
    </div>
</div>

<!-- Hidden form — JS fills and submits to approve_job_action.asp -->
<form id="approveForm" method="POST" action="approve_job_action.asp" style="display:none">
    <input type="hidden" id="afJobID" name="JobID">
</form>
<div id="rejectModal" class="modal">
    <div class="modal-content">
        <span class="close-btn" onclick="document.getElementById('rejectModal').classList.remove('active')">&times;</span>
        <h2 style="color:#dc2626;margin-bottom:20px">&#x274C; Reject Job</h2>

        <div class="alert alert-danger" id="rejectJobSummary" style="font-weight:700;margin-bottom:15px;"></div>

        <form id="rejectForm" method="POST" action="reject_job_action.asp">
            <input type="hidden" id="rejectJobID" name="JobID">
            <div class="form-group">
                <label>Reason for Rejection: <span style="color:#dc2626">*</span></label>
                <textarea name="RejectComment" id="rejectComment" rows="4" style="width:100%; padding:10px; font-size:14px;" required placeholder="อธิบายเหตุผลที่ตีกลับ เพื่อให้พนักงานนำไปแก้ไข..."></textarea>
            </div>
            <button type="submit" class="btn btn-danger">&#x274C; Confirm Reject</button>
            <button type="button" class="btn btn-secondary" onclick="document.getElementById('rejectModal').classList.remove('active')">Cancel</button>
        </form>
    </div>
</div>
<script>
var todayStr = '<%= fmtDate(Date()) %>';

// -- openRejectModal() -----------------------------------------
function openRejectModal(jobID, title) {
    document.getElementById('rejectJobID').value = jobID;
    document.getElementById('rejectJobSummary').innerHTML = 'Job #' + jobID + ': ' + title;
    document.getElementById('rejectComment').value = '';
    document.getElementById('rejectModal').classList.add('active');
}

// -- openAcceptModal() -----------------------------------------
// Called by Approve button on each row.
// Populates modal with job summary and early/late/on-time result.
function openAcceptModal(jobID, title, typeName, assignedTo, targetDate, submitDate, daysDiff, penaltyTotal, penPD) {
    document.getElementById('acceptJobID').value = jobID;
    document.getElementById('acceptDate').value  = todayStr;

    // Build result message — matches prototype openAcceptModal() exactly
    var isLate  = daysDiff > 0;
    var isEarly = daysDiff < 0;
    var alertClass = (isLate) ? 'alert-warning' : 'alert-success';
    var resultColor = isLate ? '#dc2626' : '#059669';

    var resultHTML;
    if (isLate) {
        resultHTML = '&#x26A0; ' + daysDiff + ' days LATE<br>' +
                     'Total Penalty: -' + penaltyTotal + ' points<br>' +
                     '(-' + penPD + ' pts/day &times; ' + daysDiff + ' days)';
    } else if (isEarly) {
        resultHTML = '&#x1F389; ' + Math.abs(daysDiff) + ' days EARLY!<br>Excellent work!';
    } else {
        resultHTML = '&#x2705; Submitted ON-TIME!<br>Perfect!';
    }

    document.getElementById('acceptJobDetails').innerHTML =
        '<div class="alert ' + alertClass + '">' +
        '<strong>Job #' + jobID + ': ' + title + '</strong><br><br>' +
        '<strong>Job Type:</strong> ' + typeName + '<br>' +
        '<strong>Assigned To:</strong> ' + assignedTo + '<br>' +
        '<strong>Target Date:</strong> ' + targetDate + '<br>' +
        '<strong>Submit Date:</strong> ' + submitDate + '<br><br>' +
        '<div style="font-size:18px;font-weight:700;color:' + resultColor + '">' +
            resultHTML +
        '</div></div>';

    document.getElementById('acceptModal').classList.add('active');
}

// -- confirmAccept() -------------------------------------------
// Submits the hidden form to approve_job_action.asp.
function confirmAccept() {
    document.getElementById('afJobID').value = document.getElementById('acceptJobID').value;
    document.getElementById('approveForm').submit();
}
</script>

</div><!-- /container -->
</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
