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
Dim activePage : activePage = "myjobs"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' myjobs.asp — My Jobs tab
' Shows 4 sections: Target Today / Overdue / Future / Private
' Each section runs its own focused SQL query for clarity.
' ============================================================

' -- Count each section for section headers ------------------
Dim cntToday, cntOverdue, cntFuture, cntPrivate
Dim rs

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " AND Status='Pending' AND CAST(TargetDate AS DATE) = CAST(GETDATE() AS DATE)")
cntToday = rs(0) : rs.Close : Set rs = Nothing

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " AND Status='Pending' AND CAST(TargetDate AS DATE) < CAST(GETDATE() AS DATE)")
cntOverdue = rs(0) : rs.Close : Set rs = Nothing

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " AND Status='Pending' AND CAST(TargetDate AS DATE) > CAST(GETDATE() AS DATE) AND IsPrivate=0")
cntFuture = rs(0) : rs.Close : Set rs = Nothing

Set rs = objConn.Execute("SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " AND Status='Pending' AND IsPrivate=1")
cntPrivate = rs(0) : rs.Close : Set rs = Nothing

' -- Shared SQL columns for job queries ----------------------
' Separated for readability — used by all 4 section queries
Dim sqlSelect
sqlSelect = "SELECT j.JobID, j.Title, j.TypeID AS JobTypeID, j.Urgency, " & _
            "       j.TargetDate, j.OriginalTargetDate, j.ShiftsUsed, " & _
            "       jt.TypeName, jt.PenaltyPerDay, jt.ShiftsAllowed AS MaxShifts, " & _
            "       ab.Username AS AssignedByName, ab.UserID AS AssignedByID " & _
            "FROM " & DB_SCHEMA & "Jobs j " & _
            "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
            "JOIN " & DB_SCHEMA & "Users ab    ON j.AssignedBy = ab.UserID " & _
            "WHERE j.AssignedTo = " & curUserID & " AND j.Status = 'Pending' "

' -- canShiftJob() -------------------------------------------
' Returns True if the current user can shift this job.
' Mirrors prototype canShift() logic exactly.
'   typeID      = job type as string e.g. "1", "4.1", "-2"
'   assignedByID= UserID of who created the job
'   maxShifts   = -1 means unlimited (stored as -1 in DB)
'   shiftsUsed  = how many shifts already used
Function canShiftJob(typeID, assignedByID, maxShifts, shiftsUsed)
    Dim isGiver : isGiver = (CLng(assignedByID) = curUserID)

    ' Managers can always shift any job
    If curIsManager Then canShiftJob = True : Exit Function

    ' Check shifts remaining (not applicable when unlimited)
    If maxShifts <> -1 Then
        If (maxShifts - shiftsUsed) <= 0 Then
            canShiftJob = False
            Exit Function
        End If
    End If

    Select Case CStr(typeID)
        Case "-2", "-1"        : canShiftJob = isGiver          ' Owner emergency/critical: giver only
        Case "0"               : canShiftJob = True             ' Owner/Admin: anyone
        Case "1", "2", "3"    : canShiftJob = isGiver          ' Urgent/Important: giver only
        Case "4.1","4.2","4.3" : canShiftJob = True             ' Planning: receiver/giver both OK
        Case "5"               : canShiftJob = True             ' Mr. Problem: receiver/giver both OK
        Case Else              : canShiftJob = False
    End Select
End Function

' -- renderJobTypePill() --------------------------------------
' Returns the colored job type HTML pill matching prototype CSS
Function renderJobTypePill(typeID, typeName)
    ' CSS class: jt--2, jt--1, jt-0, jt-1, jt-2, jt-3, jt-41, jt-42, jt-43, jt-5
    Dim cssClass : cssClass = "jt-" & Replace(CStr(typeID), ".", "")
    renderJobTypePill = "<span class=""" & cssClass & """>" & Server.HTMLEncode(CStr(typeName)) & "</span>"
End Function

' -- renderShiftsCount() --------------------------------------
' Returns "2 / 3" or "2 / infinity" for unlimited (-1)
Function renderShiftsCount(shiftsUsed, maxShifts)
    If CLng(maxShifts) = -1 Then
        renderShiftsCount = shiftsUsed & " / &infin;"
    Else
        renderShiftsCount = shiftsUsed & " / " & maxShifts
    End If
End Function

' -- renderShiftButton() --------------------------------------
' Returns Shift button HTML or "No shifts" message
' Passes all needed values to openShiftModal() JS function
Function renderShiftButton(jobID, origTarget, currTarget, shiftsUsed, maxShifts, title, assignedByName, typeID, assignedByID, canShift, btnLabel)
    Dim maxDisp
    If CLng(maxShifts) = -1 Then maxDisp = "Unlimited" Else maxDisp = CStr(maxShifts)

    ' Escape single quotes in title/name so JS string is safe
    Dim safeTitle : safeTitle = Replace(CStr(title), "'", "\'")
    Dim safeName  : safeName  = Replace(CStr(assignedByName), "'", "\'")

    If canShift Then
        renderShiftButton = "<button class=""btn btn-warning"" " & _
            "onclick=""openShiftModal(" & jobID & ",'" & fmtDate(origTarget) & "','" & fmtDate(currTarget) & "'," & _
            shiftsUsed & ",'" & maxDisp & "','" & safeTitle & "','" & safeName & "','" & typeID & "'," & assignedByID & ")"" >" & _
            "&#x1F4C5; " & btnLabel & "</button>"
    Else
        renderShiftButton = "<span style=""color:#dc2626;font-size:11px"">&#x274C; No shifts</span>"
    End If
End Function

' ---- renderSubmitForm() -------------------------------------------
' Returns a small POST form with a Submit button
' btnClass: btn-success (normal) or btn-danger (accept late)
Function renderSubmitForm(jobID, btnClass, btnLabel)
    renderSubmitForm = "<form method=""POST"" action=""submit_job.asp"" style=""display:inline"">" & _
        "<input type=""hidden"" name=""JobID"" value=""" & jobID & """>" & _
        "<button type=""submit"" class=""btn " & btnClass & """>" & btnLabel & "</button>" & _
        "</form>"
End Function
%>


<!-- =========================================================
     SECTION 1: TARGET TODAY
     Jobs where TargetDate = today — last chance to shift
============================================================= -->
<% If cntToday > 0 Then %>
<div class="card">
    <h2 style="color:#ea580c;font-size:24px">&#x1F514; TARGET DATE TODAY &mdash; LAST CHANCE TO SHIFT! (<%= cntToday %>)</h2>
    <div class="alert alert-warning">
        <strong>&#x26A0; IMPORTANT:</strong> Today is the target date for these jobs!<br>
        &bull; You can still SHIFT these jobs today (if shifts available)<br>
        &bull; After today, shifting will be LOCKED &mdash; penalties will apply<br>
        &bull; OR submit the job as complete
    </div>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title</th>
            <th>Assigned By</th><th>Target</th><th>Shifts</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND CAST(j.TargetDate AS DATE) = CAST(GETDATE() AS DATE) ORDER BY j.TargetDate")
        Do While Not rs.EOF
            Dim canShift : canShift = canShiftJob(rs("JobTypeID"), rs("AssignedByID"), rs("MaxShifts"), rs("ShiftsUsed"))
        %>
        <tr class="target-today-row">
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong></td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td style="color:#ea580c;font-weight:700">TODAY (<%= sys_date(rs("TargetDate")) %>)</td>
            <td><%= renderShiftsCount(rs("ShiftsUsed"), rs("MaxShifts")) %></td>
            <td>
                <%= renderSubmitForm(rs("JobID"), "btn-success", "&#x2705; Submit") %>
                <%= renderShiftButton(rs("JobID"), rs("OriginalTargetDate"), rs("TargetDate"), rs("ShiftsUsed"), rs("MaxShifts"), rs("Title"), rs("AssignedByName"), rs("JobTypeID"), rs("AssignedByID"), canShift, "SHIFT NOW!") %>
            </td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>


<!-- ========================================================
     SECTION 2: OVERDUE
     Jobs where today > TargetDate — shifting is locked
======================================================== -->
<% If cntOverdue > 0 Then %>
<div class="card">
    <h2 style="color:#dc2626;font-size:24px">&#x1F6A8; OVERDUE &mdash; CANNOT SHIFT! (<%= cntOverdue %>)</h2>
    <div class="alert alert-danger">
        <strong>&#x274C; TOO LATE TO SHIFT!</strong><br>
        &bull; These jobs passed their target date<br>
        &bull; Shifting is LOCKED &mdash; cannot escape penalties<br>
        &bull; You MUST accept the job late and take the penalty<br>
        &bull; Penalties apply: -X pts/day &times; days late
    </div>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title</th><th>Assigned By</th>
            <th>Target</th><th>Days Late</th><th>Penalty</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND CAST(j.TargetDate AS DATE) < CAST(GETDATE() AS DATE) ORDER BY j.TargetDate")
        Do While Not rs.EOF
            Dim daysLate : daysLate = DateDiff("d", (rs("TargetDate")), Date())
            Dim totalPen : totalPen = daysLate * CLng(rs("PenaltyPerDay"))
        %>
        <tr class="overdue-row">
            <td><strong>#<%= rs("JobID") %></strong><%'=sys_date(rs("TargetDate")) %></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong></td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td style="color:#dc2626;font-weight:700"><%= sys_date(rs("TargetDate")) %> </td>
            <td style="color:#dc2626;font-weight:700;font-size:16px"><%= daysLate %> days &#x26A0;</td>
            <td style="color:#dc2626;font-weight:700;font-size:16px">
                -<%= totalPen %> pts<br>
                <small>(-<%= rs("PenaltyPerDay") %>/day)</small>
            </td>
            <td>
                <%= renderSubmitForm(rs("JobID"), "btn-danger", "Accept Late") %>
                <div style="color:#dc2626;font-size:10px;margin-top:5px">&#x274C; Cannot shift</div>
            </td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>


<!-- ========================================================
     SECTION 3: FUTURE JOBS
     TargetDate > today and not private — can still shift
======================================================== -->
<% If cntFuture > 0 Then %>
<div class="card">
    <h2 style="font-size:24px">&#x1F4CB; Future Jobs &mdash; Can Shift Anytime (<%= cntFuture %>)</h2>
    <div class="alert alert-success">
        <strong>&#x2705; These jobs can be managed flexibly!</strong><br>
        &bull; Target date is in the future<br>
        &bull; Can shift now or wait until target date<br>
        &bull; OR submit early/on-time for bonus
    </div>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title</th><th>Assigned By</th>
            <th>Original Target</th><th>Current Target</th><th>Shifts</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND CAST(j.TargetDate AS DATE) > CAST(GETDATE() AS DATE) AND j.IsPrivate=0 ORDER BY j.TargetDate")
        Do While Not rs.EOF
            Dim fjCanShift : fjCanShift = canShiftJob(rs("JobTypeID"), rs("AssignedByID"), rs("MaxShifts"), rs("ShiftsUsed"))
        %>
        <tr>
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td>
                <%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %>
                <% If CLng(rs("ShiftsUsed")) > 0 Then %>
                <br><span class="shift-badge">SHIFTED <%= rs("ShiftsUsed") %>x</span>
                <% End If %>
            </td>
            <td><%= getUrgencyBadge(rs("Urgency")) %><%= safeStr(rs("Title")) %></td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td><%= sys_date(rs("OriginalTargetDate")) %></td>
            <td><strong><%= sys_date(rs("TargetDate")) %></strong></td>
            <td><%= renderShiftsCount(rs("ShiftsUsed"), rs("MaxShifts")) %></td>
            <td>
                <%= renderSubmitForm(rs("JobID"), "btn-success", "&#x2705; Submit") %>
                <%= renderShiftButton(rs("JobID"), rs("OriginalTargetDate"), rs("TargetDate"), rs("ShiftsUsed"), rs("MaxShifts"), rs("Title"), rs("AssignedByName"), rs("JobTypeID"), rs("AssignedByID"), fjCanShift, "Shift") %>
            </td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>


<!-- ========================================================
     SECTION 4: PRIVATE JOBS
     Self-assigned private jobs — only visible to owner
======================================================== -->
<% If cntPrivate > 0 Then %>
<div class="card">
    <h2 style="font-size:24px">&#x1F512; My Private Jobs (<%= cntPrivate %>)</h2>
    <p style="color:#6b7280;margin-bottom:15px">Only you can see these jobs that you assigned to yourself</p>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title</th><th>Target Date</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND j.IsPrivate=1 ORDER BY j.TargetDate")
        Do While Not rs.EOF
        %>
        <tr>
            <td><strong>#<%= rs("JobID") %></strong> <span class="private-badge">PRIVATE</span></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= safeStr(rs("Title")) %></td>
            <td><%= sys_date(rs("TargetDate")) %></td>
            <td><%= renderSubmitForm(rs("JobID"), "btn-success", "&#x2705; Submit") %></td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>


<!-- ========================================================
     ALL CLEAR — shown only when all 4 sections are empty
======================================================== -->
<% If cntToday = 0 And cntOverdue = 0 And cntFuture = 0 And cntPrivate = 0 Then %>
<div class="card">
    <div style="text-align:center;padding:60px">
        <div style="font-size:80px;margin-bottom:20px">&#x1F389;</div>
        <h2 style="font-size:28px;color:#10b981;margin-bottom:10px">All Clear!</h2>
        <p style="font-size:18px;color:#6b7280">You have no pending jobs. Great work!</p>
    </div>
</div>
<% End If %>


<!-- ========================================================
     JOB TYPES REFERENCE TABLE
     Shown at the bottom when there are any active jobs
======================================================== -->
<% If cntToday > 0 Or cntOverdue > 0 Or cntFuture > 0 Then %>
<div class="card">
    <h3 style="font-size:20px;margin-bottom:15px">&#x1F4CB; Vinova Job Types &amp; Shift Rules</h3>
    <table>
        <tr><th>Job Type</th><th>Max Shifts</th><th>Penalty/Day</th></tr>
        <tr><td><span class="jt--2">Type -2 (Owner Emergency)</span></td> <td>0</td>          <td>-20 pts/day</td></tr>
        <tr><td><span class="jt--1">Type -1 (Owner Critical)</span></td>  <td>0</td>          <td>-10 pts/day</td></tr>
        <tr><td><span class="jt-0">Type 0 (Owner/Admin)</span></td>       <td>Unlimited</td>  <td>0 pts</td></tr>
        <tr><td><span class="jt-1">Type 1 (Urgent/No Delay)</span></td>   <td>0</td>          <td>-5 pts/day</td></tr>
        <tr><td><span class="jt-2">Type 2 (Urgent)</span></td>            <td>2</td>          <td>-2 pts/day</td></tr>
        <tr><td><span class="jt-3">Type 3 (Important)</span></td>         <td>3</td>          <td>-1 pt/day</td></tr>
        <tr><td><span class="jt-41">Type 4.1 (Plan-Gen)</span></td>       <td>Unlimited</td>  <td>-1 pt/day</td></tr>
        <tr><td><span class="jt-42">Type 4.2 (Plan-Soln)</span></td>      <td>3</td>          <td>-1 pt/day</td></tr>
        <tr><td><span class="jt-43">Type 4.3 (Plan-Fashion)</span></td>   <td>2</td>          <td>-1 pt/day</td></tr>
        <tr><td><span class="jt-5">Type 5 (Mr. Problem)</span></td>       <td>Unlimited</td>  <td>0 pts</td></tr>
    </table>
    <div class="alert alert-info" style="margin-top:15px">
        <strong>&#x1F4A1; Universal Rules:</strong><br>
        &bull; Can finish job ANYTIME (before, on, or after target)<br>
        &bull; Can shift date BEFORE or ON target date only<br>
        &bull; CANNOT shift AFTER target date (penalties apply)<br>
        &bull; Penalty applies PER DAY late
    </div>
</div>
<% End If %>


<!-- ========================================================
     SHIFT DATE MODAL — matches prototype modal exactly
======================================================== -->
<div id="shiftModal" class="modal">
    <div class="modal-content">
        <span class="close-btn" onclick="document.getElementById('shiftModal').classList.remove('active')">&times;</span>
        <!-- h2 matches prototype: " Shift Target Date" -->
        <h2 style="margin-bottom:20px">&#x1F4C5; Shift Target Date</h2>

        <!-- Job summary info box — filled by openShiftModal() JS -->
        <div id="shiftJobDetails"></div>

        <div class="form-group">
            <label>Original Target Date:</label>
            <input type="text" id="origTargetDisp" readonly style="background:#f3f4f6">
        </div>
        <div class="form-group">
            <label>Current Target Date:</label>
            <input type="text" id="currTargetDisp" readonly style="background:#f3f4f6">
        </div>
        <!-- Label matches prototype: "Shifts Remaining:" -->
        <div class="form-group">
            <label>Shifts Remaining:</label>
            <input type="text" id="shiftsDisp" readonly style="background:#f3f4f6">
        </div>
        <!-- Label matches prototype exactly -->
        <div class="form-group">
            <label>New Target Date (must be &ge; today): <span style="color:#dc2626">*</span></label>
            <input type="date" id="newTargetDate">
        </div>

        <!-- Reason field — shown/hidden by JS based on job type -->
        <div class="form-group" id="shiftReasonSection" style="display:none">
            <label id="shiftReasonLabel">&#x1F4AC; Reason for Shift: <span style="color:#dc2626">*</span></label>
            <textarea id="shiftReason" rows="3" placeholder="Explain why you need to shift the date..." style="font-size:14px;padding:10px"></textarea>
        </div>

        <!-- Static rules alert — always visible, matches prototype -->
        <div class="alert alert-warning">
            <strong>&#x1F4CB; Date Change Rules:</strong><br>
            &bull; <strong>Types 1, 2, 3:</strong> Only JOB GIVER can change date (Type 1 needs reason)<br>
            &bull; <strong>Types 4.x &amp; 5:</strong> JOB RECEIVER must write reason to shift<br>
            &bull; <strong>Managers:</strong> Can override in special cases
        </div>

        <!-- Hidden values passed from button click -->
        <input type="hidden" id="shiftJobID">
        <input type="hidden" id="shiftCurrTarget">

        <!-- Button label matches prototype: " Shift Date" -->
        <button class="btn btn-warning" onclick="confirmShift()">&#x1F4C5; Shift Date</button>
        <button class="btn btn-secondary" onclick="document.getElementById('shiftModal').classList.remove('active')">Cancel</button>
    </div>
</div>

<!-- Hidden form — JS fills and submits this to shift_job.asp -->
<form id="shiftForm" method="POST" action="shift_job.asp" style="display:none">
    <input type="hidden" id="sfJobID"   name="JobID">
    <input type="hidden" id="sfNewDate" name="NewDate">
    <input type="hidden" id="sfReason"  name="Reason">
</form>

<script>
var todayStr  = '<%= fmtDate(Date()) %>';
var curUserID = <%= curUserID %>;

// -- openShiftModal() --------------------------------------
// Called by the Shift button on each job row.
// Populates the modal with job details and applies shift rules.
function openShiftModal(jobID, origTarget, currTarget, shiftsUsed, maxShifts, title, assignedByName, jobTypeID, assignedByID) {

    // Fill read-only display fields
    document.getElementById('shiftJobID').value      = jobID;
    document.getElementById('shiftCurrTarget').value = currTarget;
    document.getElementById('origTargetDisp').value  = origTarget;
    document.getElementById('currTargetDisp').value  = currTarget;
    document.getElementById('shiftsDisp').value      = shiftsUsed + ' used / ' + maxShifts + ' max';
    document.getElementById('newTargetDate').min     = todayStr;
    document.getElementById('newTargetDate').value   = '';
    document.getElementById('shiftReason').value     = '';

    // Determine if a written reason is required:
    // - Type 1: required when the current user is the job giver
    // - Types 4.1, 4.2, 4.3, 5: required when the current user is NOT the giver (i.e. is receiver)
    var isGiver     = (curUserID === parseInt(assignedByID));
    var needsReason = false;
    var reasonLabel = '&#x1F4AC; Reason for Shift (Required): <span style="color:#dc2626">*</span>';

    if (jobTypeID === '1' && isGiver) {
        needsReason = true;
        reasonLabel = '&#x26A0; Reason for Shift — Required for Type 1: <span style="color:#dc2626">*</span>';
    } else if (jobTypeID === '4.1' || jobTypeID === '4.2' || jobTypeID === '4.3' || jobTypeID === '5') {
        if (!isGiver) {
            needsReason = true;
        }
    }

    document.getElementById('shiftReasonSection').style.display = needsReason ? 'block' : 'none';
    document.getElementById('shiftReasonLabel').innerHTML = reasonLabel;

    // Build rules message for the info box
    var rulesMsg = '';
    if (jobTypeID === '1' || jobTypeID === '2' || jobTypeID === '3') {
        rulesMsg = '&bull; Types 1, 2, 3: Only the <strong>Job Giver</strong> can change the date';
        if (jobTypeID === '1') rulesMsg += '<br>&bull; Type 1 requires a written reason';
    } else if (jobTypeID === '4.1' || jobTypeID === '4.2' || jobTypeID === '4.3' || jobTypeID === '5') {
        rulesMsg = '&bull; Planning &amp; Problem jobs: <strong>Job Receiver</strong> must write a reason for shifting';
    } else {
        rulesMsg = '&bull; This job type: anyone can shift freely without a reason';
    }

    // Render job summary + rules in the modal info box
    document.getElementById('shiftJobDetails').innerHTML =
        '<div class="alert alert-info">' +
            '<strong>Job #' + jobID + ': ' + title + '</strong><br>' +
            'Assigned By: ' + assignedByName + '<br>' +
            'Original Target: ' + origTarget + ' &rarr; Current Target: ' + currTarget +
        '</div>' +
        '<div class="alert alert-warning">' +
            '<strong>&#x1F4CB; Date Change Rules:</strong><br>' + rulesMsg +
        '</div>';

    document.getElementById('shiftModal').classList.add('active');
}

// -- confirmShift() ----------------------------------------
// Validates modal inputs then submits the hidden form to shift_job.asp
function confirmShift() {
    var newDate     = document.getElementById('newTargetDate').value;
    var reason      = document.getElementById('shiftReason').value.trim();
    var currTarget  = document.getElementById('shiftCurrTarget').value;
    var needsReason = (document.getElementById('shiftReasonSection').style.display !== 'none');

    if (!newDate) {
        alert('Please select a new target date!');
        return;
    }
    if (newDate < todayStr) {
        alert('Cannot shift to a past date!\n\nToday: ' + todayStr + '\nYou selected: ' + newDate);
        return;
    }
    if (newDate === currTarget) {
        alert('New date must be different from the current target date!');
        return;
    }
    if (needsReason && !reason) {
        alert('A reason is required for shifting this job. Please fill in the Reason field.');
        return;
    }

    // Pass values to hidden form and submit
    document.getElementById('sfJobID').value   = document.getElementById('shiftJobID').value;
    document.getElementById('sfNewDate').value = newDate;
    document.getElementById('sfReason').value  = reason;
    document.getElementById('shiftForm').submit();
}
</script>

</div><!-- /container -->


</body>
</html>
<% objConn.Close : Set objConn = Nothing %>
