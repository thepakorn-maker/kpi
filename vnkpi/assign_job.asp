<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
Dim activePage : activePage = "assignjob"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' assign_job.asp — Assign New Job form
' Matches prototype renderAssignJob() exactly.
'
' Rules from prototype:
'   - Cannot assign to yourself via main list (use SELF option)
'   - Cannot assign to Mr. Problem (TypeID=5 handles that automatically)
'   - Owners (IsManager=1 with Dept=BOSS) can only receive Type 0 jobs
'   - Type 5 hides assign-to, date, private fields (auto Monday)
'   - Owner-type jobs (-2,-1,0) only shown when assigner is owner
'   - Workload preview loads via AJAX when assignee changes
' ============================================================

' -- Determine if current user is an "owner" (boss-level manager) --
' Owners can assign type -2, -1, 0 and can assign to other owners
Dim isOwner : isOwner = False
If curIsManager Then
    Dim rsDept
    Set rsDept = objConn.Execute( _
        "SELECT ISNULL(Dept, Department) AS Dept FROM " & DB_SCHEMA & "Users " & _
        "WHERE UserID = " & curRealUserID)
    If Not rsDept.EOF Then
        If InStr(UCase(CStr(rsDept("Dept"))), "BOSS") > 0 Then isOwner = True
    End If
    rsDept.Close : Set rsDept = Nothing
End If

' -- Today's date for min date on target date picker ----------
Dim todayStr : todayStr = fmtDate(Date())

' -- Fetch all active users for the assign-to dropdown --------
' Grouped by dept for optgroup display — matches prototype grouping
' Exclude: current user (use SELF instead), Mr. Problem (TypeID 5 auto-assigns)
' "Mr. Problem" identified by Dept containing "Meeting" or Username = "mrproblem"
Dim rsAll
Set rsAll = objConn.Execute( _
    "SELECT UserID, Username, IsManager, " & _
    "ISNULL(Dept, Department) AS Dept " & _
    "FROM " & DB_SCHEMA & "Users " & _
    "WHERE IsActive = 1 " & _
    "AND UserID <> " & curRealUserID & " " & _
    "AND LOWER(Username) <> 'mrproblem' " & _
    "ORDER BY IsManager DESC, Dept, Username")

' Build grouped arrays for optgroup rendering
' Groups: Owners, Managers, MER, Sales, Design, Purchase+Account, Operations
Dim grpOwner(50), grpMgr(50), grpMer(50), grpSales(50), grpDesign(50)
Dim grpPurchAcc(50), grpOps(50)
Dim cOwner, cMgr, cMer, cSales, cDesign, cPurchAcc, cOps
cOwner=0 : cMgr=0 : cMer=0 : cSales=0 : cDesign=0 : cPurchAcc=0 : cOps=0

' Each entry stores "UserID|Username|Dept"
Do While Not rsAll.EOF
     uID   = CLng(rsAll("UserID"))
     uName = CStr(rsAll("Username"))
     uMgr  = CBool(rsAll("IsManager"))
     uDept = CStr(rsAll("Dept"))
     entry = uID & "|" & uName & "|" & uDept
     deptUp = UCase(uDept)

    If uMgr And InStr(deptUp, "BOSS") > 0 Then
        ' Owner (boss-level manager)
        grpOwner(cOwner) = entry : cOwner = cOwner + 1
    ElseIf uMgr Then
        ' Non-boss manager
        grpMgr(cMgr) = entry : cMgr = cMgr + 1
    ElseIf InStr(deptUp, "MER") > 0 Then
        grpMer(cMer) = entry : cMer = cMer + 1
    ElseIf InStr(deptUp, "SALES") > 0 Then
        grpSales(cSales) = entry : cSales = cSales + 1
    ElseIf InStr(deptUp, "GRAPHIC") > 0 Then
        grpDesign(cDesign) = entry : cDesign = cDesign + 1
    ElseIf InStr(deptUp, "PURCHASE") > 0 Or InStr(deptUp, "ACCOUNT") > 0 Then
        grpPurchAcc(cPurchAcc) = entry : cPurchAcc = cPurchAcc + 1
    Else
        grpOps(cOps) = entry : cOps = cOps + 1
    End If
    rsAll.MoveNext
Loop
rsAll.Close : Set rsAll = Nothing

' -- Helper: render one optgroup from a filled array ----------
' Returns HTML string for all options in the group
Function renderOptGroup(label, grp, cnt)
    If cnt = 0 Then renderOptGroup = "" : Exit Function
    Dim html : html = "<optgroup label=""" & label & """>"
    Dim i, parts
    For i = 0 To cnt - 1
        parts = Split(grp(i), "|")
        html = html & "<option value=""" & parts(0) & """>" & _
               Server.HTMLEncode(parts(1)) & " (" & Server.HTMLEncode(parts(2)) & ")</option>"
    Next
    html = html & "</optgroup>"
    renderOptGroup = html
End Function

' -- Fetch job types for the type selector ---------------------
' Owner types (-2,-1,0) only shown when isOwner = true
Dim rsTypes
Set rsTypes = objConn.Execute( _
    "SELECT TypeID, TypeName, ShiftsAllowed, PenaltyPerDay " & _
    "FROM " & DB_SCHEMA & "JobTypes " & _
    "ORDER BY CAST(TypeID AS FLOAT)")
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">&#x2795; Assign New Job</h2>
    <p style="color:#6b7280;margin-bottom:30px;font-size:16px">
        Follow the steps below to assign a task.
        <% If Not isOwner Then %>
        <br><span style="color:#f59e0b">&#x26A0; Note: Only owners can assign jobs to owner-level managers</span>
        <% End If %>
    </p>

    <form method="POST" action="assign_job_action.asp" id="assignForm">

        <!-- -- Step 1: Assign To -- -->
        <div class="form-group" id="assignToField">
            <label style="font-size:16px">1&#xFE0F;&#x20E3; Assign To: <span style="color:#dc2626">*</span></label>
            <select id="newJobAssignTo" name="AssignTo" style="font-size:16px"
                    onchange="showWorkload(); handleAssignToChange();">
                <option value="">-- Select Person --</option>
                <option value="SELF">&#x1F512; ASSIGN TO MYSELF (Private Job)</option>

                <!-- Owners only shown if current user is also an owner -->
                <% If isOwner Then %>
                <%= renderOptGroup("&#x1F454; OWNERS", grpOwner, cOwner) %>
                <% End If %>

                <%= renderOptGroup("&#x1F454; MANAGERS",          grpMgr,      cMgr) %>
                <%= renderOptGroup("&#x1F4E6; MER TEAMS",         grpMer,      cMer) %>
                <%= renderOptGroup("&#x1F4BC; SALES TEAMS",       grpSales,    cSales) %>
                <%= renderOptGroup("&#x1F3A8; DESIGN TEAM",       grpDesign,   cDesign) %>
                <%= renderOptGroup("&#x1F4B0; PURCHASE & ACCOUNT",grpPurchAcc, cPurchAcc) %>
                <%= renderOptGroup("&#x1F527; OPERATIONS TEAMS",  grpOps,      cOps) %>
            </select>
        </div>

        <!-- -- Workload preview (filled by AJAX) -- -->
        <div id="workloadDisplay"></div>

        <!-- -- Step 2: Job Type -- -->
        <div class="form-group">
            <label style="font-size:16px">2&#xFE0F;&#x20E3; Job Type: <span style="color:#dc2626">*</span></label>
            <select id="newJobType" name="JobType" style="font-size:16px"
                    onchange="handleJobTypeChange()">
                <option value="">-- Select Job Type --</option>
                <%
                Do While Not rsTypes.EOF
                    Dim tID    : tID    = CStr(rsTypes("TypeID"))
                    Dim tName  : tName  = CStr(rsTypes("TypeName"))
                    Dim tShift : tShift = rsTypes("ShiftsAllowed")
                    Dim tPen   : tPen   = CLng(rsTypes("PenaltyPerDay"))

                    ' Owner types (-2,-1,0) only visible to owners
                    Dim showType : showType = True
                    If tID = "-2" Or tID = "-1" Or tID = "0" Then
                        If Not isOwner Then showType = False
                    End If

                    If showType Then
                        Dim shiftDisp
                        If IsNull(tShift) Then
                            shiftDisp = "&#x221E; shifts"
                        Else
                            shiftDisp = tShift & " shifts"
                        End If

                        Dim penDisp
                        If tPen = 0 Then
                            penDisp = "0 pts/day"
                        Else
                            penDisp = tPen & " pts/day"
                        End If

                        Dim optStyle : optStyle = ""
                        If tID = "-2" Or tID = "-1" Then optStyle = " style=""color:#dc2626;font-weight:700"""
                %>
                <option value="<%= tID %>"<%= optStyle %>>
                    <%= Server.HTMLEncode(tName) %> — <%= penDisp %> penalty | <%= shiftDisp %> allowed
                </option>
                <%
                    End If
                    rsTypes.MoveNext
                Loop
                rsTypes.Close : Set rsTypes = Nothing
                %>
            </select>
        </div>

        <!-- -- Type 5 info alert (shown when Mr. Problem selected) -- -->
        <div id="type5Alert" style="display:none" class="alert alert-info">
            <strong>&#x1F4CB; Mr. Problem — Meeting Agenda Item</strong><br>
            &bull; Automatically assigned to "Mr. Problem" for Monday meeting<br>
            &bull; Target date will be set to next Monday automatically<br>
            &bull; This will appear in Meeting Mode for discussion
        </div>

        <!-- -- Step 3: Urgency -- -->
        <div class="form-group" id="urgencyField">
            <label style="font-size:16px">3&#xFE0F;&#x20E3; Urgency Level:</label>
            <select id="newJobUrgency" name="Urgency" style="font-size:16px">
                <option value="normal">&#x1F4CC; Normal — Regular priority</option>
                <option value="urgent">&#x26A0; Urgent — High priority, respond soon</option>
                <option value="emergency">&#x1F534; Emergency — Critical, immediate action required</option>
            </select>
            <div class="alert alert-info" style="margin-top:10px;font-size:14px">
                <strong>&#x2139; Note:</strong> Urgency level is for communication only.
                It helps receivers understand priority but doesn't affect KPI penalties
                (those are determined by job type).
            </div>
        </div>

        <!-- -- Step 4: Title -- -->
        <div class="form-group">
            <label style="font-size:16px">4&#xFE0F;&#x20E3; Job Title: <span style="color:#dc2626">*</span></label>
            <input type="text" id="newJobTitle" name="Title"
                   placeholder="e.g., Sample approval for ABC Corporation"
                   style="font-size:16px">
        </div>

        <!-- -- Step 5: Target Date (hidden for type 5) -- -->
        <div class="form-group" id="targetDateField">
            <label style="font-size:16px">5&#xFE0F;&#x20E3; Target Date: <span style="color:#dc2626">*</span></label>
            <input type="date" id="newJobTarget" name="TargetDate"
                   min="<%= todayStr %>" style="font-size:16px">
        </div>

        <!-- -- Private checkbox (hidden for type 5) -- -->
        <div class="form-group" id="privateCheckField">
            <label style="display:flex;align-items:center;gap:10px;font-size:16px">
                <input type="checkbox" id="newJobPrivate" name="IsPrivate"
                       value="1" style="width:auto;height:20px">
                <span>&#x1F512; Make PRIVATE (only works if assigning to yourself —
                nobody else will see this job)</span>
            </label>
        </div>

        <div class="alert alert-info">
            <strong>&#x1F4A1; Job Assignment Rules:</strong><br>
            &bull; Choose appropriate job type as it affects penalty points if late<br>
            &bull; Private jobs only work when you assign to yourself<br>
            &bull; Target date should be realistic for the job complexity<br>
            &bull; Job giver must accept completion when receiver submits
        </div>

        <button type="submit" class="btn btn-success"
                style="font-size:18px;padding:18px 40px;margin-top:20px">
            &#x2705; Assign This Job
        </button>

    </form>
</div>

</div><!-- /container -->
</body>


<script>
// -- handleJobTypeChange() -------------------------------------
// Show/hide fields based on selected job type.
// Type 5 (Mr. Problem): hide assign-to, date, private — auto Monday.
function handleJobTypeChange() {
    var type = document.getElementById('newJobType').value;
    var assignToField    = document.getElementById('assignToField');
    var targetDateField  = document.getElementById('targetDateField');
    var privateCheckField= document.getElementById('privateCheckField');
    var type5Alert       = document.getElementById('type5Alert');
    var urgencyField     = document.getElementById('urgencyField');

    if (type === '5') {
        assignToField.style.display     = 'none';
        targetDateField.style.display   = 'none';
        privateCheckField.style.display = 'none';
        type5Alert.style.display        = 'block';
    } else {
        assignToField.style.display     = 'block';
        targetDateField.style.display   = 'block';
        privateCheckField.style.display = 'block';
        type5Alert.style.display        = 'none';
    }
}

// -- handleAssignToChange() ------------------------------------
// When assigning to an Owner, force job type to Type 0 (no penalty).
// Owners should only receive Type 0 jobs per prototype rule.
function handleAssignToChange() {
    var assignTo = document.getElementById('newJobAssignTo').value;
    var typeSelect = document.getElementById('newJobType');

    // Check if selected user is an owner by looking at option text
    var selectedOption = typeSelect.options[typeSelect.selectedIndex];

    // We pass owner IDs from server as a data attribute on the select
    var ownerIDs = document.getElementById('newJobAssignTo')
                           .getAttribute('data-owner-ids') || '';

    if (ownerIDs.split(',').indexOf(assignTo) !== -1) {
        // Auto-select Type 0 for owners
        for (var i = 0; i < typeSelect.options.length; i++) {
            if (typeSelect.options[i].value === '0') {
                typeSelect.selectedIndex = i;
                break;
            }
        }
        // Show info
        var info = document.getElementById('ownerTypeInfo');
        if (!info) {
            info = document.createElement('div');
            info.id = 'ownerTypeInfo';
            info.className = 'alert alert-info';
            info.style.marginTop = '10px';
            info.innerHTML = '<strong>&#x2139; Owner Assignment:</strong> Jobs to owners use Type 0 automatically (no KPI penalty). Use Urgency Level to communicate priority.';
            typeSelect.parentElement.appendChild(info);
        }
    } else {
        var info = document.getElementById('ownerTypeInfo');
        if (info) info.remove();
    }
}

// -- showWorkload() --------------------------------------------
// AJAX call to workload_ajax.asp when assignee changes.
// Shows the selected person's current pending jobs.
function showWorkload() {
    var assignTo = document.getElementById('newJobAssignTo').value;
    var display  = document.getElementById('workloadDisplay');

    if (!assignTo || assignTo === '' || assignTo === 'SELF') {
        display.innerHTML = '';
        return;
    }

    // Show loading state
    display.innerHTML = '<div class="alert alert-info">&#x23F3; Loading workload...</div>';

    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'workload_ajax.asp?UserID=' + encodeURIComponent(assignTo), true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4) {
            if (xhr.status === 200) {
                display.innerHTML = xhr.responseText;
            } else {
                display.innerHTML = '';
            }
        }
    };
    xhr.send();
}

// -- Form validation before submit ----------------------------
document.getElementById('assignForm').onsubmit = function (e) {
    var type = document.getElementById('newJobType').value;
    var title = document.getElementById('newJobTitle').value.trim();
    var assignTo = document.getElementById('newJobAssignTo').value;
    var target = document.getElementById('newJobTarget').value;
    var isPrivate = document.getElementById('newJobPrivate').checked;

    if (!type) {
        alert('Please select a Job Type.');
        document.getElementById('newJobType').focus();
        e.preventDefault(); return false;
    }
    if (!title) {
        alert('Please enter a Job Title.');
        document.getElementById('newJobTitle').focus();
        e.preventDefault(); return false;
    }

    if (type !== '5') {
        if (!assignTo) {
            alert('Please select who to assign this job to.');
            document.getElementById('newJobAssignTo').focus();
            e.preventDefault(); return false;
        }
        if (!target) {
            alert('Please select a Target Date.');
            document.getElementById('newJobTarget').focus();
            e.preventDefault(); return false;
        }
        if (isPrivate && assignTo !== 'SELF') {
            alert('Private jobs can only be assigned to yourself.\n\nPlease either uncheck "Make PRIVATE" or change "Assign To" to yourself (SELF).');
            e.preventDefault(); return false;
        }
    }
    return true;
};
</script>

</html>
<% objConn.Close : Set objConn = Nothing %>
