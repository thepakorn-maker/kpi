<%@ Language="VBScript" CODEPAGE="874" %>
<%
    Response.CharSet = "windows-874" 
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
' myjobs.asp - My Jobs tab
' Shows 5 sections: 
'   0. Waiting for Approval (NEW)
'   1. Target Today / Overdue / Future / Private
' ============================================================

' -- Count each section for section headers ------------------
Dim cntWaitingApproval, cntRejected, cntToday, cntOverdue, cntFuture, cntPrivate
Dim rs, sql

' Count: Waiting for Approval
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Submitted'"
Set rs = objConn.Execute(sql)
cntWaitingApproval = rs(0) : rs.Close : Set rs = Nothing

' Count: Rejected jobs
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Pending' AND RejectComment IS NOT NULL"
Set rs = objConn.Execute(sql)
cntRejected = rs(0) : rs.Close : Set rs = Nothing

' Count: Target today
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Pending' AND RejectComment IS NULL " & _
      "AND CAST(TargetDate AS DATE) = CAST(GETDATE() AS DATE)"
Set rs = objConn.Execute(sql)
cntToday = rs(0) : rs.Close : Set rs = Nothing

' Count: Overdue
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Pending' AND RejectComment IS NULL " & _
      "AND CAST(GETDATE() AS DATE) > CAST(TargetDate AS DATE)"
Set rs = objConn.Execute(sql)
cntOverdue = rs(0) : rs.Close : Set rs = Nothing

' Count: Future
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Pending' AND RejectComment IS NULL " & _
      "AND CAST(TargetDate AS DATE) > CAST(GETDATE() AS DATE) AND IsPrivate=0"
Set rs = objConn.Execute(sql)
cntFuture = rs(0) : rs.Close : Set rs = Nothing

' Count: Private
sql = "SELECT COUNT(*) FROM " & DB_SCHEMA & "Jobs WHERE AssignedTo=" & curUserID & " " & _
      "AND Status='Pending' AND RejectComment IS NULL AND IsPrivate=1"
Set rs = objConn.Execute(sql)
cntPrivate = rs(0) : rs.Close : Set rs = Nothing

' -- Shared SQL columns for job queries ----------------------
Dim sqlSelect
sqlSelect = "SELECT j.JobID, j.Title, j.TypeID AS JobTypeID, j.Urgency, " & _
            "       j.TargetDate, j.OriginalTargetDate, j.ShiftsUsed, " & _
            "       j.LastPenaltyDate, j.RejectComment, j.RejectDate, " & _
            "       j.UpdatedAt, j.SubmitDate, " & _
            "       jt.TypeName, jt.PenaltyPerDay, jt.ShiftsAllowed AS MaxShifts, " & _
            "       ab.Username AS AssignedByName, ab.UserID AS AssignedByID " & _
            "FROM " & DB_SCHEMA & "Jobs j " & _
            "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
            "JOIN " & DB_SCHEMA & "Users ab ON j.AssignedBy = ab.UserID " & _
            "WHERE j.AssignedTo = " & curUserID & " AND j.Status = 'Pending' "

' -- canShiftJob() -------------------------------------------
Function canShiftJob(typeID, assignedByID, maxShifts, shiftsUsed)
    Dim isGiver : isGiver = (CLng(assignedByID) = curUserID)

    If curIsManager Then canShiftJob = True : Exit Function

    If maxShifts <> -1 Then
        If (maxShifts - shiftsUsed) <= 0 Then
            canShiftJob = False
            Exit Function
        End If
    End If

    Select Case CStr(typeID)
        Case "-2", "-1"        : canShiftJob = isGiver
        Case "0"               : canShiftJob = True
        Case "1", "2", "3"    : canShiftJob = isGiver
        Case "4.1","4.2","4.3" : canShiftJob = True
        Case "5"               : canShiftJob = True
        Case Else              : canShiftJob = False
    End Select
End Function

' -- renderJobTypePill() ------------------------------------
Function renderJobTypePill(typeID, typeName)
    Dim cssClass : cssClass = "jt-" & Replace(CStr(typeID), ".", "")
    renderJobTypePill = "<span class=""" & cssClass & """>" & Server.HTMLEncode(CStr(typeName)) & "</span>"
End Function

' -- renderShiftsCount() ------------------------------------
Function renderShiftsCount(shiftsUsed, maxShifts)
    If CLng(maxShifts) = -1 Then
        renderShiftsCount = shiftsUsed & " / &infin;"
    Else
        renderShiftsCount = shiftsUsed & " / " & maxShifts
    End If
End Function

' -- renderShiftButton() ------------------------------------
Function renderShiftButton(jobID, origTarget, currTarget, shiftsUsed, maxShifts, title, assignedByName, typeID, assignedByID, canShift, btnLabel)
    Dim maxDisp
    If CLng(maxShifts) = -1 Then maxDisp = "Unlimited" Else maxDisp = CStr(maxShifts)

    Dim safeTitle : safeTitle = Replace(CStr(title), "'", "\'")
    Dim safeName  : safeName  = Replace(CStr(assignedByName), "'", "\'")

    If canShift Then
        renderShiftButton = "<button class=""btn btn-warning"" " & _
            "onclick=""openShiftModal(" & jobID & ",'" & fmtDate(origTarget) & "','" & fmtDate(currTarget) & "'," & _
            shiftsUsed & ",'" & maxDisp & "','" & safeTitle & "','" & safeName & "','" & typeID & "'," & assignedByID & ")"" >" & _
            "&#x1F4C5; " & btnLabel & "</button>"
    Else
        renderShiftButton = "<span style=""color:#dc2626;font-size:11px"">&#x2717; No shifts</span>"
    End If
End Function

' -- renderSubmitForm() ------------------------------------
Function renderSubmitForm(jobID, btnClass, btnLabel)
    Dim confirmMsg
    If InStr(btnClass, "danger") > 0 Then
        confirmMsg = "Accept this job LATE? Penalties will apply. Are you sure?"
    Else
        confirmMsg = "Submit this job as complete. Are you sure?"
    End If
    renderSubmitForm = "<form method=""POST"" action=""submit_job.asp"" style=""display:inline"" " & _
        "onsubmit=""return confirm('" & confirmMsg & "')"">" & _
        "<input type=""hidden"" name=""JobID"" value=""" & jobID & """>" & _
        "<button type=""submit"" class=""btn " & btnClass & """>" & btnLabel & "</button>" & _
        "</form>"
End Function

' -- getResultBadge() ----------------------------------------
Function getResultBadge(targetDate, submitDate, penaltyPerDay)
    Dim daysDiff : daysDiff = DateDiff("d", CDate(targetDate), CDate(submitDate))
    Dim penalty : penalty = 0
    If daysDiff > 0 And CLng(penaltyPerDay) < 0 Then penalty = daysDiff * CLng(penaltyPerDay)

    Dim resultHTML, resultColor, rowClass
    If daysDiff > 0 Then
        resultColor = "#dc2626"
        resultHTML = "&#x26A0; " & daysDiff & " days LATE<br><small>(" & penalty & " pts)</small>"
        rowClass = "overdue-row"
    ElseIf daysDiff < 0 Then
        resultColor = "#059669"
        resultHTML = "&#x1F389; " & Abs(daysDiff) & " days EARLY!"
        rowClass = "early-finish"
    Else
        resultColor = "#059669"
        resultHTML = "&#x2705; On-Time"
        rowClass = ""
    End If

    getResultBadge = "<div style='color:" & resultColor & ";font-weight:700;font-size:14px'>" & resultHTML & "</div>|" & rowClass
End Function
%>

<!-- SECTION 0: WAITING FOR APPROVAL -->
<% If cntWaitingApproval > 0 Then %>
<div class="card">
    <h2 style="color:#2563eb;font-size:24px">&#x23F3; WAITING FOR APPROVAL (<%= cntWaitingApproval %>)</h2>
    <div class="alert alert-info">
        <strong>&#x1F4A1; Your Submission Status:</strong><br>
        These jobs have been submitted and are waiting for your manager's approval.<br>
        &bull; <span style="color:#059669">Green = Submitted Early</span> &nbsp;&nbsp;
        <span style="color:#000">Black = Submitted On-Time</span> &nbsp;&nbsp;
        <span style="color:#dc2626">Red = Submitted Late (penalties will apply)</span>
    </div>
    <table>
        <tr>
            <th>ID</th>
            <th>Job Type</th>
            <th>Title</th>
            <th>Assigned By</th>
            <th>Target Date</th>
            <th>Submit Date</th>
            <th>Result</th>
        </tr>
        <%
        Dim sqlSubmitted
        sqlSubmitted = "SELECT j.JobID, j.Title, j.TypeID AS JobTypeID, j.Urgency, " & _
                       "       j.TargetDate, j.SubmitDate, j.UpdatedAt, " & _
                       "       jt.TypeName, jt.PenaltyPerDay, " & _
                       "       ab.Username AS AssignedByName " & _
                       "FROM " & DB_SCHEMA & "Jobs j " & _
                       "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
                       "JOIN " & DB_SCHEMA & "Users ab ON j.AssignedBy = ab.UserID " & _
                       "WHERE j.AssignedTo = " & curUserID & " AND j.Status = 'Submitted' " & _
                       "ORDER BY j.SubmitDate DESC"
        
        Set rs = objConn.Execute(sqlSubmitted)
        Do While Not rs.EOF
            Dim resultParts : resultParts = Split(getResultBadge(rs("TargetDate"), rs("SubmitDate"), rs("PenaltyPerDay")), "|")
            Dim resultHTML : resultHTML = resultParts(0)
            Dim rowClass : rowClass = resultParts(1)
        %>
        <tr class="<%= rowClass %>">
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td>
                <%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong>
                <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
                <% End If %>
            </td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td><%= fmtDate(rs("TargetDate")) %></td>
            <td><%= fmtDate(rs("SubmitDate")) %></td>
            <td><%= resultHTML %></td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>

<!-- SECTION: REJECTED JOBS -->
<% If cntRejected > 0 Then %>
<div class="card" style="border: 2px solid #dc2626;">
    <h2 style="color:#dc2626;font-size:24px">&#x2717; REJECTED &mdash; NEEDS REVISION (<%= cntRejected %>)</h2>
    <div class="alert alert-danger">
        <strong>&#x26A0; Action Required:</strong> Your job was rejected! Please read the comment and resubmit.
    </div>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title &amp; Feedback</th>
            <th>Assigned By</th><th>Target</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND j.RejectComment IS NOT NULL ORDER BY j.RejectDate DESC")
        Do While Not rs.EOF
        %>
        <tr style="background-color: #fef2f2;">
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td>
                <%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong>
                <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated by Manager on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
                <% End If %>
                <div style="background:#fee2e2; color:#991b1b; padding:8px; margin-top:8px; border-radius:4px; font-size:13px; border-left:3px solid #dc2626;">
                    <strong>&#x1F4AC; Manager Comment (<%= sys_date(rs("RejectDate")) %>):</strong><br>
                    <%= safeStr(rs("RejectComment")) %>
                </div>
            </td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td style="color:#ea580c;font-weight:700"><%= sys_date(rs("TargetDate")) %></td>
            <td>
                <%= renderSubmitForm(rs("JobID"), "btn-success", "&#x2705; Resubmit") %>
            </td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>

<!-- SECTION 1: TARGET TODAY -->
<% If cntToday > 0 Then %>
<div class="card">
    <h2 style="color:#ea580c;font-size:24px">&#x1F514; TARGET DATE TODAY - LAST CHANCE TO SHIFT! (<%= cntToday %>)</h2>
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
        Set rs = objConn.Execute(sqlSelect & "AND j.RejectComment IS NULL AND CAST(j.TargetDate AS DATE) = CAST(GETDATE() AS DATE) ORDER BY j.TargetDate")
        Do While Not rs.EOF
            Dim canShift : canShift = canShiftJob(rs("JobTypeID"), rs("AssignedByID"), rs("MaxShifts"), rs("ShiftsUsed"))
        %>
        <tr class="target-today-row">
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong>
            <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated by Manager on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
            <% End If %>
            </td>
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

<!-- SECTION 2: OVERDUE -->
<% If cntOverdue > 0 Then %>
<div class="card">
    <h2 style="color:#dc2626;font-size:24px">&#x1F6A8; OVERDUE - CANNOT SHIFT! (<%= cntOverdue %>)</h2>
    <div class="alert alert-danger">
        <strong>&#x2717; TOO LATE TO SHIFT!</strong><br>
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
        Dim alreadyDeduct, daysDeduct, remaining
        Set rs = objConn.Execute(sqlSelect & "AND j.RejectComment IS NULL AND CAST(j.TargetDate AS DATE) &lt; CAST(GETDATE() AS DATE) ORDER BY j.TargetDate")
        Do While Not rs.EOF
            Dim daysLate : daysLate = DateDiff("d", (rs("TargetDate")), Date())
            Dim totalPen : totalPen = daysLate * CLng(rs("PenaltyPerDay"))
            alreadyDeduct = 0
            If Not IsNull(rs("LastPenaltyDate")) Then
                daysDeduct = DateDiff("d", rs("TargetDate"), rs("LastPenaltyDate"))
                If daysDeduct > 0 Then
                    alreadyDeduct = daysDeduct * CLng(rs("PenaltyPerDay"))
                End If
            End If
            remaining = totalPen - alreadyDeduct
        %>
        <tr class="overdue-row">
            <td><strong>#<%= rs("JobID") %></strong></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= getUrgencyBadge(rs("Urgency")) %><strong><%= safeStr(rs("Title")) %></strong>
            <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated by Manager on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
            <% End If %>            
            </td>
            <td><%= safeStr(rs("AssignedByName")) %></td>
            <td style="color:#dc2626;font-weight:700"><%= sys_date(rs("TargetDate")) %> </td>
            <td style="color:#dc2626;font-weight:700;font-size:16px"><%= daysLate %> days &#x26A0;</td>
            <td style="color:#dc2626;font-weight:700;font-size:16px">
                <%= totalPen %> pts<br>
                <small style="color:#6b7280">
                    Deducted: <%= alreadyDeduct %> pts<br>
                    Remaining: <%= remaining %> pts
                </small>
                <small>(-<%= rs("PenaltyPerDay") %>/day)</small>
            </td>
            <td>
                <%= renderSubmitForm(rs("JobID"), "btn-danger", "Accept Late") %>
                <div style="color:#dc2626;font-size:10px;margin-top:5px">&#x2717; Cannot shift</div>
            </td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>

<!-- SECTION 3: FUTURE JOBS -->
<% If cntFuture > 0 Then %>
<div class="card">
    <h2 style="font-size:24px">&#x1F4CB; Future Jobs - Can Shift Anytime (<%= cntFuture %>)</h2>
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
        Set rs = objConn.Execute(sqlSelect & "AND j.RejectComment IS NULL AND CAST(j.TargetDate AS DATE) &gt; CAST(GETDATE() AS DATE) AND j.IsPrivate=0 ORDER BY j.TargetDate")
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
            <td><%= getUrgencyBadge(rs("Urgency")) %><%= safeStr(rs("Title")) %>
            <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated by Manager on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
            <% End If %>
            </td>
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

<!-- SECTION 4: PRIVATE JOBS -->
<% If cntPrivate > 0 Then %>
<div class="card">
    <h2 style="font-size:24px">&#x1F512; My Private Jobs (<%= cntPrivate %>)</h2>
    <p style="color:#6b7280;margin-bottom:15px">Only you can see these jobs that you assigned to yourself</p>
    <table>
        <tr>
            <th>ID</th><th>Job Type</th><th>Title</th><th>Target Date</th><th>Actions</th>
        </tr>
        <%
        Set rs = objConn.Execute(sqlSelect & "AND j.RejectComment IS NULL AND j.IsPrivate=1 ORDER BY j.TargetDate")
        Do While Not rs.EOF
        %>
        <tr>
            <td><strong>#<%= rs("JobID") %></strong> <span class="private-badge">PRIVATE</span></td>
            <td><%= renderJobTypePill(rs("JobTypeID"), rs("TypeName")) %></td>
            <td><%= safeStr(rs("Title")) %>
            <% If Not IsNull(rs("UpdatedAt")) Then %>
                    <div style="font-size:12px; color:#2563eb; margin-top:6px; padding:4px 8px; background:#eff6ff; border-left:3px solid #3b82f6; border-radius:4px; display:inline-block;">
                        &#x270F; <em>Updated by Manager on <%= sys_date(rs("UpdatedAt")) %></em>
                    </div>
            <% End If %>
            </td>
            <td><%= sys_date(rs("TargetDate")) %></td>
            <td><%= renderSubmitForm(rs("JobID"), "btn-success", "&#x2705; Submit") %></td>
        </tr>
        <% rs.MoveNext : Loop : rs.Close : Set rs = Nothing %>
    </table>
</div>
<% End If %>

<!-- ALL CLEAR -->
<% If cntWaitingApproval = 0 And cntToday = 0 And cntOverdue = 0 And cntFuture = 0 And cntPrivate = 0 Then %>
<div class="card">
    <div style="text-align:center;padding:60px">
        <div style="font-size:80px;margin-bottom:20px">&#x1F389;</div>
        <h2 style="font-size:28px;color:#10b981;margin-bottom:10px">All Clear!</h2>
        <p style="font-size:18px;color:#6b7280">You have no pending jobs. Great work!</p>
    </div>
</div>
<% End If %>

<!-- JOB TYPES REFERENCE TABLE -->
<% If cntWaitingApproval > 0 Or cntToday > 0 Or cntOverdue > 0 Or cntFuture > 0 Then %>
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

<!-- SHIFT DATE MODAL -->
<div id="shiftModal" class="modal">
    <div class="modal-content">
        <span class="close-btn" onclick="document.getElementById('shiftModal').classList.remove('active')">&times;</span>
        <h2 style="margin-bottom:20px">&#x1F4C5; Shift Target Date</h2>

        <div id="shiftJobDetails"></div>

        <div class="form-group">
            <label>Original Target Date:</label>
            <input type="text" id="origTargetDisp" readonly style="background:#f3f4f6">
        </div>
        <div class="form-group">
            <label>Current Target Date:</label>
            <input type="text" id="currTargetDisp" readonly style="background:#f3f4f6">
        </div>
        <div class="form-group">
            <label>Shifts Remaining:</label>
            <input type="text" id="shiftsDisp" readonly style="background:#f3f4f6">
        </div>
        <div class="form-group">
            <label>New Target Date (must be &ge; today): <span style="color:#dc2626">*</span></label>
            <input type="date" id="newTargetDate">
        </div>

        <div class="form-group" id="shiftReasonSection" style="display:none">
            <label id="shiftReasonLabel">&#x1F4AC; Reason for Shift: <span style="color:#dc2626">*</span></label>
            <textarea id="shiftReason" rows="3" placeholder="Explain why you need to shift the date..." style="font-size:14px;padding:10px"></textarea>
        </div>

        <div class="alert alert-warning">
            <strong>&#x1F4CB; Date Change Rules:</strong><br>
            &bull; <strong>Types 1, 2, 3:</strong> Only JOB GIVER can change date (Type 1 needs reason)<br>
            &bull; <strong>Types 4.x &amp; 5:</strong> JOB RECEIVER must write reason to shift<br>
            &bull; <strong>Managers:</strong> Can override in special cases
        </div>

        <input type="hidden" id="shiftJobID">
        <input type="hidden" id="shiftCurrTarget">

        <button class="btn btn-warning" onclick="confirmShift()">&#x1F4C5; Shift Date</button>
        <button class="btn btn-secondary" onclick="document.getElementById('shiftModal').classList.remove('active')">Cancel</button>
    </div>
</div>

<form id="shiftForm" method="POST" action="shift_job.asp" style="display:none">
    <input type="hidden" id="sfJobID"   name="JobID">
    <input type="hidden" id="sfNewDate" name="NewDate">
    <input type="hidden" id="sfReason"  name="Reason">
</form>

<script>
var todayStr  = '<%= fmtDate(Date()) %>';
var curUserID = <%= curUserID %>;

function openShiftModal(jobID, origTarget, currTarget, shiftsUsed, maxShifts, title, assignedByName, jobTypeID, assignedByID) {
    document.getElementById('shiftJobID').value      = jobID;
    document.getElementById('shiftCurrTarget').value = currTarget;
    document.getElementById('origTargetDisp').value  = origTarget;
    document.getElementById('currTargetDisp').value  = currTarget;
    document.getElementById('shiftsDisp').value      = shiftsUsed + ' used / ' + maxShifts + ' max';
    document.getElementById('newTargetDate').min     = todayStr;
    document.getElementById('newTargetDate').value   = '';
    document.getElementById('shiftReason').value     = '';

    var isGiver     = (curUserID === parseInt(assignedByID));
    var needsReason = false;
    var reasonLabel = '&#x1F4AC; Reason for Shift (Required): <span style="color:#dc2626">*<\/span>';

    if (jobTypeID === '1' && isGiver) {
        needsReason = true;
        reasonLabel = '&#x26A0; Reason for Shift - Required for Type 1: <span style="color:#dc2626">*<\/span>';
    } else if (jobTypeID === '4.1' || jobTypeID === '4.2' || jobTypeID === '4.3' || jobTypeID === '5') {
        if (!isGiver) {
            needsReason = true;
        }
    }

    document.getElementById('shiftReasonSection').style.display = needsReason ? 'block' : 'none';
    document.getElementById('shiftReasonLabel').innerHTML = reasonLabel;

    var rulesMsg = '';
    if (jobTypeID === '1' || jobTypeID === '2' || jobTypeID === '3') {
        rulesMsg = '&bull; Types 1, 2, 3: Only the <strong>Job Giver<\/strong> can change the date';
        if (jobTypeID === '1') rulesMsg += '<br>&bull; Type 1 requires a written reason';
    } else if (jobTypeID === '4.1' || jobTypeID === '4.2' || jobTypeID === '4.3' || jobTypeID === '5') {
        rulesMsg = '&bull; Planning &amp; Problem jobs: <strong>Job Receiver<\/strong> must write a reason for shifting';
    } else {
        rulesMsg = '&bull; This job type: anyone can shift freely without a reason';
    }

    document.getElementById('shiftJobDetails').innerHTML =
        '<div class="alert alert-info">' +
            '<strong>Job #' + jobID + ': ' + title + '<\/strong><br>' +
            'Assigned By: ' + assignedByName + '<br>' +
            'Original Target: ' + origTarget + ' &rarr; Current Target: ' + currTarget +
        '<\/div>' +
        '<div class="alert alert-warning">' +
            '<strong>&#x1F4CB; Date Change Rules:<\/strong><br>' + rulesMsg +
        '<\/div>';

    document.getElementById('shiftModal').classList.add('active');
}

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

    document.getElementById('sfJobID').value   = document.getElementById('shiftJobID').value;
    document.getElementById('sfNewDate').value = newDate;
    document.getElementById('sfReason').value  = reason;

    if (!confirm('Confirm shift target date to ' + newDate + '?\nThis will update the DB.')) return;
    document.getElementById('shiftForm').submit();
}
</script>

</div><!-- /container -->

</body>
</html>
<% objConn.Close : Set objConn = Nothing %>