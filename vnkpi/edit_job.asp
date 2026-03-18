<%@ Language="VBScript" CODEPAGE="874" %>
<% Response.CharSet = "windows-874" : Response.CodePage = 874 %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
Dim activePage : activePage = "assigned"
%>
<!-- #include file="header.asp" -->
<%
' ============================================================
' edit_job.asp — Edit Pending Job
' Allows job assigners to edit Title and TargetDate for jobs
' in Status = 'Pending' only.
' ============================================================

Dim jobID : jobID = Request.QueryString("JobID")
If jobID = "" Or Not IsNumeric(jobID) Then
    Response.Write "<div class='card'><h3 style='color:#dc2626;'>Invalid Job ID</h3></div>"
    Response.End
End If

' -- Fetch the job and verify ownership and status --------
Dim rsJob, sqlJob
sqlJob = "SELECT j.JobID, j.Title, j.TargetDate, j.Status, j.TypeID, " & _
         "       jt.TypeName, u.Username AS AssignedToName " & _
         "FROM " & DB_SCHEMA & "Jobs j " & _
         "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
         "JOIN " & DB_SCHEMA & "Users u ON j.AssignedTo = u.UserID " & _
         "WHERE j.JobID = " & CLng(jobID) & " AND j.AssignedBy = " & curRealUserID

Set rsJob = objConn.Execute(sqlJob)

If rsJob.EOF Then
    Response.Write "<div class='card'><h3 style='color:#dc2626;'>&#x274C; Job not found or you don't have permission to edit it.</h3></div>"
    Response.End
End If

' -- Check if job is still in Pending status --------
If rsJob("Status") <> "Pending" Then
    Response.Write "<div class='card'><h3 style='color:#dc2626;'>&#x26A0; Cannot edit this job. Status is already '" & rsJob("Status") & "'.</h3>" & _
    "<p>Only jobs in 'Pending' status can be edited.</p></div>"
    Response.End
End If

' -- Extract job data --------
Dim jTitle      : jTitle      = rsJob("Title")
Dim jTarget     : jTarget     = rsJob("TargetDate")
Dim jTypeName   : jTypeName   = rsJob("TypeName")
Dim jAssignedTo : jAssignedTo = rsJob("AssignedToName")
Dim todayStr    : todayStr    = fmtDate(Date())

rsJob.Close : Set rsJob = Nothing
%>

<div class="card">
    <h2 style="font-size:28px;margin-bottom:20px">
        &#x270F; Edit Job #<%= jobID %>
    </h2>

    <div class="alert alert-info" style="margin-bottom:30px">
        <strong>&#x1F4A1; Job Details (Read-only):</strong><br>
        &bull; <strong>Assigned To:</strong> <%= Server.HTMLEncode(jAssignedTo) %><br>
        &bull; <strong>Job Type:</strong> <%= Server.HTMLEncode(jTypeName) %>
    </div>

    <form method="POST" action="edit_job_action.asp" id="editForm">
        <input type="hidden" name="JobID" value="<%= jobID %>">

        <!-- -- Step 1: Job Title -- -->
        <div class="form-group">
            <label style="font-size:16px">1&#xFE0F;&#x20E3; Job Title (Add or modify tasks): <span style="color:#dc2626">*</span></label>
            <input type="text" id="editJobTitle" name="Title" 
                   value="<%= Server.HTMLEncode(jTitle) %>" 
                   style="font-size:16px" required>
            <small style="color:#6b7280;display:block;margin-top:5px">
                Update the job description or add additional tasks as needed.
            </small>
        </div>

        <!-- -- Step 2: Target Date -- -->
        <div class="form-group">
            <label style="font-size:16px">2&#xFE0F;&#x20E3; Target Date: <span style="color:#dc2626">*</span></label>
            <input type="date" id="editJobTarget" name="TargetDate" 
                   value="<%= fmtDate(jTarget) %>" 
                   min="<%= todayStr %>" 
                   style="font-size:16px" required>
            <small style="color:#6b7280;display:block;margin-top:5px">
                The target date must be today or later.
            </small>
        </div>

        <div class="alert alert-info" style="margin-top:20px;margin-bottom:20px">
            <strong>&#x2139; Edit Rules:</strong><br>
            &bull; You can only edit jobs that are in 'Pending' status<br>
            &bull; Once the job is submitted, it will move to 'Submitted' and cannot be edited<br>
            &bull; Target date changes will be reflected in the receiver's view
        </div>

        <div style="display:flex;gap:10px;margin-top:30px">
            <button type="submit" class="btn btn-primary" style="font-size:16px;padding:15px 30px">
                &#x270F; Save Changes
            </button>
            <a href="assigned_list.asp" class="btn btn-secondary" style="font-size:16px;padding:15px 30px;text-decoration:none;display:inline-flex;align-items:center">
                &#x274C; Cancel
            </a>
        </div>
    </form>
</div>

</div><!-- /container -->
</body>
</html>
<%
objConn.Close : Set objConn = Nothing
%>