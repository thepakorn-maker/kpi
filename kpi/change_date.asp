<%@ Language=VBScript %>

<%
UserID=request.cookies("UserID")
%>
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Change Target Date - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
    <link rel="stylesheet" href="/js/jquery-ui.css">
 <script src="https://code.jquery.com/jquery-3.7.1.js"></script>
  <script src="/js/jquery-ui.js"></script>

</head>
<body>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <%
    Dim jobID : jobID = CLng(Request.QueryString("jobID"))
    
    'Dim sqlJob : sqlJob = "SELECT j.*, u.Name AS AssignedName FROM Jobs j INNER JOIN Users u ON j.AssignedTo = u.UserID WHERE j.JobID = " & jobID
    'Dim rsJob : Set rsJob = conn.Execute(sqlJob)

    Dim sqlJob : sqlJob = "SELECT j.*, jt.ShiftsAllowed, jt.typename,u.Name AS AssignedName FROM Jobs j " & _
                      "INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
                      "INNER JOIN Users u ON j.AssignedTo = u.UserID " & _
                      "WHERE j.JobID = " & jobID
     Dim rsJob : Set rsJob = conn.Execute(sqlJob)
    
    If rsJob.EOF Then
        Response.Write "<h2 style='text-align:center; color:red;'>Job not found</h2>"
        Response.End
    End If
    
    ' ตรวจสอบสิทธิ์: เป็นเจ้าของงาน หรือ Manager และถึงวันแล้ว
    Dim canChange : canChange = False
    Dim currentDate : currentDate = Date()
    Dim targetDate : targetDate = (rsJob("TargetDate"))
    Dim shiftsAllowed : shiftsAllowed = rsJob("ShiftsAllowed")
    Dim shiftsUsed : shiftsUsed = clng("0" & rsJob("ShiftsUsed"))
    'response.write iso_date(targetDate) & " " & iso_date(currentDate)
    'response.end
    If LCase(rsJob("AssignedTo")) = LCase(Request.Cookies("UserID")) And iso_date(currentDate) >= iso_date(targetDate) Then
        If IsNull(shiftsAllowed) Then
            canChange = True  ' Unlimited
        ElseIf shiftsUsed < shiftsAllowed Then
            canChange = True
        End If
        
    ElseIf Request.Cookies("IsManager") Then
        canChange = True  ' Manager สามารถเปลี่ยนได้ทุกเมื่อ (ตามสเปก override)
    End If
    %>
<!-- แสดงข้อมูล Shifts -->
<div style="margin: 1rem 0; padding: 1rem; background: #fefce8; border-radius: 8px;">
    Job Type: <%= rsJob("TypeName") %><br>
    Shifts Allowed: <%= IIf(IsNull(shiftsAllowed), "Unlimited", shiftsAllowed) %><br>
    Shifts Used: <%= shiftsUsed %><br>
    <% If Not canChange And Not request.cookies("IsManager") Then %>
        <strong style="color: #dc2626;">ไม่สามารถเลื่อนวันที่ได้: ถึง limit หรือยังไม่ถึงวันกำหนด</strong>
    <% End If %>
</div>

<%
    If Request.Form("action") = "change" and Request.Form("newTargetDate")<>Request.Form("oldTargetDate") _
       and Request.Form("newTargetDate")<>"" and canChange Then
        Dim newDate : newDate = ora_date(Request.Form("newTargetDate"))
        oldDate = ora_date(Request.Form("oldTargetDate"))
        sql="INSERT into jobshifthistory(jobid,oldtargetdate,newtargetdate,ShiftedBy)"
        sql=sql & " values('" & jobID & "','" & oldDate & "','" & newDate & "','" & UserID & "')"
        conn.Execute(sql)
        sql="UPDATE Jobs SET TargetDate = '" & newDate & "'"
        If not(request.cookies("IsManager"))  And Not IsNull(shiftsAllowed) Then
          sql=sql & ", ShiftsUsed = ShiftsUsed + 1"
        end if
        sql=sql & " WHERE JobID = " & jobID
        conn.Execute(sql)
        Response.Redirect "dashboard.asp"
    End If
    %>

    <div class="user-card" style="max-width: 600px; margin: 2rem auto;">
        <h2 style="text-align:center; margin-bottom: 1rem;">Change Target Date</h2>
        
        <div class="job-card">
            <div class="job-title"><%= rsJob("Title") %></div>
            <div class="job-given">Assigned to: <%= rsJob("AssignedName") %></div>
            <div class="target-date">
                Current Target Date: <strong><%= sys_date(rsJob("TargetDate")) %></strong>
            </div>
            <% If iso_date(currentDate) < iso_date(targetDate) Then %>
            <div style="color: #dc2626; font-weight: 600; margin: 1rem 0;">
                &#9888; Date locked until target date arrives
            </div>
            <% End If %>
        </div>

        <% If canChange Then %>
        <form method="post">
            <input type="hidden" name="action" value="change">
            <div style="margin: 1.5rem 0;">
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 600;">New Target Date</label>
                <input type=hidden name=oldTargetDate value="<%= sys_date(targetDate) %>">
                <input type="text" name="newTargetDate" id="newTargetDate" required value="<%= sys_date(targetDate) %>" style="width: 100%; padding: 12px; border-radius: 8px; border: 1px solid #d1d5db;">
            </div>
            <div style="text-align: center;">
                <button type="submit" style="background: #3b82f6; color: white; padding: 14px 40px; border: none; border-radius: 8px; font-size: 18px; cursor: pointer;">
                    Save New Date
                </button>
            </div>
        </form>
        <% Else %>
        <div style="text-align: center; color: #9ca3af; margin: 2rem;">
            You cannot change the date at this time.
        </div>
        <% End If %>
    </div>
<script language=javascript>
    $("#newTargetDate").datepicker();
    $("#newTargetDate").datepicker("option", "dateFormat", "d/m/yy");
</script>
    <!-- Switch User -->
    <div class="switch-user">
        Switch User (Demo): 
        <strong><%= Request.Cookies("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% rsJob.Close : Call CloseConnection() %>
</body>
</html>