<%@ Language=VBScript %>
<%
UserID=request.cookies("UserID")

%>
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <!--#include file="includes/include.asp" -->

    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <div class="user-card">
        <div class="user-card-header">
            <div>
                <h2><%= request.cookies("UserName") %></h2>
                <div class="position"><%= currentDepartment %></div>
            </div>
            <div class="kpi-score"><%= currentKPI %></div>
        </div>

        <div class="stats">
<%
        Dim pendingCount : pendingCount = 0
        Dim overdueCount : overdueCount = 0
        
        Dim sqlStats : sqlStats = "SELECT " & _
            "COUNT(*) AS TotalPending, " & _
            "SUM(CASE WHEN TargetDate < GETDATE() THEN 1 ELSE 0 END) AS TotalOverdue " & _
            "FROM Jobs " & _
            "WHERE AssignedTo = " & request.cookies("UserID") & " AND Status = 'Pending' " & _
            "AND (IsPrivate = 0 OR AssignedBy = " & request.cookies("UserID") & ")"
        
        Dim rsStats : Set rsStats = conn.Execute(sqlStats)
        If Not rsStats.EOF Then
            pendingCount = rsStats("TotalPending")
            overdueCount = rsStats("TotalOverdue")
        End If
        rsStats.Close
        Set rsStats = Nothing
        %>

        <div class="stat-box">
            <div class="label">Pending Jobs</div>
            <div class="number"><%= pendingCount %></div>
        </div>
        
        <% If overdueCount > 0 Then %>
        <div class="stat-box overdue">
            <div class="label">Overdue</div>
            <div class="number"><%= overdueCount %></div>
        </div>
        <% End If %>
        </div>

    <%
    ' ดึงงาน Pending ของ user ปัจจุบัน (รวม Private ถ้าเป็นของตัวเอง)
    Dim sqlJobs, rsJobs
    sqlJobs = "SELECT j.*, jt.TypeName, DATEDIFF(day, j.TargetDate, GETDATE()) AS DaysLate,jt.ShiftsAllowed " & _
              "FROM Jobs j INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
              "WHERE j.AssignedTo = '" & UserID & "' AND j.Status = 'Pending' " & _
              "AND (j.IsPrivate = 0 OR j.AssignedBy = '" & UserID & "') " & _
              "ORDER BY j.TargetDate ASC"
    Set rsJobs = conn.Execute(sqlJobs)
    
    Dim jobCount : jobCount = 0
    Do While Not rsJobs.EOF
        jobCount = jobCount + 1
        Dim isToday : isToday = (sys_date(rsJobs("TargetDate")) = sys_date(Date))
        TypeID=rsJobs("TypeID")
        ShiftsAllowed=rsJobs("ShiftsAllowed")
        ShiftsUsed=clng("0" & rsJobs("ShiftsUsed"))
    %>
        <div class="job-card type<%= Replace(rsJobs("TypeID"),".","") %> <%= IIf(isToday, "blink-today", "") %>">
            <div class="job-header">
                <div class="job-id">#<%= jobCount %></div>
                <div class="job-type"><%= rsJobs("TypeName") %> <%= IIf(rsJobs("IsPrivate"), "(Private)", "") %></div>
                <div class="target-date">
                    Target Date<br>
                    <strong><%= sys_date(rsJobs("TargetDate")) %></strong>
                    <% If rsJobs("DaysLate") > 0 Then %>
                        <div class="overdue"><%= rsJobs("DaysLate") %> days overdue</div>
                    <% ElseIf isToday Then %>
                        <div class="due-today">Due TODAY!</div>
                    <% End If %>
                </div>
            </div>
            <div class="job-title"><%= rsJobs("Title") %></div>
            <div class="job-given">Given by: <%= rsJobs("AssignedBy") %></div>
            <div class="job-footer">
                <!--button class="btn-complete">&#10004; Complete Job</button>
                <button class="btn-change <%'=IIf(isToday, "enabled", "") %>">Change Date</button-->
    <form method="post" action="process_complete_job.asp" style="display: inline;">
        <input type="hidden" name="jobID" value="<%= rsJobs("JobID") %>">
        <input type="hidden" name="returnPage" value="dashboard">
        <button type="submit" class="btn-complete">&#10004; Complete Job</button>
    </form>

    <% 
    ' ตรวจสอบว่าถึง Target Date แล้วหรือยัง เพื่อปลดล็อก Change Date
    Dim isTodayOrPast : isTodayOrPast = ( datediff("d",rsJobs("TargetDate"),Date())>=0 )
    %><%'=clng("0" & ShiftsAllowed) & "@" & ShiftsUsed%>
    <button class="btn-change <%= IIf(isTodayOrPast and TypeID<>"5" and (clng("0" & ShiftsAllowed)>ShiftsUsed or isnull(ShiftsAllowed)), "enabled", "") %>" 
            onclick="if(this.classList.contains('enabled')) { window.location='change_date.asp?jobID=<%= rsJobs("JobID") %>'; }">
        Change Date
    </button>               
        </div>
    </div>
    <%
        rsJobs.MoveNext
    Loop
    rsJobs.Close
    %>
    </div>  <!-- div class=user-card -->


<!-- ส่วนงานที่ตัวเองมอบหมาย (My Given Jobs) -->
<div class="user-card" style="margin-top: 2rem; background: #fffbeb;">
    <h2 style="font-size: 22px; font-weight: 700; margin-bottom: 1rem; color: #1e3a8a;">
        My Given Jobs (งานที่ฉันมอบหมาย)
    </h2>

    <%
    ' นับจำนวนงานที่ตัวเองมอบหมาย (Pending)
    Dim givenPending : givenPending = 0
    Dim sqlGivenCount : sqlGivenCount = "SELECT COUNT(*) AS Total FROM Jobs WHERE AssignedBy = " & request.cookies("UserID") & " AND Status = 'Pending'"
    Dim rsGivenCount : Set rsGivenCount = conn.Execute(sqlGivenCount)
    If Not rsGivenCount.EOF Then givenPending = rsGivenCount("Total")
    rsGivenCount.Close
    %>

    <div class="stat-box" style="margin-bottom: 1rem;">
        <div class="label">Total Given Pending</div>
        <div class="number"><%= givenPending %></div>
    </div>

    <%
    Dim sqlGiven, rsGiven
    sqlGiven = "SELECT j.*, u.Name AS AssigneeName, u.Department AS AssigneeDept, jt.TypeName " & _
               "FROM Jobs j " & _
               "INNER JOIN Users u ON j.AssignedTo = u.UserID " & _
               "INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
               "WHERE j.AssignedBy = " & request.cookies("UserID") & " AND j.Status = 'Pending' " & _
               "ORDER BY j.TargetDate ASC"
    
    Set rsGiven = conn.Execute(sqlGiven)
    
    Dim givenNum : givenNum = 0
    Do While Not rsGiven.EOF
        givenNum = givenNum + 1
        Dim isTodayGiven : isTodayGiven = (iso_date(rsGiven("TargetDate")) = iso_date(Date))
    %>
    <div class="job-card type<%= Replace(rsGiven("TypeID"),".","") %> <%= IIf(isTodayGiven, "blink-today", "") %>">
        <div class="job-header">
            <div class="job-id">#<%= givenNum %></div>
            <div class="job-type"><%= rsGiven("TypeName") %></div>
            <div class="target-date">
                Target Date<br>
                <strong><%= sys_date(rsGiven("TargetDate")) %></strong>
                <% If iso_date(rsGiven("TargetDate")) < iso_date(Date) Then %>
                    <div class="overdue"><%= DATEDIFF("d", iso_date(rsGiven("TargetDate")), iso_date(Date)) %>d overdue</div>
                <% ElseIf isTodayGiven Then %>
                    <div class="due-today">Due TODAY!</div>
                <% End If %>
            </div>
        </div>
        <div class="job-title"><%= rsGiven("Title") %></div>
        <div class="job-given">Given to: <%= rsGiven("AssigneeName") %> - <%= rsGiven("AssigneeDept") %></div>
    </div>
    <%
        rsGiven.MoveNext
    Loop
    rsGiven.Close
    Set rsGiven = Nothing
    %>

    <% If givenPending = 0 Then %>
    <div style="text-align: center; color: #9ca3af; padding: 2rem;">
        ยังไม่มีงานที่คุณมอบหมายให้คนอื่น (Pending)
    </div>
    <% End If %>
</div>


    <!-- Switch User -->
    <!--#include file="switch_user.asp"-->

    <!--div class="switch-user">
        Switch User : 
        <strong><%'= request.cookies("UserName") %> - <%= currentDepartment %></strong>
    </div-->

    <% Call CloseConnection() %>
</body>
</html>