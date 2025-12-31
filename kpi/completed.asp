<%@ Language=VBScript %>

<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Job Completed - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
<%
UserID=Session("UserID")
%>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <%
    ' ฟังก์ชันแสดงวันที่ dd/mm/yyyy (ซ้ำไว้ในไฟล์นี้เพื่อไม่ error)
    Function FormatThaiDate(d)
        If IsDate(d) Then
            FormatThaiDate = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
        Else
            FormatThaiDate = "-"
        End If
    End Function
    %>

    <div style="padding: 1rem;">
        <h1 style="font-size: 24px; font-weight: 700; text-align: center; margin: 1.5rem 0;">
            &#10004; Job Completed
        </h1>
        <p style="text-align: center; color: #6b7280; margin-bottom: 2rem;">
            All completed jobs organized by person with KPI scores noted
        </p>

        <%
        ' นับ Total Completed Jobs ทั้งหมด
        Dim totalCompleted : totalCompleted = 0
        Dim sqlTotal : sqlTotal = "SELECT COUNT(*) AS Total FROM Jobs WHERE Status = 'Completed'"

        Dim rsTotal : Set rsTotal = conn.Execute(sqlTotal)
        If Not rsTotal.EOF Then totalCompleted = rsTotal("Total")
        rsTotal.Close
        %>

        <p style="text-align: center; font-size: 18px; font-weight: 600; margin-bottom: 2rem;">
            Total Completed Jobs: <%= totalCompleted %>
        </p>

        <%
        Dim sqlUsers, rsUsers
        sqlUsers = "SELECT UserID, Name, Department, KPIScore FROM Users "
        if Session("UserID")<>"" then
        sqlUsers =  sqlUsers & " where UserID='" & UserID & "'"
        end if

        sqlUsers =  sqlUsers & " ORDER BY Name"
        Set rsUsers = conn.Execute(sqlUsers)

        Do While Not rsUsers.EOF
            Dim userID : userID = rsUsers("UserID")

            ' นับสถิติงานเสร็จของคนนี้
            Dim completedCount : completedCount = 0
            Dim onTimeCount : onTimeCount = 0
            Dim lateCount : lateCount = 0

            Dim sqlCount : sqlCount = "SELECT " & _
                "COUNT(*) AS Total, " & _
                "SUM(CASE WHEN CompletedDate <= TargetDate THEN 1 ELSE 0 END) AS OnTime, " & _
                "SUM(CASE WHEN CompletedDate > TargetDate THEN 1 ELSE 0 END) AS Late " & _
                "FROM Jobs WHERE AssignedTo = " & userID & " AND Status = 'Completed'"

            Dim rsCount : Set rsCount = conn.Execute(sqlCount)
            If Not rsCount.EOF Then
                completedCount = rsCount("Total")
                onTimeCount = rsCount("OnTime")
                lateCount = rsCount("Late")
            End If
            rsCount.Close

            If completedCount > 0 Then
        %>
        <div class="user-card" style="background: #f0fdf4; margin: 1.5rem;">
            <div class="user-card-header">
                <div>
                    <h2><%= rsUsers("Name") %></h2>
                    <div class="position"><%= rsUsers("Department") %></div>
                </div>
                <div class="kpi-score" style="color: #16a34a;"><%= rsUsers("KPIScore") %></div>
            </div>

            <div class="stats">
                <div class="stat-box">
                    <div class="label">Total Completed</div>
                    <div class="number"><%= completedCount %></div>
                </div>
                <div class="stat-box" style="background: #f0fdf0;border:1px solid #ddddee;border-radius:2px;">
                    <div class="label">On Time</div>
                    <div class="number" style="color: #16a34a;"><%= onTimeCount %></div>
                </div>
                <% If lateCount > 0 Then %>
                <div class="stat-box overdue" style="background: #fdf0f0;border:1px solid #eedddd;border-radius:2px;">
                    <div class="label">Completed Late</div>
                    <div class="number"><%= lateCount %></div>
                </div>
                <% End If %>
            </div>

            <%
            ' ดึงงานที่เสร็จแล้วของคนนี้
            Dim sqlJobs : sqlJobs = "SELECT j.*, jt.PenaltyPerDay, jt.typename," & _
                "DATEDIFF(day, j.TargetDate, j.CompletedDate) AS RawDaysLate " & _
                "FROM Jobs j INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
                "WHERE j.AssignedTo = " & userID & " AND j.Status = 'Completed' " & _
                "ORDER BY j.CompletedDate DESC"

            Dim rsJobs : Set rsJobs = conn.Execute(sqlJobs)
            Dim jobNum : jobNum = 0

            Do While Not rsJobs.EOF
                jobNum = jobNum + 1

                ' คำนวณวันล่าช้า (ถ้าเสร็จก่อนหรือตรงเวลา -> 0)
                Dim daysLate : daysLate = rsJobs("RawDaysLate")
                If daysLate < 0 Then daysLate = 0

                ' คำนวณ KPI Impact (ถ้าไม่มีหัก -> 0)
                Dim penaltyPerDay : penaltyPerDay = rsJobs("PenaltyPerDay")
                Dim kpiImpact : kpiImpact = daysLate * penaltyPerDay
            %>
            <div class="job-card" style="border-left-color: <%= IIf(daysLate > 0, "#dc2626", "#22c55e") %>;">
                <div class="job-header">
                    <div class="job-id">#<%= jobNum %></div>
                    <div class="job-type" style="" ><%= "Type " & rsJobs("TypeID") & " (" & rsJobs("TypeName") & ")" %> <%= IIf(rsJobs("IsPrivate") = 1, "(Private)", "") %></div>
                    <div class="job-type" style="background:#fefebb" ><%= rsJobs("status") %> </div>

                    <div class="target-date">
                        Target: <%= sys_date(rsJobs("TargetDate")) %><br>
                        Completed: <%= sys_date(rsJobs("CompletedDate")) %>
                        <% If daysLate > 0 Then %>
                            <div class="overdue"><%= daysLate %> days late</div>
                        <% Else %>
                            <div style="color:#16a34a; font-weight:600;">On Time</div>
                        <% End If %>
                    </div>
                </div>
                <div class="job-title"><%= rsJobs("Title") %></div>
                <div class="job-given">Given by: <%= rsJobs("AssignedBy") %></div>
                <div class="job-footer" style="justify-content: flex-end;">
                    <div style="font-size: 18px; font-weight: 700; color: <%= IIf(kpiImpact < 0, "#dc2626", "#16a34a") %>;">
                        KPI Impact: <%= kpiImpact %> pts
                    </div>
                </div>
            </div>
            <%
                rsJobs.MoveNext
            Loop
            rsJobs.Close
            %>
        </div>
        <%
            End If
            rsUsers.MoveNext
        Loop
        rsUsers.Close
        %>
    </div>

    <!-- Switch User -->
    <div class="switch-user">
        Switch User (Demo): 
        <strong><%= Session("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% Call CloseConnection() %>
</body>
</html>