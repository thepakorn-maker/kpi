<%@ Language=VBScript %>

<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="windows-874">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Job WIP - Project Management Dashboard</title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body leftmargin=0 rightmargin=0>
    <!--#include file="includes/include.asp" -->
    <!--#include file="includes/header.asp" -->
    <!--#include file="includes/navbar.asp" -->

    <div astyle="padding: 1rem;">
        <h1 style="font-size: 24px; font-weight: 700; text-align: center; margin: 1rem 0;">
            Job WIP (Work in Progress)
        </h1>
        <p style="text-align: center; color: #6b7280; margin-bottom: 2rem;">
            All pending jobs organized by person, sorted by target date
        </p>

        <%
        Dim sqlUsers, rsUsers
        sqlUsers = "SELECT UserID, Name, Department, KPIScore FROM Users ORDER BY Name"
        Set rsUsers = conn.Execute(sqlUsers)

        Do While Not rsUsers.EOF
            Dim userID : userID = rsUsers("UserID")
            Dim pendingCount : pendingCount = 0
            Dim overdueCount : overdueCount = 0

            ' นับ Pending และ Overdue (ไม่รวม Private ของคนอื่น)
            Dim sqlCount
            sqlCount = "SELECT COUNT(*) AS TotalPending, " & _
                       "SUM(CASE WHEN TargetDate < GETDATE() THEN 1 ELSE 0 END) AS TotalOverdue " & _
                       "FROM Jobs WHERE AssignedTo = " & userID & " AND Status = 'Pending' " & _
                       "AND (IsPrivate = 0 OR AssignedBy = " & userID & ")"
            Dim rsCount : Set rsCount = conn.Execute(sqlCount)
            If Not rsCount.EOF Then
                pendingCount = rsCount("TotalPending")
                overdueCount = rsCount("TotalOverdue")
            End If
            rsCount.Close
        %>
        <div class="user-card" style="aborder:1px solid black;">
            <div class="user-card-header">
                <div>
                    <h2><%= rsUsers("Name") %></h2>
                    <div class="position"><%= rsUsers("Department") %></div>
                </div>
                <div class="kpi-score"><%= rsUsers("KPIScore") %></div>
            </div>

            <div class="stats">
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
            ' ดึงงาน Pending ของคนนี้
            Dim sqlJobs2, rsJobs2
            sqlJobs2 = "SELECT j.*, jt.TypeName, DATEDIFF(day, j.TargetDate, GETDATE()) AS DaysLate " & _
                       "FROM Jobs j INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
                       "WHERE j.AssignedTo = " & userID & " AND j.Status = 'Pending' " & _
                       "AND (j.IsPrivate = 0 OR j.AssignedBy = " & userID & ") " & _
                       "ORDER BY j.TargetDate ASC"
                       'response.write sqlJobs2
                       'response.end
            Set rsJobs2 = conn.Execute(sqlJobs2)
            on error resume next
            Dim jobNum : jobNum = 0
            Do While Not rsJobs2.EOF
                jobNum = jobNum + 1
                Dim isToday2 : isToday2 = (sys_date(rsJobs2("TargetDate")) = sys_date(Date))
            %>
            <div class="job-card type<%= Replace(rsJobs2("TypeID"),".","") %> <%= IIf(isToday2, "blink-today", "") %>">
                <div class="job-header">
                    <div class="job-id">#<%= jobNum %></div>
                    <div class="job-type"><%= "Type " & rsJobs2("TypeID") & " (" & rsJobs2("TypeName") & ")" %> <%= IIf(rsJobs2("IsPrivate"), "(Private)", "") %></div>
                    <div class="job-type" style="background:#fefebb"><%=rsJobs2("status")%></div>
                    <div class="target-date">
                        Target Date<br>
                        <strong><%'= rsJobs2("TargetDate")%><%= sys_date(cdate(rsJobs2("TargetDate"))) %></strong>
                        <% If rsJobs2("DaysLate") > 0 Then %>
                            <div class="overdue"><%= rsJobs2("DaysLate") %>d overdue</div>
                        <% ElseIf isToday2 Then %>
                            <div class="due-today">Due TODAY!</div>
                        <% End If %>
                    </div>
                </div>
                <div class="job-title"><%= rsJobs2("Title") %></div>
                <div class="job-given">Given by: <%= rsJobs2("AssignedBy") %></div>
                <div class="job-footer">
                <% If rsJobs2("AssignedTo") = Request.Cookies("UserID") Then %>
                <form method="post" action="process_complete_job.asp" style="display: inline;">
                <input type="hidden" name="jobID" value="<%= rsJobs2("JobID") %>">
                <input type="hidden" name="returnPage" value="alljobs">
                <button type="submit" class="btn-complete">&#10004; Complete Job</button>
               </form>
               <% End If %>
</div>

            </div>
            <%
                rsJobs2.MoveNext
            Loop
            rsJobs2.Close
            %>
        </div>
        <%
            rsUsers.MoveNext
        Loop
        rsUsers.Close
        %>
    </div>

    <!-- Switch User -->
    <div class="switch-user">
        Switch User : 
        <strong><%= Session("UserName") %> - <%= currentDepartment %></strong>
    </div>

    <% Call CloseConnection() %>
</body>
</html>