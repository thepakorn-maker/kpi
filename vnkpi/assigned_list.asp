<%@ Language="VBScript" CODEPAGE="874" LCID="1033" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
Dim activePage : activePage = "assigned"
%>
<!-- #include file="header.asp" -->
<%
' -- Helper: job type pill -------------------------------------
Function jtPill(typeID, typeName)
    Dim cls : cls = "jt-" & Replace(CStr(typeID), ".", "")
    jtPill = "<span class=""" & cls & """>" & Server.HTMLEncode(CStr(typeName)) & "</span>"
End Function

Dim rs, sql

' Query jobs assigned BY current user with status='Pending'
sql = "SELECT j.*, u.Username as ReceiverName, j.TypeID AS JobTypeID, jt.TypeName " & _
      "FROM " & DB_SCHEMA & "Jobs j " & _
      "LEFT JOIN " & DB_SCHEMA & "Users u ON j.AssignedTo = u.UserID " & _
      "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
      "WHERE j.AssignedBy = " & curRealUserID & " AND j.Status = 'Pending' " & _
      "ORDER BY j.TargetDate ASC"

Set rs = objConn.Execute(sql)
%>

<div class="card">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
        <h2 style="font-size:28px;">&#x1F4D1; My Assigned Jobs (Pending)</h2>
        <p style="color:#6b7280;">Track all jobs you've assigned that are still pending</p>
    </div>

    <%
    If rs.EOF Then
    %>
        <div style="text-align:center;padding:60px">
            <div style="font-size:60px;margin-bottom:15px">&#10003;</div>
            <p style="font-size:18px;color:#6b7280;">No pending assigned jobs</p>
            <a href="assign_job.asp" class="btn btn-primary" style="margin-top:15px;display:inline-block;">+ Create New Job</a>
        </div>
    <%
    Else
    %>
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Job Type</th>
                    <th>Title</th>
                    <th>Assigned To</th>
                    <th>Target Date</th>

                    <th>Status</th>
                    <th style="text-align:center;">Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                Do While Not rs.EOF
                    Dim jID     : jID     = CLng(rs("JobID"))
                    Dim jType   : jType   = Server.HTMLEncode(CStr(rs("TypeName")))
                    Dim isRejected : isRejected = Not IsNull(rs("RejectComment"))
                    Dim targetDate : targetDate = rs("TargetDate")
                    Dim isOverdue  : isOverdue = (CDate(targetDate) < Date)
                    Dim rowClass : rowClass = ""
                    If isRejected Then rowClass = "rejected-row"
                    If isOverdue Then rowClass = "overdue-row"
                %>
                <tr class="<%= rowClass %>">
                    <td><strong>#<%= jID %></strong></td>   
                    <td><%= jtPill(rs("JobTypeID"), rs("TypeName")) %></td>
                    <td>
                        <div style="font-weight:600;"><%= safeStr(rs("Title")) %></div>
                        <small style="color:#6b7280;">Type: <%= rs("TypeID") %></small>
                        <% If isRejected Then %>
                            <div style="margin-top:5px;padding:5px 10px;background:#fecaca;color:#991b1b;border-radius:4px;font-size:12px;">
                                <strong>Reason for Rejection:</strong> <%= safeStr(rs("RejectComment")) %>
                            </div>
                        <% End If %>
                    </td>
                    <td>
                        <strong><%= safeStr(rs("ReceiverName")) %></strong>
                    </td>
                    <td>
                        <span style="font-weight:600;<%= iif(isOverdue, "color:#dc2626;", "") %>">
                            <%= fmtDate(targetDate) %>
                        </span>
                        <% If isOverdue Then %><br><small style="color:#dc2626;">(Overdue)</small><% End If %>
                    </td>


                    <td>
                        <% If isRejected Then %>
                            <span style="background:#ef4444;color:white;padding:2px 8px;border-radius:12px;font-size:12px;">Rejected</span>
                        <% Else %>
                            <span style="background:#3b82f6;color:white;padding:2px 8px;border-radius:12px;font-size:12px;">Pending</span>
                        <% End If %>
                    </td>
                    <td style="text-align:center;">
                        <a href="edit_job.asp?JobID=<%= rs("JobID") %>" class="btn btn-secondary" style="padding:5px 12px;font-size:13px;text-decoration:none;">
                            &#x270F; Edit
                        </a>
                    </td>
                </tr>
                <% rs.MoveNext : Loop %>
            </tbody>
        </table>
    <%
    End If
    rs.Close : Set rs = Nothing
    %>
</div>

</body>
</html>
<% objConn.Close : Set objConn = Nothing %>