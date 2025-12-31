<%@ Language=VBScript %>

<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->
<%
response.charset="windows-874"


Dim loadUserID : loadUserID = CLng(Request.QueryString("userID"))

Dim pendingCount : pendingCount = 0
Dim sqlCount : sqlCount = "SELECT COUNT(*) AS Total FROM Jobs " & _
    "WHERE AssignedTo = " & loadUserID & " AND Status = 'Pending' " & _
    "AND (IsPrivate = 0 OR AssignedBy = " & loadUserID & ")"

Dim rsCount : Set rsCount = conn.Execute(sqlCount)
If Not rsCount.EOF Then pendingCount = rsCount("Total")
rsCount.Close

' พื้นหลังหลักสีเหลืองอ่อน (เหมือนภาพ)
Response.Write "<div style=""background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:1rem;"">"
Response.Write "<div style=""font-size:18px; font-weight:700; color:#92400e; margin-bottom:1rem;"">"
Response.Write "Existing Jobs:"
Response.Write "</div>"

If pendingCount = 0 Then
    Response.Write "<div style=""text-align:center; color:#92400e; padding:2rem;"">"
    Response.Write "No pending jobs"
    Response.Write "</div>"
Else
    Dim sqlJobs : sqlJobs = "SELECT j.JobID, j.Title, j.TypeID, j.TargetDate, j.ShiftsUsed, jt.TypeName, jt.ShiftsAllowed, " & _
        "u.Name AS GivenByName, DATEDIFF(day, j.TargetDate, GETDATE()) AS DaysLate " & _
        "FROM Jobs j " & _
        "INNER JOIN JobTypes jt ON j.TypeID = jt.TypeID " & _
        "LEFT JOIN Users u ON j.AssignedBy = u.UserID " & _
        "WHERE j.AssignedTo = " & loadUserID & " AND j.Status = 'Pending' " & _
        "AND (j.IsPrivate = 0 OR j.AssignedBy = " & loadUserID & ") " & _
        "ORDER BY j.TargetDate ASC"
    
    Dim rsJobs : Set rsJobs = conn.Execute(sqlJobs)
    
    Do While Not rsJobs.EOF
        Dim daysLate : daysLate = rsJobs("DaysLate")
        Dim isOverdue : isOverdue = (daysLate > 0)
        Dim shiftsText
        If IsNull(rsJobs("ShiftsAllowed")) Then
            shiftsText = "Unlimited"
        Else
            shiftsText = rsJobs("ShiftsUsed") & "/" & rsJobs("ShiftsAllowed")
        End If
        
        Response.Write "<div style=""background:white; margin-bottom:0.8rem; padding:1rem; border-radius:8px; border-left:4px solid " & _
            IIf(isOverdue, "#dc2626", "#3b82f6") & "; box-shadow:0 1px 3px rgba(0,0,0,0.1);"">"
        
        Response.Write "<div style=""display:flex; justify-content:space-between; align-items:flex-start;"">"
        Response.Write "<div style=""flex:1;"">"
        Response.Write "<div style=""font-weight:700; color:#1f2937; margin-bottom:0.3rem;"">" & rsJobs("Title") & "</div>"
        Response.Write "<div style=""font-size:14px; color:#6b7280;"">Given by: " & rsJobs("GivenByName") & "</div>"
        If isOverdue Then
            Response.Write "<div style=""font-size:14px; color:#dc2626; font-weight:700; margin-top:0.3rem;"">"
            Response.Write "Due: " & sys_date(rsJobs("TargetDate")) & " (OVERDUE!)"
            Response.Write "</div>"
        Else
            Response.Write "<div style=""font-size:14px; color:#1e3a8a; margin-top:0.3rem;"">"
            Response.Write "Due: " & sys_date(rsJobs("TargetDate"))
            Response.Write "</div>"
        End If
        Response.Write "</div>"
        
        Response.Write "<div style=""margin-left:1rem; text-align:right;"">"
        Response.Write "<span style=""background:#fbbf24; color:#92400e; padding:0.3rem 0.8rem; border-radius:9999px; font-size:13px; font-weight:600;"">"
        Response.Write "Type " & rsJobs("TypeID")
        Response.Write "</span><br>"
        Response.Write "<span style=""font-size:13px; color:#6b7280; margin-top:0.5rem; display:block;"">"
        Response.Write "Shifts: " & shiftsText
        Response.Write "</span>"
        Response.Write "</div>"
        Response.Write "</div>"
        Response.Write "</div>"
        
        rsJobs.MoveNext
    Loop
    rsJobs.Close
End If

Response.Write "</div>"

Call CloseConnection()
%>