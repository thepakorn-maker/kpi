<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
' ============================================================
' workload_ajax.asp — AJAX endpoint for assign_job.asp
' Returns HTML snippet showing selected user's pending jobs.
' Called by showWorkload() JS function via XMLHttpRequest GET.
' URL: workload_ajax.asp?UserID=<int>
'
' Matches prototype showWorkload() display exactly:
'   - 0 jobs: green "Available!" alert
'   - N jobs: amber alert with workload table
'     Columns: Job title / Target Date (color-coded) / From / Type
' ============================================================

Response.ContentType = "text/html"
Response.Charset = "utf-8"

Dim targetID : targetID = CLng(Request.QueryString("UserID"))

' Validate — must be a real active user
If targetID = 0 Then Response.End

' Get username for heading
Dim rsUser
Set rsUser = objConn.Execute( _
    "SELECT Username, ISNULL(Dept, Department) AS Dept " & _
    "FROM " & DB_SCHEMA & "Users " & _
    "WHERE UserID = " & targetID & " AND IsActive = 1")

If rsUser.EOF Then
    rsUser.Close : Set rsUser = Nothing
    objConn.Close : Set objConn = Nothing
    Response.End
End If

Dim userName : userName = CStr(rsUser("Username"))
Dim userDept : userDept = CStr(rsUser("Dept"))
rsUser.Close : Set rsUser = Nothing

' Fetch pending non-private jobs for this user, sorted by target date
Dim rs
Set rs = objConn.Execute( _
    "SELECT j.Title, j.TargetDate, ab.Username AS AssignedByName, " & _
    "       jt.TypeName, j.TypeID " & _
    "FROM " & DB_SCHEMA & "Jobs j " & _
    "JOIN " & DB_SCHEMA & "JobTypes jt ON j.TypeID = jt.TypeID " & _
    "JOIN " & DB_SCHEMA & "Users ab    ON j.AssignedBy = ab.UserID " & _
    "WHERE j.AssignedTo = " & targetID & " " & _
    "  AND j.Status = 'Pending' " & _
    "  AND j.IsPrivate = 0 " & _
    "ORDER BY j.TargetDate ASC")

If rs.EOF Then
    ' No pending jobs — show green "Available" alert
    rs.Close : Set rs = Nothing
    objConn.Close : Set objConn = Nothing
    Response.Write "<div class=""alert"" style=""background:#ecfdf5;border-left:4px solid #059669;margin-bottom:20px"">"
    Response.Write "<strong>&#x2705; " & Server.HTMLEncode(userName) & " — Available!</strong><br>"
    Response.Write "No pending jobs — good time to assign work!"
    Response.Write "</div>"
    Response.End
End If

' Has jobs — show amber workload table
Dim jobCount : jobCount = 0
Dim rowsHTML : rowsHTML = ""

Do While Not rs.EOF
    Dim td     : td     = CDate(rs("TargetDate"))
    Dim dLeft  : dLeft  = DateDiff("d", Date(), td)
    Dim tdStr  : tdStr  = fmtDate(td)
    Dim tdColor, tdNote

    If dLeft < 0 Then
        tdColor = "#dc2626"
        tdNote  = "<br>&#x1F534; OVERDUE!"
    ElseIf dLeft = 0 Then
        tdColor = "#f59e0b"
        tdNote  = "<br>&#x1F525; DUE TODAY!"
    ElseIf dLeft <= 2 Then
        tdColor = "#f97316"
        tdNote  = "<br>&#x26A0; SOON"
    Else
        tdColor = "#059669"
        tdNote  = ""
    End If

    ' Job type pill CSS class
    Dim tCls : tCls = "jt-" & Replace(CStr(rs("TypeID")), ".", "")

    rowsHTML = rowsHTML & _
        "<tr style=""background:#fff;border-bottom:1px solid #e5e7eb"">" & _
        "<td style=""padding:8px"">" & Server.HTMLEncode(CStr(rs("Title"))) & "</td>" & _
        "<td style=""padding:8px;color:" & tdColor & ";font-weight:bold"">" & _
            tdStr & tdNote & "</td>" & _
        "<td style=""padding:8px"">" & Server.HTMLEncode(CStr(rs("AssignedByName"))) & "</td>" & _
        "<td style=""padding:8px""><span class=""" & tCls & """>" & _
            Server.HTMLEncode(CStr(rs("TypeName"))) & "</span></td>" & _
        "</tr>"

    jobCount = jobCount + 1
    rs.MoveNext
Loop
rs.Close : Set rs = Nothing
objConn.Close : Set objConn = Nothing

' Render amber workload panel
Response.Write "<div class=""alert"" style=""background:#fef3c7;border-left:4px solid #f59e0b;margin-bottom:20px"">"
Response.Write "<strong>&#x1F4CA; " & Server.HTMLEncode(userName) & "'s Current Workload: " & jobCount & " pending jobs</strong>"
Response.Write "<div style=""margin-top:15px;max-height:300px;overflow-y:auto"">"
Response.Write "<table style=""width:100%;font-size:14px"">"
Response.Write "<tr style=""background:#fbbf24;color:#000"">"
Response.Write "<th style=""padding:8px"">Job</th>"
Response.Write "<th style=""padding:8px"">Target Date</th>"
Response.Write "<th style=""padding:8px"">From</th>"
Response.Write "<th style=""padding:8px"">Type</th>"
Response.Write "</tr>"
Response.Write rowsHTML
Response.Write "</table></div>"
Response.Write "<div style=""margin-top:10px;font-size:13px;color:#92400e"">"
Response.Write "&#x1F4A1; <strong>Tip:</strong> Check existing target dates above to avoid overloading."
Response.Write "</div></div>"
%>
