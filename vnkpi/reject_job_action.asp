<%@ Language="VBScript" CODEPAGE="874" LCID="1033" %>
<% Response.CharSet = "windows-874" : Response.CodePage = 874 %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->

<%
' ============================================================
' reject_job_action.asp
' รับค่า POST (JobID, RejectComment) เพื่อตีกลับงานให้พนักงานแก้ไข
' ============================================================

Dim jobID         : jobID = Request.Form("JobID")
Dim rejectComment : rejectComment = Trim(Request.Form("RejectComment"))

If jobID <> "" And rejectComment <> "" Then
    ' Escape single quotes เพื่อป้องกัน SQL Error
    Dim safeComment : safeComment = Replace(rejectComment, "'", "''")

    ' อัปเดตสถานะกลับเป็น Pending และบันทึก Reject Comment
    Dim sqlReject
    sqlReject = "UPDATE " & DB_SCHEMA & "Jobs " & _
                "SET Status = 'Pending', " & _
                "    RejectDate = CAST(GETDATE() AS DATE), " & _
                "    RejectComment = N'" & safeComment & "' " & _
                "WHERE JobID = " & CLng(jobID) & " AND AssignedBy = " & curRealUserID
                
    objConn.Execute(sqlReject)
End If

' กลับไปหน้า Approve List
Response.Redirect "approve_list.asp"
%>