<%@ Language="VBScript" CODEPAGE="874" LCID="1033" %>
<% Response.CharSet = "windows-874" : Response.CodePage = 874 %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->

<%
' ============================================================
' edit_job_action.asp
' รับค่า POST เพื่ออัปเดต Title และ TargetDate สำหรับงานที่ยัง Pending
' ============================================================

Dim jobID      : jobID      = Request.Form("JobID")
Dim newTitle   : newTitle   = Trim(Request.Form("Title"))
Dim newTarget  : newTarget  = Trim(Request.Form("TargetDate"))

If jobID <> "" And newTitle <> "" And newTarget <> "" Then
    ' ป้องกัน Single Quote ทำให้ SQL พัง
    Dim safeTitle : safeTitle = Replace(newTitle, "'", "''")

    ' อัปเดตข้อมูล โดยเช็คให้แน่ใจว่าเป็นงานของผู้มอบหมายคนนี้จริงๆ และยัง Pending อยู่
    Dim sqlUpdate
    sqlUpdate = "UPDATE " & DB_SCHEMA & "Jobs " & _
                "SET Title = N'" & safeTitle & "', " & _
                "    TargetDate = '" & newTarget & "', " & _
                "    UpdatedAt = CAST(GETDATE() AS DATE), " & _
                "    UpdatedBy = " & curRealUserID & " " & _
                "WHERE JobID = " & CLng(jobID) & " " & _
                "  AND AssignedBy = " & curRealUserID & " " & _
                "  AND Status = 'Pending'"
                
    objConn.Execute(sqlUpdate)
End If

' เนื่องจากปกติผู้จัดการจะเข้ามาแก้งานจากหน้ารวม (เช่น dashboard) เมื่อแก้เสร็จให้เด้งกลับไปที่เดิม
' ในที่นี้ขอกำหนดเป็น dashboard.asp ไว้ก่อน สามารถปรับเปลี่ยนได้ตาม Workflow จริงครับ
Response.Redirect "dashboard.asp"
%>