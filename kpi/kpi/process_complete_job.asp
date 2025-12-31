<%@ Language=VBScript %>

<!--#include file="db/connect.asp" -->
<%
UserID=request.cookies("UserID")
If request.cookies("UserID") = "" Then Response.Redirect "login.asp"

Dim jobID, returnPage
jobID = CLng(Request.Form("jobID"))
returnPage = Request.Form("returnPage")
If returnPage = "" Then returnPage = "dashboard"

' ตรวจสอบว่างานนี้เป็นของ user ปัจจุบันและยัง Pending
Dim sqlCheck : sqlCheck = "SELECT AssignedTo FROM Jobs WHERE JobID = " & jobID & " AND Status = 'Pending'"
Dim rsCheck : Set rsCheck = conn.Execute(sqlCheck)

If Not rsCheck.EOF Then
    If CLng(rsCheck("AssignedTo")) = CLng(request.cookies("UserID")) Then
        conn.Execute "UPDATE Jobs SET Status = 'Completed', CompletedDate = GETDATE() WHERE JobID = " & jobID
        
        ' บันทึก log (optional)
        conn.Execute "INSERT INTO KPILog (UserID, JobID, ChangeAmount, Reason) VALUES (" & request.cookies("UserID") & ", " & jobID & ", 0, 'Job completed')"
    End If
End If
rsCheck.Close
Set rsCheck = Nothing

Response.Redirect returnPage & ".asp"
%>