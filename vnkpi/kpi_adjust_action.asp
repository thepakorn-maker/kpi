<%@ Language="VBScript" %>
<!-- #include file="conn.asp" -->
<!-- #include file="auth_check.asp" -->
<!-- #include file="functions.asp" -->
<%
If Not curIsManager Then Response.Redirect "dashboard.asp" : Response.End
' ============================================================
' kpi_adjust_action.asp — POST: apply quick KPI adjustment
' Called by the KPI Awards form in meeting.asp
' Calls sp_AddMeetingKPIAdjustment (no ProblemID — standalone award)
' ============================================================

Dim targetUserID : targetUserID = CLng(Request.Form("TargetUserID"))
Dim points       : points       = CLng(Request.Form("Points"))
Dim reason       : reason       = Trim(Request.Form("Reason"))

If targetUserID = 0 Then
    Response.Redirect "meeting.asp?err=" & Server.URLEncode("Please select a team member.")
    Response.End
End If
If points = 0 Then
    Response.Redirect "meeting.asp?err=" & Server.URLEncode("Adjustment is 0 points. Please select a point value.")
    Response.End
End If
If reason = "" Then
    Response.Redirect "meeting.asp?err=" & Server.URLEncode("Please enter a reason for the adjustment.")
    Response.End
End If
Dim refJobID : refJobID = 0
If Request.Form("RefJobID") <> "" Then refJobID = CLng(Request.Form("RefJobID"))

Dim sql
sql = "EXEC " & DB_SCHEMA & "sp_AddMeetingKPIAdjustment " & _
      "@TargetUserID=" & targetUserID                       & ", " & _
      "@Points="       & points                             & ", " & _
      "@Reason='"      & Replace(reason, "'", "''")        & "', " & _
      "@AdjustedBy="   & curRealUserID & ", " & _
      "@RefJobID="     & refJobID     ' <- เพิ่ม (0 = ไม่มี job เกี่ยวข้อง)

On Error Resume Next
objConn.Execute sql
Dim spError : spError = Err.Description
On Error GoTo 0
objConn.Close : Set objConn = Nothing

If spError <> "" Then
    Response.Redirect "meeting.asp?err=" & Server.URLEncode(spError)
    Response.End
End If

Response.Redirect "meeting.asp"
Response.End
%>
