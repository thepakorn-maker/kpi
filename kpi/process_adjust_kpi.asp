<%@ Language=VBScript %>

<!--#include file="includes/include.asp" -->
<!--#include file="db/connect.asp" -->
<%
Response.ContentType = "application/json"

Dim userID, changeAmt, success, newScore, reason
success = False
newScore = 0

If request.cookies("IsManager") And Request.Form("userID") <> "" And Request.Form("change") <> "" Then
    userID = CLng(Request.Form("userID"))
    changeAmt = CInt(Request.Form("change"))  ' +1 หรือ -1
    reason = IIf(changeAmt = 1, "Meeting +1 Point", "Meeting -1 Point")

    ' อัปเดต KPI
    Dim sqlUpdate : sqlUpdate = "UPDATE Users SET KPIScore = KPIScore + " & changeAmt & " WHERE UserID = " & userID
    conn.Execute sqlUpdate

    ' ดึงคะแนนใหม่
    Dim rsScore : Set rsScore = conn.Execute("SELECT KPIScore FROM Users WHERE UserID = " & userID)
    If Not rsScore.EOF Then newScore = rsScore("KPIScore")
    rsScore.Close

    ' บันทึก log
    conn.Execute "INSERT INTO KPILog (UserID, ChangeAmount, Reason) VALUES (" & userID & ", " & changeAmt & ", '" & reason & "')"

    success = True
End If

' ส่ง JSON กลับ
Response.Write "{ ""success"": " & LCase(success) & ", ""newScore"": " & newScore & " }"
Call CloseConnection()
%>