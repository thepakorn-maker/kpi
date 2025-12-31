<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="db/connect.asp" -->
<%
response.charset="windows-874"

Dim loadUserID : loadUserID = CLng(Request.QueryString("userID"))

Dim pendingCount : pendingCount = 0
Dim sqlLoad : sqlLoad = "SELECT COUNT(*) AS Total FROM Jobs " & _
    "WHERE AssignedTo = " & loadUserID & " AND Status = 'Pending' " & _
    "AND (IsPrivate = 0 OR AssignedBy = " & loadUserID & ")"  ' รวม Private ของตัวเอง

Dim rsLoad : Set rsLoad = conn.Execute(sqlLoad)
If Not rsLoad.EOF Then pendingCount = rsLoad("Total")
rsLoad.Close
Set rsLoad = Nothing

Response.Write "<strong>จำนวนงาน Pending ปัจจุบัน:</strong> " & pendingCount & " งาน<br>"
If pendingCount >= 5 Then
    Response.Write "<span style='color:#dc2626;'> งานค้างเยอะ – พิจารณามอบหมายให้คนอื่น</span>"
ElseIf pendingCount >= 3 Then
    Response.Write "<span style='color:#f59e0b;'> งานค้างปานกลาง</span>"
Else
    Response.Write "<span style='color:#16a34a;'>พร้อมรับงานเพิ่ม</span>"
End If

Call CloseConnection()
%>