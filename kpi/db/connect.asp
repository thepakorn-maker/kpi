<% 
' =============================================
' db/connect.asp - Database Connection for Vinova KPI System
' =============================================

Dim conn

' ---------- แก้ไขส่วนนี้ตาม server ของคุณ ----------
Dim ConnectionString
ConnectionString = "Provider=SQLOLEDB;Data Source=.;" & _
                   "Initial Catalog=vndb;" & _
                   "User ID=vn;" & _
                   "Password=V13579$++"

' ถ้าใช้ Windows Authentication (แนะนำสำหรับภายในบริษัท)
' ConnectionString = "Provider=SQLOLEDB;Data Source=YOUR_SERVER_NAME;" & _
'                    "Initial Catalog=VinovaKPI;Integrated Security=SSPI;"

' ---------- สร้าง Connection ----------
On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open ConnectionString

If Err.Number <> 0 Then
    Response.Write "<h3 style='color:red;'>Database Connection Error: " & Err.Description & "</h3>"
    Response.End
End If
On Error GoTo 0

' ฟังก์ชันช่วยปิด Connection (เรียกตอนท้ายทุกหน้า)
Sub CloseConnection()
    If IsObject(conn) Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub
%>