<%
' ============================================================
' conn.asp — Database connection include
' Usage: <!-- #include file="conn.asp" -->
' Provides: objConn (open ADODB.Connection)
' ============================================================

Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")

Dim strConn
strConn = "Provider=SQLOLEDB;" & _
          "Data Source=(local);" & _
          "Initial Catalog=vndb;" & _
          "User ID=vn;" & _
          "Password=V13579$++;"

' Schema prefix constant — use vn.TableName in all SQL
Const DB_SCHEMA = "vn."

On Error Resume Next
objConn.Open strConn

If Err.Number <> 0 Then
    Response.Write "<h2 style='color:red;font-family:Arial'>Database connection failed</h2>"
    Response.Write "<p style='font-family:Arial'>Error: " & Err.Description & "</p>"
    Response.End
End If
On Error GoTo 0
%>
