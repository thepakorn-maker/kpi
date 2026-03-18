<%@ language="VBScript" %>
<%
  Option Explicit

  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<HTML><HEAD><TITLE>The page cannot be displayed</TITLE>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=Windows-874">
<STYLE type="text/css">
  BODY { font: 8pt/12pt tahoma }
  H1 { font: 13pt/15pt tahoma }
  H2 { font: 8pt/12pt tahoma }
  A:link { color: red }
  A:visited { color: maroon }
</STYLE>
</HEAD><BODY><TABLE width=500 border=0 cellspacing=10><TR><TD>

<font size=+1 face="tahoma">The page cannot be displayed</font>
<hr>
Technical Information (for support personnel)

<ul>
<li><font style="font-size:15px"><b>ข้อความผิดพลาดมีดังนี้:</b><br>
<%
'dim con, sql
'set con=Server.CreateObject("ADODB.Connection")
'con.Open "Provider=sqlncli11;server=(local);database=vndb;uid=sa;pwd=Vinova9465##$$"
  Dim bakCodepage
  on error resume next
	bakCodepage = Session.Codepage
	Session.Codepage = 1252
  on error goto 0
  dim err_string
  err_string=""
  response.write "<FONT style='font-family:tahoma;font-size:14px;' color=red>"
  Response.Write Server.HTMLEncode(objASPError.Category)
  err_string=err_string & Server.HTMLEncode(objASPError.Category)
  if objASPError.ASPCode > "" Then 
  Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
  err_string=err_string & Server.HTMLEncode(", " & objASPError.ASPCode)
  end if
  Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
  err_string=err_string & Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & chr(13) & chr(10)
  
  If objASPError.ASPDescription > "" Then 
    Response.Write Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
    err_string=err_string &  Server.HTMLEncode(objASPError.ASPDescription) & chr(13) & chr(10) 
  ElseIf (objASPError.Description > "") Then 
	Response.Write Server.HTMLEncode(objASPError.Description) & "<br>" 
	err_string=err_string & Server.HTMLEncode(objASPError.Description) & chr(13) & chr(10)
  end if
  blnErrorWritten = False
  ' Only show the Source if it is available and the request is from the same machine as IIS
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
    If (strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      Response.Write Server.HTMLEncode(objASPError.File)
      err_string=err_string & Server.HTMLEncode(objASPError.File)
      if objASPError.Line > 0 Then 
      Response.Write ", line " & objASPError.Line
      err_string=err_string & ", line " & objASPError.Line
      end if
      If objASPError.Column > 0 Then 
      Response.Write ", column " & objASPError.Column
      err_string=err_string & ", column " & objASPError.Column
      end if
      Response.Write "<br>"
      err_string=err_string & chr(13) & chr(10)
      Response.Write "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
      'err_string=err_string & "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
      err_string=err_string & Server.HTMLEncode(objASPError.Source) & chr(13) & chr(10)
      If objASPError.Column > 0 Then 
      Response.Write String((objASPError.Column - 1), "-") & "^<br>"
      err_string=err_string & String((objASPError.Column - 1), "-") & "^" & chr(13) & chr(10)
      end if
      Response.Write "</b></font>"
      blnErrorWritten = True
    End If
  End If
  If Not blnErrorWritten And objASPError.File <> "?" Then
    Response.Write "<b>" & Server.HTMLEncode(objASPError.File)
    err_string=err_string & Server.HTMLEncode(objASPError.File)
    if objASPError.Line > 0 Then 
    Response.Write Server.HTMLEncode(", line " & objASPError.Line)
    err_string=err_string & Server.HTMLEncode(", line " & objASPError.Line)
    end if
    if objASPError.Column > 0 Then 
    Response.Write ", column " & objASPError.Column
    err_string=err_string & ", column " & objASPError.Column
    end if
    Response.Write "</b><br>"
  End If
  response.write "</font>"
  err_string=replace(err_string,"'","''")
  err_string=replace(err_string,chr(34),"""""")
  'Response.Write err_string
  'sql="INSERT into ERR_TAB(err_desc,session_id,user_code) values('" & err_string & "','" & Session.SessionID & "','" & request.cookies("user_code") & "')"
  'con.Execute(sql)
  'con.Close
  'set con=nothing
%>
</font>
</li>
<BR><BR>
<li>Browser Type:<br>
<%= Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT")) %>
<br><br></li>
<li>Page:<br>
<%
  strMethod = Request.ServerVariables("REQUEST_METHOD")
  Response.Write strMethod & " "
  If strMethod = "POST" Then
    Response.Write Request.TotalBytes & " bytes to "
  End If
  Response.Write Request.ServerVariables("SCRIPT_NAME")
  Response.Write "</li>"
  If strMethod = "POST" Then
    Response.Write "<p><li>POST Data:<br>"
    ' On Error in case Request.BinaryRead was executed in the page that triggered the error.
    On Error Resume Next
    If Request.TotalBytes > lngMaxFormBytes Then
      Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
    Else
      Response.Write Server.HTMLEncode(Request.Form)
    End If
    On Error Goto 0
    Response.Write "</li>"
  End If
%>
<br><br></li>
<li>Time:<br>
<%
  datNow = Now()
  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
  on error resume next
	Session.Codepage = bakCodepage 
  on error goto 0
%>
<br><br></li>
<!--comment>
<li>More information:<br>
<%  
  strQueryString = "prd=iis&sbp=&pver=5.0&ID=500;100&cat=" & Server.URLEncode(objASPError.Category) & "&os=&over=&hrd=&Opt1=" & Server.URLEncode(objASPError.ASPCode)  & "&Opt2=" & Server.URLEncode(objASPError.Number) & "&Opt3=" & Server.URLEncode(objASPError.Description) 
  strURL = "http://www.microsoft.com/ContentRedirect.asp?" & strQueryString
%>
  <ul>
  <li>Click on <a href="<%= strURL %>">Microsoft Support</a> for a links to articles about this error.</li>
  <li>Go to <a href="http://go.microsoft.com/fwlink/?linkid=8180" target="_blank">Microsoft Product Support Services</a> and perform a title search for the words <b>HTTP</b> and <b>500</b>.</li>
  <li>Open <b>IIS Help</b>, which is accessible in IIS Manager (inetmgr), and search for topics titled <b>Web Site Administration</b>, and <b>About Custom Error Messages</b>.</li>
  <li>In the IIS Software Development Kit (SDK) or at the <a href="http://go.microsoft.com/fwlink/?LinkId=8181">MSDN Online Library</a>, search for topics titled <b>Debugging ASP Scripts</b>, <b>Debugging Components</b>, and <b>Debugging ISAPI Extensions and Filters</b>.</li>
  </ul>
</li>
</comment-->
</ul>

</TD></TR></TABLE></BODY></HTML>