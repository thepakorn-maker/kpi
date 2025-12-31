<%@ Language=VBScript %>
<%
Response.Cookies("UserID") = ""
'Response.Cookies("UserID").Expires = DateAdd("d", -1, Now())

Response.Cookies("UserName") = ""
'Response.Cookies("UserName").Expires = DateAdd("d", -1, Now())

Response.Cookies("IsManager") = ""
'Response.Cookies("IsManager").Expires = DateAdd("d", -1, Now())

Response.Cookies("KPIScore") = ""
'Response.Cookies("KPIScore").Expires = DateAdd("d", -1, Now())

Response.Redirect "login.asp"
%>