<%
' ping.asp
' This file exists only to keep the session alive.
Session("LastPing") = Now() 
Response.Write "Active"
%>