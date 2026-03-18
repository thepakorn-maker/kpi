<%@ Language="VBScript" %>
<%
' ============================================================
' logout.asp — Destroy session and redirect to login
' ============================================================
Session.Abandon
Response.Redirect "login.asp"
Response.End
%>
