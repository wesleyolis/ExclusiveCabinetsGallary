<%@ LANGUAGE = VBScript %>

<% Session.Timeout = 20 %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>

<%
if Request.Form("Action") = "Login" Then
if Session("Admin") <> true THEN
if Request.Form("user") = "excab" Then
if Request.Form("pass") = "Kathy" Then
Session("Admin") = true 
Session("msg") = "Sucessful Login"
Response.redirect "default.asp?Email=" + Request.QueryString("Email")
End IF
End IF
ELSE
Session("Admin") = true
Session("msg") = "Already Login"
Response.redirect "default.asp?Email=" + Request.QueryString("Email")
END IF
Else
if Request.Form("Action") = "Logout" then
Session("Admin") = false
Session("msg") = "Sucessful Logout"
Response.redirect "default.asp?Email=" + Request.QueryString("Email")
End IF
END IF

Session("msg") = "Incorrect User or pass"
Response.redirect "default.asp?Email=" + Request.QueryString("Email")

%>