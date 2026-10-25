<%@ Language=VBScript %>
<%
' Clear all session variables
Session.Contents.RemoveAll()
Session.Abandon()

' Redirect to login page
Response.Redirect("login.asp?msg=You+have+been+logged+out+successfully")
Response.End
%>
