<!--#include file="includes/db_config.asp"-->
<%
If IsAuthenticated() Then
    Response.Redirect("dashboard.asp")
Else
    Response.Redirect("login.asp")
End If
%>
