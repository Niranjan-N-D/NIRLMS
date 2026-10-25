<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

If Not IsAdmin() Then
    Response.Redirect("books.asp?msg=Only+Admin+can+delete+books")
    Response.End
End If

Dim bookId, conn
bookId = SafeInt(Request.QueryString("id"), 0)

If bookId > 0 Then
    Set conn = GetConnection()
    On Error Resume Next
    conn.Execute("DELETE FROM Books WHERE id = " & bookId)
    conn.Close
    Set conn = Nothing
    On Error GoTo 0
End If

Response.Redirect("books.asp?msg=Book+deleted+successfully")
Response.End
%>
