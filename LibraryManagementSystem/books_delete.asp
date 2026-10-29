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
    ' Soft delete: preserve borrowing history and audit trail while hiding book from catalog
    conn.Execute("UPDATE Books SET is_deleted = 1, is_available = 0, copies_available = 0 WHERE id = " & bookId)
    conn.Close
    Set conn = Nothing
    On Error GoTo 0
End If

Response.Redirect("books.asp?msg=Book+deleted+successfully")
Response.End
%>
