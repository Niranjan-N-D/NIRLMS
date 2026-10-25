<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

If Not IsStaff() Then
    Response.Redirect("books.asp")
    Response.End
End If

Dim bookId, errorMessage, conn
bookId = SafeInt(Request.QueryString("id"), 0)
errorMessage = ""

If bookId = 0 Then
    Response.Redirect("books.asp")
    Response.End
End If

Set conn = GetConnection()

' Handle Form Submission (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim title, isbn, publisher, priceVal, authorId, categoryId, copies
    title = Trim(Request.Form("title"))
    isbn = Trim(Request.Form("isbn"))
    publisher = Trim(Request.Form("publisher"))
    priceVal = SafeFloat(Request.Form("price"), 350.00)
    authorId = SafeInt(Request.Form("author_id"), 0)
    categoryId = SafeInt(Request.Form("category_id"), 0)
    copies = SafeInt(Request.Form("copies"), 0)

    If title = "" Or isbn = "" Then
        errorMessage = "Book Title and ISBN are required."
    Else
        Dim updateSql, isAvailable
        If copies > 0 Then isAvailable = 1 Else isAvailable = 0

        updateSql = "UPDATE Books SET " & _
                    "title = " & SQLQuote(title) & ", " & _
                    "isbn = " & SQLQuote(isbn) & ", " & _
                    "publisher = " & SQLQuote(publisher) & ", " & _
                    "price = " & priceVal & ", " & _
                    "author_id = " & authorId & ", " & _
                    "category_id = " & categoryId & ", " & _
                    "copies_available = " & copies & ", " & _
                    "is_available = " & isAvailable & " " & _
                    "WHERE id = " & bookId

        On Error Resume Next
        conn.Execute(updateSql)
        If Err.Number <> 0 Then
            errorMessage = "Failed to update book: " & Err.Description
        Else
            conn.Close
            Set conn = Nothing
            Response.Redirect("books.asp?msg=Book+updated+successfully")
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' Fetch Current Book Details
Dim rsBook
Set rsBook = conn.Execute("SELECT id, title, isbn, publisher, price, author_id, category_id, copies_available, is_available FROM Books WHERE id = " & bookId)

If rsBook.EOF Then
    conn.Close
    Set conn = Nothing
    Response.Redirect("books.asp")
    Response.End
End If

' Fetch Authors and Categories for Select Lists
Dim rsAuthors, rsCategories
Set rsAuthors = conn.Execute("SELECT id, name FROM Authors ORDER BY name ASC")
Set rsCategories = conn.Execute("SELECT id, name FROM Categories ORDER BY name ASC")

RenderHeader "Edit Book #" & bookId
%>

<div class="card" style="max-width: 650px; margin: 0 auto;">
    <div class="card-header">
        <h2 style="margin:0; color:#2c3e50;">Edit Book Details</h2>
        <a href="books.asp" class="btn btn-secondary">Cancel</a>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
    <% End If %>

    <form action="books_edit.asp?id=<%= bookId %>" method="POST">
        <div class="form-group">
            <label for="title">Book Title *</label>
            <input type="text" id="title" name="title" value="<%= CleanText(rsBook("title") & "") %>" required>
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:15px;">
            <div class="form-group">
                <label for="isbn">ISBN Code *</label>
                <input type="text" id="isbn" name="isbn" value="<%= CleanText(rsBook("isbn") & "") %>" required>
            </div>
            <div class="form-group">
                <label for="publisher">Publisher</label>
                <input type="text" id="publisher" name="publisher" value="<%= CleanText(rsBook("publisher") & "") %>">
            </div>
            <div class="form-group">
                <label for="price">Price of Book (&#8377;)</label>
                <input type="number" step="0.01" id="price" name="price" value="<%= SafeFloat(rsBook("price"), 0.00) %>" required>
            </div>
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:15px;">
            <div class="form-group">
                <label for="author_id">Author</label>
                <select id="author_id" name="author_id">
                    <option value="0">-- Select Author --</option>
                    <%
                    If Not rsAuthors Is Nothing And rsAuthors.State = 1 Then
                        Do While Not rsAuthors.EOF
                            Dim selAuthor
                            If CStr(rsAuthors("id")) = CStr(rsBook("author_id")) Then selAuthor = "selected" Else selAuthor = ""
                    %>
                        <option value="<%= rsAuthors("id") %>" <%= selAuthor %>><%= CleanText(rsAuthors("name") & "") %></option>
                    <%
                            rsAuthors.MoveNext
                        Loop
                    End If
                    %>
                </select>
            </div>

            <div class="form-group">
                <label for="category_id">Category</label>
                <select id="category_id" name="category_id">
                    <option value="0">-- Select Category --</option>
                    <%
                    If Not rsCategories Is Nothing And rsCategories.State = 1 Then
                        Do While Not rsCategories.EOF
                            Dim selCat
                            If CStr(rsCategories("id")) = CStr(rsBook("category_id")) Then selCat = "selected" Else selCat = ""
                    %>
                        <option value="<%= rsCategories("id") %>" <%= selCat %>><%= CleanText(rsCategories("name") & "") %></option>
                    <%
                            rsCategories.MoveNext
                        Loop
                    End If
                    %>
                </select>
            </div>
        </div>

        <div class="form-group">
            <label for="copies">Copies Available</label>
            <input type="number" id="copies" name="copies" value="<%= rsBook("copies_available") %>" min="0" required>
        </div>

        <div style="margin-top:20px; text-align:right;">
            <button type="submit" class="btn btn-primary" style="padding:10px 20px;">Update Book</button>
        </div>
    </form>
</div>

<%
rsBook.Close
conn.Close
Set conn = Nothing
RenderFooter
%>
