<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

If Not IsStaff() Then
    Response.Redirect("books.asp")
    Response.End
End If

Dim errorMessage, conn
errorMessage = ""
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
    copies = SafeInt(Request.Form("copies"), 1)

    If title = "" Or isbn = "" Then
        errorMessage = "Book Title and ISBN are required fields."
    Else
        Dim insertSql, isAvailable
        If copies > 0 Then isAvailable = 1 Else isAvailable = 0

        insertSql = "INSERT INTO Books (title, isbn, publisher, price, author_id, category_id, copies_available, is_available) VALUES (" & _
                    SQLQuote(title) & ", " & _
                    SQLQuote(isbn) & ", " & _
                    SQLQuote(publisher) & ", " & _
                    priceVal & ", " & _
                    authorId & ", " & _
                    categoryId & ", " & _
                    copies & ", " & _
                    isAvailable & ")"

        On Error Resume Next
        conn.Execute(insertSql)
        If Err.Number <> 0 Then
            errorMessage = "Failed to insert book: " & Err.Description
        Else
            conn.Close
            Set conn = Nothing
            Response.Redirect("books.asp?msg=Book+added+successfully")
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' Fetch Authors and Categories for Select Lists
Dim rsAuthors, rsCategories
Set rsAuthors = conn.Execute("SELECT id, name FROM Authors ORDER BY name ASC")
Set rsCategories = conn.Execute("SELECT id, name FROM Categories ORDER BY name ASC")

RenderHeader "Add New Book"
%>

<div class="card" style="max-width: 650px; margin: 0 auto;">
    <div class="card-header">
        <h2 style="margin:0; color:#2c3e50;">Add New Book</h2>
        <a href="books.asp" class="btn btn-secondary">Back to Books</a>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
    <% End If %>

    <form action="books_add.asp" method="POST">
        <div class="form-group">
            <label for="title">Book Title *</label>
            <input type="text" id="title" name="title" required placeholder="e.g. ಕಾನೂರು ಹೆಗ್ಗಡಿತಿ (Kanooru Heggadithi) or Masterlink Systems">
        </div>

        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:15px;">
            <div class="form-group">
                <label for="isbn">ISBN Code *</label>
                <input type="text" id="isbn" name="isbn" required placeholder="e.g. 978-0123456789">
            </div>
            <div class="form-group">
                <label for="publisher">Publisher</label>
                <input type="text" id="publisher" name="publisher" placeholder="e.g. Sapna / BEL">
            </div>
            <div class="form-group">
                <label for="price">Price of Book (&#8377;)</label>
                <input type="number" step="0.01" id="price" name="price" value="350.00" required placeholder="e.g. 450.00">
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
                    %>
                        <option value="<%= rsAuthors("id") %>"><%= CleanText(rsAuthors("name") & "") %></option>
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
                    %>
                        <option value="<%= rsCategories("id") %>"><%= CleanText(rsCategories("name") & "") %></option>
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
            <input type="number" id="copies" name="copies" value="1" min="0" required>
        </div>

        <div style="margin-top:20px; text-align:right;">
            <button type="submit" class="btn btn-success" style="padding:10px 20px;">Save Book</button>
        </div>
    </form>
</div>

<%
conn.Close
Set conn = Nothing
RenderFooter
%>
