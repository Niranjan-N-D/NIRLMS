<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim searchKeyword, conn, rs, sql
searchKeyword = Trim(Request.QueryString("q"))

Set conn = GetConnection()

sql = "SELECT b.id, b.title, b.isbn, b.publisher, b.price, b.copies_available, b.is_available, a.name AS author_name, c.name AS category_name, " & _
      "(SELECT COALESCE(AVG(rating), 0) FROM Feedbacks WHERE book_id = b.id) AS avg_rating, " & _
      "(SELECT COUNT(*) FROM Feedbacks WHERE book_id = b.id) AS review_count " & _
      "FROM Books b " & _
      "LEFT JOIN Authors a ON b.author_id = a.id " & _
      "LEFT JOIN Categories c ON b.category_id = c.id "

If searchKeyword <> "" Then
    sql = sql & "WHERE b.title LIKE " & SQLQuote("%" & searchKeyword & "%") & " OR b.isbn LIKE " & SQLQuote("%" & searchKeyword & "%") & " OR a.name LIKE " & SQLQuote("%" & searchKeyword & "%") & " "
End If

sql = sql & "ORDER BY b.id DESC"

On Error Resume Next
Set rs = conn.Execute(sql)
On Error GoTo 0

RenderHeader "Book Management & Catalog"
%>

<div class="card">
    <div class="card-header">
        <h2 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Book Catalog & Inventory", "Browse & Request Books") %></h2>
        <% If IsStaff() Then %>
            <a href="books_add.asp" class="btn btn-success">+ Add New Book</a>
        <% End If %>
    </div>

    <!-- Search Form -->
    <form action="books.asp" method="GET" style="display:flex; gap:10px; margin-bottom: 20px;">
        <input type="text" name="q" value="<%= CleanText(searchKeyword) %>" placeholder="Search by title, author, or ISBN (e.g. Kannada / English)..." style="flex:1; padding:8px 12px; border:1px solid #ccc; border-radius:4px;">
        <button type="submit" class="btn btn-primary">Search</button>
        <% If searchKeyword <> "" Then %>
            <a href="books.asp" class="btn btn-secondary">Clear</a>
        <% End If %>
    </form>

    <table>
        <thead>
            <tr>
                <th>ID</th>
                <th>Title (ಪುಸ್ತಕದ ಹೆಸರು)</th>
                <th>Author (ಲೇಖಕರು)</th>
                <th>Category</th>
                <th>ISBN</th>
                <th>Price (&#8377;)</th>
                <th>Rating</th>
                <th>Status</th>
                <th>Action</th>
            </tr>
        </thead>
        <tbody>
            <%
            If rs Is Nothing Or rs.State = 0 Or rs.EOF Then
            %>
                <tr>
                    <td colspan="9" style="text-align:center; color:#95a5a6;">No books found matching criteria.</td>
                </tr>
            <%
            Else
                Do While Not rs.EOF
                    Dim avgR, revCount, starStr
                    avgR = Round(SafeInt(rs("avg_rating"), 0), 1)
                    revCount = SafeInt(rs("review_count"), 0)
                    If avgR > 0 Then
                        starStr = "⭐ " & avgR & " (" & revCount & ")"
                    Else
                        starStr = "<span style='color:#95a5a6; font-size:12px;'>No reviews</span>"
                    End If
            %>
                <tr>
                    <td>#<%= rs("id") %></td>
                    <td><strong style="font-size:15px; color:#2c3e50;"><%= CleanText(rs("title") & "") %></strong></td>
                    <td><%= CleanText(rs("author_name") & "") %></td>
                    <td><%= CleanText(rs("category_name") & "") %></td>
                    <td><code><%= CleanText(rs("isbn") & "") %></code></td>
                    <td><strong style="color:#27ae60;">&#8377;<%= FormatNumber(SafeFloat(rs("price"), 0), 2) %></strong></td>
                    <td><%= starStr %></td>
                    <td>
                        <% If CBool(rs("is_available")) And SafeInt(rs("copies_available"), 0) > 0 Then %>
                            <span class="badge bg-success">Available (<%= rs("copies_available") %>)</span>
                        <% Else %>
                            <span class="badge bg-danger">Out of Stock</span>
                        <% End If %>
                    </td>
                    <td>
                        <div style="display:flex; gap:5px; flex-wrap:wrap;">
                            <% If HasPermission("feedback") Then %>
                                <a href="feedback.asp?book_id=<%= rs("id") %>" class="btn btn-warning" style="padding:4px 8px; font-size:12px;" title="Write Review">⭐ Review</a>
                            <% End If %>

                            <% If IsStaff() Then %>
                                <a href="books_edit.asp?id=<%= rs("id") %>" class="btn btn-primary" style="padding:4px 8px; font-size:12px;">Edit</a>
                                <a href="books_delete.asp?id=<%= rs("id") %>" class="btn btn-danger" style="padding:4px 8px; font-size:12px;" onclick="return confirm('Delete this book?');">Delete</a>
                            <% End If %>

                            <% If CBool(rs("is_available")) And SafeInt(rs("copies_available"), 0) > 0 Then %>
                                <% If HasPermission("requests") Then %>
                                    <a href="request_book.asp?id=<%= rs("id") %>" class="btn btn-success" style="padding:4px 8px; font-size:12px;">Request Book</a>
                                <% End If %>
                            <% End If %>
                        </div>
                    </td>
                </tr>
            <%
                    rs.MoveNext
                Loop
                rs.Close
            End If

            conn.Close
            Set conn = Nothing
            %>
        </tbody>
    </table>
</div>

<% RenderFooter %>
