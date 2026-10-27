<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim searchKeyword, conn, rs, sql
Dim avgR, revCount, starStr, availCopies, borrowedCopies, totalBookCopies
Dim statusBadgeHtml, statusSubtextHtml

searchKeyword = Trim(Request.QueryString("q"))

Set conn = GetConnection()

sql = "SELECT b.id, b.title, b.isbn, b.publisher, b.price, b.copies_available, b.is_available, " & _
      "a.name AS author_name, c.name AS category_name, " & _
      "(SELECT COALESCE(AVG(rating), 0) FROM Feedbacks WHERE book_id = b.id) AS avg_rating, " & _
      "(SELECT COUNT(*) FROM Feedbacks WHERE book_id = b.id) AS review_count, " & _
      "(SELECT COUNT(*) FROM Borrowings WHERE book_id = b.id AND status IN ('Borrowed', 'Overdue')) AS active_borrowed " & _
      "FROM Books b " & _
      "LEFT JOIN Authors a ON b.author_id = a.id " & _
      "LEFT JOIN Categories c ON b.category_id = c.id "

If searchKeyword <> "" Then
    sql = sql & "WHERE b.title LIKE " & SQLQuote("%" & searchKeyword & "%") & _
                " OR b.isbn LIKE " & SQLQuote("%" & searchKeyword & "%") & _
                " OR a.name LIKE " & SQLQuote("%" & searchKeyword & "%") & _
                " OR c.name LIKE " & SQLQuote("%" & searchKeyword & "%") & " "
End If

sql = sql & "ORDER BY b.id DESC"

On Error Resume Next
Set rs = conn.Execute(sql)
On Error GoTo 0

RenderHeader IIf(IsStaff(), "Book Catalog & Inventory", "Browse & Request Books")
%>

<style>
/* Status Badge Pill & Subtext (Exact User Spec Match) */
.status-pill-wrap {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
}
.status-pill {
    display: inline-block;
    padding: 7px 20px;
    border-radius: 24px;
    font-size: 15px;
    font-weight: 700;
    color: #ffffff;
    line-height: 1.2;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
.status-pill-available {
    background-color: #27ae60;
}
.status-pill-loaned {
    background-color: #f39c12;
}
.status-pill-out {
    background-color: #e74c3c;
}
.status-subtext {
    font-size: 13.5px;
    font-weight: 500;
    color: #576574;
    margin-top: 6px;
    letter-spacing: 0.2px;
}

/* Action Button Vertical Stack (Exact User Spec Match) */
.action-stack {
    display: flex;
    flex-direction: column;
    gap: 7px;
    width: 125px;
}
.action-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    width: 100%;
    padding: 7px 12px;
    border-radius: 7px;
    font-size: 13.5px;
    font-weight: 700;
    color: #ffffff !important;
    text-decoration: none;
    border: none;
    cursor: pointer;
    box-sizing: border-box;
    transition: opacity 0.15s ease, transform 0.1s ease;
    text-align: center;
}
.action-btn:hover {
    opacity: 0.92;
    transform: translateY(-1px);
}
.btn-act-review {
    background-color: #f39c12;
}
.btn-act-edit {
    background-color: #3498db;
}
.btn-act-delete {
    background-color: #e74c3c;
}
.btn-act-request {
    background-color: #2ecc71;
}
.btn-act-reserve {
    background-color: #e67e22;
}
</style>

<div class="card">
    <div class="card-header" style="display:flex; justify-content:space-between; align-items:center;">
        <h2 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Book Catalog & Inventory", "Browse & Request Books") %></h2>
        <% If IsStaff() Then %>
            <a href="books_add.asp" class="btn btn-success">+ Add New Book</a>
        <% End If %>
    </div>

    <!-- Search Form -->
    <form action="books.asp" method="GET" style="display:flex; gap:10px; margin-bottom: 20px;">
        <input type="text" name="q" value="<%= CleanText(searchKeyword) %>" placeholder="Search by title, author, category, or ISBN..." style="flex:1; padding:8px 12px; border:1px solid #ccc; border-radius:4px;">
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
                <th style="min-width:160px;">Status</th>
                <th style="width:135px; text-align:center;">Action</th>
            </tr>
        </thead>
        <tbody>
            <%
            If rs Is Nothing Or rs.State = 0 Or rs.EOF Then
            %>
                <tr>
                    <td colspan="9" style="text-align:center; color:#95a5a6; padding:24px;">No books found matching your criteria.</td>
                </tr>
            <%
            Else
                Do While Not rs.EOF
                    avgR = Round(SafeFloat(rs("avg_rating"), 0.0), 1)
                    revCount = SafeInt(rs("review_count"), 0)
                    availCopies = SafeInt(rs("copies_available"), 0)
                    borrowedCopies = SafeInt(rs("active_borrowed"), 0)
                    totalBookCopies = availCopies + borrowedCopies

                    If avgR > 0 Then
                        starStr = "⭐ " & avgR & " (" & revCount & ")"
                    Else
                        starStr = "<span style='color:#95a5a6; font-size:12px;'>No reviews</span>"
                    End If

                    ' Clean Status Pill & Subtext (Matches Screenshot)
                    If availCopies > 0 Then
                        statusBadgeHtml = "<span class='status-pill status-pill-available'>Available (" & availCopies & ")</span>"
                        If borrowedCopies > 0 Then
                            statusSubtextHtml = "<div class='status-subtext'>" & borrowedCopies & " borrowed &bull; " & totalBookCopies & " total</div>"
                        Else
                            statusSubtextHtml = "<div class='status-subtext'>" & totalBookCopies & " in stock</div>"
                        End If
                    ElseIf borrowedCopies > 0 Then
                        statusBadgeHtml = "<span class='status-pill status-pill-loaned'>Loaned Out (" & borrowedCopies & ")</span>"
                        statusSubtextHtml = "<div class='status-subtext'>" & borrowedCopies & " borrowed &bull; " & totalBookCopies & " total</div>"
                    Else
                        statusBadgeHtml = "<span class='status-pill status-pill-out'>Out of Stock</span>"
                        statusSubtextHtml = "<div class='status-subtext'>0 copies available</div>"
                    End If
            %>
                <tr>
                    <td>#<%= rs("id") %></td>
                    <td><strong style="font-size:14.5px; color:#2c3e50;"><%= CleanText(rs("title") & "") %></strong></td>
                    <td><%= CleanText(rs("author_name") & "") %></td>
                    <td><%= CleanText(rs("category_name") & "") %></td>
                    <td><code><%= CleanText(rs("isbn") & "") %></code></td>
                    <td><strong style="color:#27ae60;">&#8377;<%= FormatNumber(SafeFloat(rs("price"), 0), 2) %></strong></td>
                    <td><%= starStr %></td>
                    <td>
                        <div class="status-pill-wrap">
                            <%= statusBadgeHtml %>
                            <%= statusSubtextHtml %>
                        </div>
                    </td>
                    <td>
                        <div class="action-stack">
                            <% If HasPermission("feedback") Then %>
                                <a href="feedback.asp?book_id=<%= rs("id") %>" class="action-btn btn-act-review" title="Write Review">★ Review</a>
                            <% End If %>

                            <% If IsStaff() Then %>
                                <a href="books_edit.asp?id=<%= rs("id") %>" class="action-btn btn-act-edit">Edit</a>
                                <a href="books_delete.asp?id=<%= rs("id") %>" class="action-btn btn-act-delete" onclick="return confirm('Delete this book?');">Delete</a>
                            <% End If %>

                            <% If availCopies > 0 Then %>
                                <% If HasPermission("requests") Then %>
                                <a href="request_book.asp?id=<%= rs("id") %>" class="action-btn btn-act-request">Request Book</a>
                                <% End If %>
                            <% ElseIf totalBookCopies > 0 Then %>
                                <% If HasPermission("requests") Then %>
                                <a href="request_book.asp?id=<%= rs("id") %>" class="action-btn btn-act-reserve" title="Reserve when returned">Reserve</a>
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
