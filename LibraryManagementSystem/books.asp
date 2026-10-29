<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim searchKeyword, categoryId, categoryName, conn, rs, sql, dbErrorMsg
Dim avgR, revCount, starStr, availCopies, borrowedCopies, totalBookCopies
Dim statusBadgeHtml, statusSubtextHtml
Dim searchPlaceholder, clearSearchUrl

searchKeyword = Trim(Request.QueryString("q"))
categoryId = SafeInt(Request.QueryString("category_id"), 0)
categoryName = ""
dbErrorMsg = ""

Set conn = GetConnection()

' 1. Lookup Category Name if category_id filter provided
If categoryId > 0 Then
    On Error Resume Next
    Dim rsCatCheck
    Set rsCatCheck = conn.Execute("SELECT name FROM Categories WHERE id = " & categoryId)
    If Err.Number = 0 Then
        If Not (rsCatCheck Is Nothing) Then
            If rsCatCheck.State = 1 Then
                If Not rsCatCheck.EOF Then
                    categoryName = rsCatCheck("name") & ""
                End If
                rsCatCheck.Close
            End If
            Set rsCatCheck = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
End If

If categoryId > 0 And categoryName = "" Then
    categoryName = "Category #" & categoryId
End If

' 2. Construct Books Query
sql = "SELECT b.id, b.title, b.isbn, b.publisher, b.price, b.copies_available, b.is_available, " & _
      "a.name AS author_name, c.name AS category_name, " & _
      "(SELECT COALESCE(AVG(rating), 0) FROM Feedbacks WHERE book_id = b.id) AS avg_rating, " & _
      "(SELECT COUNT(*) FROM Feedbacks WHERE book_id = b.id) AS review_count, " & _
      "(SELECT COUNT(*) FROM Borrowings WHERE book_id = b.id AND status IN ('Borrowed', 'Overdue')) AS active_borrowed " & _
      "FROM Books b " & _
      "LEFT JOIN Authors a ON b.author_id = a.id " & _
      "LEFT JOIN Categories c ON b.category_id = c.id "

Dim whereParts
whereParts = ""

If categoryId > 0 Then
    whereParts = "b.category_id = " & categoryId
End If

If searchKeyword <> "" Then
    Dim kwCond
    kwCond = "(b.title LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR b.isbn LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR a.name LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR c.name LIKE " & SQLQuote("%" & searchKeyword & "%") & ")"
    If whereParts <> "" Then
        whereParts = whereParts & " AND " & kwCond
    Else
        whereParts = kwCond
    End If
End If

If whereParts <> "" Then
    sql = sql & "WHERE " & whereParts & " "
End If

sql = sql & "ORDER BY b.id DESC"

On Error Resume Next
Set rs = conn.Execute(sql)
If Err.Number <> 0 Then
    dbErrorMsg = Err.Description
    Err.Clear
End If
On Error GoTo 0

' Determine Search Placeholders and URLs
If categoryId > 0 Then
    searchPlaceholder = "Search within " & categoryName & "..."
    clearSearchUrl = "books.asp?category_id=" & categoryId
Else
    searchPlaceholder = "Search by title, author, category, or ISBN..."
    clearSearchUrl = "books.asp"
End If

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
    background-color: #f1c40f;
    color: #2c3e50 !important;
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
    <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <h2 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Book Catalog & Inventory", "Browse & Request Books") %></h2>
        <div style="display:flex; gap:10px; align-items:center;">
            <a href="categories.asp" class="btn" style="background:#f1f5f9; color:#0B1320; border:1px solid #cbd5e1; font-weight:600; text-decoration:none;">🏷️ Browse Categories</a>
            <% If IsStaff() Then %>
                <a href="books_add.asp" class="btn btn-success">+ Add New Book</a>
            <% End If %>
        </div>
    </div>

    <!-- Active Category Filter Banner -->
    <% If categoryId > 0 Then %>
        <div style="background:#e0f2fe; border:1px solid #7dd3fc; border-radius:10px; padding:12px 18px; margin-bottom:18px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;">
            <div style="display:flex; align-items:center; gap:12px;">
                <span style="font-size:24px;">🏷️</span>
                <div>
                    <div style="font-size:12px; text-transform:uppercase; letter-spacing:0.5px; color:#0369a1; font-weight:700;">Category Filter Active</div>
                    <div style="font-size:16px; font-weight:800; color:#0c4a6e;">
                        <%= Server.HTMLEncode(categoryName) %>
                    </div>
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:10px;">
                <a href="categories.asp" class="btn" style="background:#0284c7; color:#fff; font-size:13px; font-weight:600; padding:6px 14px; border-radius:6px; text-decoration:none;">&larr; All Categories</a>
                <a href="books.asp" class="btn" style="background:#fff; color:#b91c1c; border:1px solid #fca5a5; font-size:13px; font-weight:600; padding:6px 14px; border-radius:6px; text-decoration:none;">✕ Clear Filter</a>
            </div>
        </div>
    <% End If %>

    <!-- Search Form -->
    <form action="books.asp" method="GET" style="display:flex; gap:10px; margin-bottom: 20px;">
        <% If categoryId > 0 Then %>
            <input type="hidden" name="category_id" value="<%= categoryId %>">
        <% End If %>
        <input type="text" name="q" value="<%= Server.HTMLEncode(searchKeyword) %>" placeholder="<%= Server.HTMLEncode(searchPlaceholder) %>" style="flex:1; padding:8px 12px; border:1px solid #ccc; border-radius:4px;">
        <button type="submit" class="btn btn-primary">Search</button>
        <% If searchKeyword <> "" Then %>
            <a href="<%= clearSearchUrl %>" class="btn btn-secondary">Clear Search</a>
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
            Dim hasBooks
            hasBooks = False
            If Not (rs Is Nothing) Then
                If rs.State = 1 Then
                    If Not rs.EOF Then
                        hasBooks = True
                    End If
                End If
            End If

            If Not hasBooks Then
            %>
                <tr>
                    <td colspan="9" style="text-align:center; color:#95a5a6; padding:24px;">
                        <% If dbErrorMsg <> "" Then %>
                            <span style="color:#e74c3c; font-weight:600;">Query Error: <%= Server.HTMLEncode(dbErrorMsg) %></span>
                        <% ElseIf categoryId > 0 Then %>
                            No books currently registered under <strong><%= Server.HTMLEncode(categoryName) %></strong>.
                        <% Else %>
                            No books found matching your criteria.
                        <% End If %>
                    </td>
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

            If Not (conn Is Nothing) Then
                If conn.State = 1 Then conn.Close
                Set conn = Nothing
            End If
            %>
        </tbody>
    </table>
</div>

<% RenderFooter %>
