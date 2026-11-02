<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim searchKeyword, categoryId, categoryName, conn, rs, sql, dbErrorMsg
Dim searchPlaceholder, clearSearchUrl

searchKeyword = Trim(Request.QueryString("q"))
categoryId = SafeInt(Request.QueryString("category_id"), 0)
categoryName = ""
dbErrorMsg = ""

' --------------------------------------------------------------------
' Pagination & Lazy-Load Parameters (50 books per chunk for instant load)
' --------------------------------------------------------------------
Dim page, pageSize, isAjax, offsetVal, totalBooksCount, totalPages, rsCount
page = SafeInt(Request.QueryString("page"), 1)
If page < 1 Then page = 1

pageSize = SafeInt(Request.QueryString("page_size"), 50)
If pageSize < 10 Then pageSize = 50
If pageSize > 250 Then pageSize = 250

isAjax = (Request.QueryString("ajax") = "1")

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

' 2. Construct WHERE clause
Dim whereParts
whereParts = "COALESCE(b.is_deleted, 0) = 0"

If categoryId > 0 Then
    whereParts = whereParts & " AND b.category_id = " & categoryId
End If

If searchKeyword <> "" Then
    Dim kwCond
    kwCond = "(b.title LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR b.isbn LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR a.name LIKE " & SQLQuote("%" & searchKeyword & "%") & _
             " OR c.name LIKE " & SQLQuote("%" & searchKeyword & "%") & ")"
    whereParts = whereParts & " AND " & kwCond
End If

' 3. Fast Total Count (Indexed)
Dim countSql
countSql = "SELECT COUNT(*) AS total_cnt FROM Books b "
If searchKeyword <> "" Then
    countSql = countSql & "LEFT JOIN Authors a ON b.author_id = a.id LEFT JOIN Categories c ON b.category_id = c.id "
End If
countSql = countSql & "WHERE " & whereParts

totalBooksCount = 0
On Error Resume Next
Set rsCount = conn.Execute(countSql)
If Not (rsCount Is Nothing Or rsCount.State = 0) Then
    If Not rsCount.EOF Then
        totalBooksCount = SafeInt(rsCount("total_cnt"), 0)
    End If
    rsCount.Close
    Set rsCount = Nothing
End If
On Error GoTo 0

totalPages = (totalBooksCount + pageSize - 1) \ pageSize
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages
offsetVal = (page - 1) * pageSize

' 4. Paginated Books Query (LIMIT + OFFSET)
sql = "SELECT b.id, b.title, b.isbn, b.publisher, b.price, b.copies_available, b.is_available, " & _
      "a.name AS author_name, c.name AS category_name, " & _
      "COALESCE(fr.avg_rating, 0) AS avg_rating, " & _
      "COALESCE(fr.review_count, 0) AS review_count, " & _
      "COALESCE(ab.active_borrowed, 0) AS active_borrowed " & _
      "FROM Books b " & _
      "LEFT JOIN Authors a ON b.author_id = a.id " & _
      "LEFT JOIN Categories c ON b.category_id = c.id " & _
      "LEFT JOIN (SELECT book_id, AVG(rating) AS avg_rating, COUNT(*) AS review_count FROM Feedbacks GROUP BY book_id) fr ON fr.book_id = b.id " & _
      "LEFT JOIN (SELECT book_id, COUNT(*) AS active_borrowed FROM Borrowings WHERE status IN ('Borrowed', 'Overdue') GROUP BY book_id) ab ON ab.book_id = b.id " & _
      "WHERE " & whereParts & " " & _
      "ORDER BY b.id ASC " & _
      "LIMIT " & pageSize & " OFFSET " & offsetVal

' --------------------------------------------------------------------
' Helper Subroutine to Render a Book Row (Used by both main page and AJAX lazy-loader)
' --------------------------------------------------------------------
Sub RenderBookRow(rec)
    Dim rAvg, rCount, rStar, rAvail, rBorrowed, rTotal
    Dim sBadge, sSubtext
    rAvg = Round(SafeFloat(rec("avg_rating"), 0.0), 1)
    rCount = SafeInt(rec("review_count"), 0)
    rAvail = SafeInt(rec("copies_available"), 0)
    rBorrowed = SafeInt(rec("active_borrowed"), 0)
    rTotal = rAvail + rBorrowed

    ' Parse bilingual title (Separate Kannada and English)
    Dim rawTitle, mainTitle, subTitle, pOpen
    rawTitle = CleanText(rec("title") & "")
    mainTitle = rawTitle
    subTitle = ""
    pOpen = InStr(rawTitle, "(")
    If pOpen > 1 Then
        mainTitle = Trim(Left(rawTitle, pOpen - 1))
        subTitle = Trim(Mid(rawTitle, pOpen + 1))
        If Right(subTitle, 1) = ")" Then
            subTitle = Left(subTitle, Len(subTitle) - 1)
        End If
    End If

    ' Status Pill
    If rAvail > 0 Then
        sBadge = "<span class='status-pill status-pill-available'>● " & rAvail & " Available</span>"
        If rBorrowed > 0 Then
            sSubtext = "<div class='status-subtext'>" & rBorrowed & " on loan &bull; " & rTotal & " total</div>"
        Else
            sSubtext = "<div class='status-subtext'>" & rTotal & " in stock</div>"
        End If
    ElseIf rBorrowed > 0 Then
        sBadge = "<span class='status-pill status-pill-loaned'>● Loaned Out</span>"
        sSubtext = "<div class='status-subtext'>" & rBorrowed & " on loan &bull; " & rTotal & " total</div>"
    Else
        sBadge = "<span class='status-pill status-pill-out'>● Out of Stock</span>"
        sSubtext = "<div class='status-subtext'>0 available</div>"
    End If
%>
    <tr id="book-row-<%= rec("id") %>">
        <td class="td-id">#<%= rec("id") %></td>
        <td class="td-title">
            <div class="book-title-kannada"><%= mainTitle %></div>
            <% If subTitle <> "" Then %>
                <div class="book-title-english"><%= subTitle %></div>
            <% End If %>
        </td>
        <td class="td-author">
            <span class="author-name"><%= CleanText(rec("author_name") & "") %></span>
        </td>
        <td class="td-category">
            <span class="cat-pill"><%= CleanText(rec("category_name") & "") %></span>
        </td>
        <td class="td-code">
            <span class="isbn-badge"><%= CleanText(rec("isbn") & "") %></span>
        </td>
        <td class="td-price">
            <span class="price-val">&#8377;<%= FormatNumber(SafeFloat(rec("price"), 0), 2) %></span>
        </td>
        <td class="td-status">
            <%= sBadge %>
            <%= sSubtext %>
        </td>
        <td class="td-actions">
            <div class="action-grid">
                <% If HasPermission("feedback") Then %>
                    <a href="feedback.asp?book_id=<%= rec("id") %>" class="action-btn btn-act-review" title="Write Review">★ Review</a>
                <% End If %>

                <% If rAvail > 0 Then %>
                    <% If HasPermission("requests") Then %>
                    <a href="request_book.asp?id=<%= rec("id") %>" class="action-btn btn-act-request" title="Request Book">📖 Request</a>
                    <% End If %>
                <% ElseIf rTotal > 0 Then %>
                    <% If HasPermission("requests") Then %>
                    <a href="request_book.asp?id=<%= rec("id") %>" class="action-btn btn-act-reserve" title="Reserve when returned">⏳ Reserve</a>
                    <% End If %>
                <% End If %>

                <% If IsStaff() Then %>
                    <a href="books_edit.asp?id=<%= rec("id") %>" class="action-btn btn-act-edit" title="Edit details">✏️ Edit</a>
                    <a href="books_delete.asp?id=<%= rec("id") %>" class="action-btn btn-act-delete" onclick="return confirm('Delete this book?');" title="Delete">🗑️ Delete</a>
                <% End If %>
            </div>
        </td>
    </tr>
<%
End Sub

' --------------------------------------------------------------------
' AJAX Handler: Return ONLY table rows for background lazy loading
' --------------------------------------------------------------------
If isAjax Then
    Response.ContentType = "text/html"
    On Error Resume Next
    Set rs = conn.Execute(sql)
    If Not (rs Is Nothing) Then
        If rs.State = 1 Then
            Do While Not rs.EOF
                RenderBookRow rs
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    End If
    If Not (conn Is Nothing) Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Response.End
End If

' --------------------------------------------------------------------
' Normal Page Load: Execute first chunk and render full page
' --------------------------------------------------------------------
On Error Resume Next
Set rs = conn.Execute(sql)
If Err.Number <> 0 Then
    dbErrorMsg = Err.Description
    Err.Clear
End If
On Error GoTo 0

' Fetch Categories for Filter Dropdown
Dim rsAllCats
Set rsAllCats = conn.Execute("SELECT id, name FROM Categories WHERE COALESCE(is_deleted, 0) = 0 ORDER BY name ASC")

RenderHeader IIf(IsStaff(), "Book Catalog & Inventory", "Browse & Request Books")
%>

<style>
/* Modern Card & Table Container - Fit to Screen at 100% Zoom */
.books-card {
    background: #ffffff;
    border-radius: 14px;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04), 0 1px 3px rgba(0, 0, 0, 0.02);
    border: 1px solid #e2e8f0;
    padding: 16px 18px;
    margin-bottom: 24px;
    width: 100%;
    box-sizing: border-box;
}
.books-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    margin-bottom: 16px;
}
.books-title-group h2 {
    font-size: 20px;
    font-weight: 800;
    color: #0f172a;
    margin: 0 0 2px 0;
    letter-spacing: -0.3px;
}
.books-title-group p {
    font-size: 13px;
    color: #64748b;
    margin: 0;
}

/* Toolbar & Filters */
.filter-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 10px;
    background: #f8fafc;
    padding: 10px 14px;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    margin-bottom: 16px;
}
.filter-form {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
    flex: 1;
    max-width: 800px;
}
.search-wrapper {
    position: relative;
    flex: 1;
    min-width: 200px;
    display: flex;
    align-items: center;
}
.search-icon {
    position: absolute;
    left: 10px;
    color: #94a3b8;
    pointer-events: none;
}
.search-input {
    width: 100%;
    height: 36px;
    padding: 0 12px 0 34px;
    border: 1.5px solid #cbd5e1;
    border-radius: 7px;
    font-size: 13px;
    color: #0f172a;
    background: #ffffff;
    transition: all 0.15s ease;
}
.search-input:focus {
    outline: none;
    border-color: #0077f6;
    box-shadow: 0 0 0 3px rgba(0, 119, 246, 0.12);
}
.category-select {
    height: 36px;
    padding: 0 10px;
    border: 1.5px solid #cbd5e1;
    border-radius: 7px;
    font-size: 12.5px;
    color: #1e293b;
    background: #ffffff;
    cursor: pointer;
    max-width: 220px;
}
.category-select:focus {
    outline: none;
    border-color: #0077f6;
}

/* Full Width Fit-To-Screen Table Architecture */
.table-responsive {
    width: 100%;
    overflow-x: hidden;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    background: #ffffff;
}
.custom-table {
    width: 100%;
    table-layout: fixed;
    border-collapse: collapse;
    font-size: 12.5px;
    background: #ffffff;
}
.custom-table th {
    background: #f8fafc;
    color: #475569;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 9px 8px;
    text-align: left;
    border-bottom: 2px solid #e2e8f0;
    vertical-align: bottom;
}
.th-sub {
    font-size: 9.5px;
    font-weight: 500;
    color: #94a3b8;
    text-transform: none;
    letter-spacing: normal;
    display: block;
    margin-top: 1px;
}
.custom-table td {
    padding: 8px 8px;
    border-bottom: 1px solid #f1f5f9;
    vertical-align: middle;
    box-sizing: border-box;
}
.custom-table tr:hover td {
    background-color: #f8fafc;
}

/* Column Specific Styles - Budgeted Proportions */
.td-id {
    color: #94a3b8;
    font-size: 11.5px;
    font-weight: 700;
    text-align: center;
    padding-left: 4px !important;
    padding-right: 4px !important;
}
.td-title {
    overflow: hidden;
}
.book-title-kannada {
    font-size: 13.5px;
    font-weight: 700;
    color: #0f172a;
    line-height: 1.3;
    word-break: break-word;
}
.book-title-english {
    font-size: 11.5px;
    color: #64748b;
    margin-top: 1px;
    line-height: 1.25;
    word-break: break-word;
}
.td-author {
    color: #334155;
    font-weight: 600;
    font-size: 12px;
    word-break: break-word;
    line-height: 1.3;
}
.cat-pill {
    display: inline-block;
    padding: 2px 7px;
    border-radius: 10px;
    background: #f1f5f9;
    color: #0369a1;
    font-size: 11px;
    font-weight: 700;
    border: 1px solid #e0f2fe;
    word-break: break-word;
    line-height: 1.2;
}
.isbn-badge {
    display: inline-block;
    padding: 2px 6px;
    border-radius: 5px;
    background: #faf5ff;
    color: #7e22ce;
    font-size: 10.5px;
    font-weight: 700;
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
    border: 1px solid #f3e8ff;
    white-space: nowrap;
}
.price-val {
    color: #059669;
    font-weight: 800;
    font-size: 12.5px;
    white-space: nowrap;
}

/* Status Pill */
.status-pill {
    display: inline-block;
    padding: 2px 7px;
    border-radius: 10px;
    font-size: 11px;
    font-weight: 700;
    line-height: 1.2;
    white-space: nowrap;
}
.status-pill-available {
    background: #dcfce7;
    color: #15803d;
    border: 1px solid #bbf7d0;
}
.status-pill-loaned {
    background: #fef3c7;
    color: #b45309;
    border: 1px solid #fde68a;
}
.status-pill-out {
    background: #fee2e2;
    color: #b91c1c;
    border: 1px solid #fecaca;
}
.status-subtext {
    font-size: 10px;
    color: #64748b;
    margin-top: 2px;
    white-space: nowrap;
}

/* 2x2 Sleek Compact Action Grid */
.td-actions {
    text-align: center;
    padding: 6px 4px !important;
}
.action-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 3px;
    width: 100%;
    max-width: 124px;
    margin: 0 auto;
}
.action-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 2px;
    padding: 4px 3px;
    border-radius: 5px;
    font-size: 10px;
    font-weight: 700;
    color: #ffffff !important;
    text-decoration: none;
    border: none;
    cursor: pointer;
    box-sizing: border-box;
    transition: transform 0.12s ease, opacity 0.12s ease;
    white-space: nowrap;
    line-height: 1.1;
}
.action-btn:hover {
    transform: translateY(-1px);
    opacity: 0.92;
}
.btn-act-review { background: #d97706; }
.btn-act-request { background: #16a34a; }
.btn-act-reserve { background: #ea580c; }
.btn-act-edit { background: #0284c7; }
.btn-act-delete { background: #dc2626; }

/* Pagination & Toolbar Controls */
.size-pill {
    padding: 3px 8px;
    border-radius: 5px;
    background: #ffffff;
    color: #475569;
    font-size: 11.5px;
    font-weight: 700;
    text-decoration: none;
    border: 1px solid #cbd5e1;
    transition: all 0.15s ease;
}
.size-pill.active {
    background: #0077f6;
    color: #ffffff;
    border-color: #0077f6;
}
.page-btn {
    padding: 6px 12px;
    border-radius: 7px;
    background: #ffffff;
    color: #1e293b;
    border: 1.5px solid #cbd5e1;
    font-size: 12.5px;
    font-weight: 700;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    gap: 5px;
    transition: all 0.15s ease;
}
.page-btn:hover {
    background: #f8fafc;
    border-color: #0077f6;
    color: #0077f6;
}
.page-btn.disabled {
    opacity: 0.4;
    pointer-events: none;
}
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
</style>

<div class="books-card">
    <!-- Header -->
    <div class="books-header">
        <div class="books-title-group">
            <h2>Book Catalog &amp; Inventory</h2>
            <p>Showing collection of <strong><%= totalBooksCount %></strong> registered library volumes</p>
        </div>
        <div style="display:flex; gap:10px; align-items:center;">
            <a href="categories.asp" class="btn" style="background:#f1f5f9; color:#0f172a; border:1px solid #cbd5e1; font-weight:700; border-radius:10px; padding:9px 16px; text-decoration:none;">🏷️ Browse Categories</a>
            <% If IsStaff() Then %>
                <a href="books_add.asp" class="btn btn-success" style="font-weight:700; border-radius:10px; padding:9px 18px;">+ Add New Book</a>
            <% End If %>
        </div>
    </div>

    <!-- Active Category Banner -->
    <% If categoryId > 0 Then %>
        <div style="background:#e0f2fe; border:1px solid #7dd3fc; border-radius:12px; padding:12px 18px; margin-bottom:18px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;">
            <div style="display:flex; align-items:center; gap:12px;">
                <span style="font-size:24px;">🏷️</span>
                <div>
                    <div style="font-size:11.5px; text-transform:uppercase; letter-spacing:0.5px; color:#0369a1; font-weight:800;">Filtered by Category</div>
                    <div style="font-size:16px; font-weight:800; color:#0c4a6e;">
                        <%= Server.HTMLEncode(categoryName) %> (<%= totalBooksCount %> books)
                    </div>
                </div>
            </div>
            <a href="books.asp" class="btn" style="background:#ffffff; color:#b91c1c; border:1px solid #fca5a5; font-size:12.5px; font-weight:700; padding:6px 14px; border-radius:8px; text-decoration:none;">✕ Clear Category</a>
        </div>
    <% End If %>

    <!-- Filter & Search Toolbar -->
    <div class="filter-bar">
        <form action="books.asp" method="GET" class="filter-form">
            <input type="hidden" name="page_size" value="<%= pageSize %>">
            
            <div class="search-wrapper">
                <svg class="search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input type="text" name="q" value="<%= Server.HTMLEncode(searchKeyword) %>" placeholder="Search by title, author, or accession code..." class="search-input">
            </div>

            <select name="category_id" class="category-select" onchange="this.form.submit()">
                <option value="0">All Categories</option>
                <% If Not (rsAllCats Is Nothing) Then
                    Do While Not rsAllCats.EOF %>
                        <option value="<%= rsAllCats("id") %>" <%= IIf(categoryId = SafeInt(rsAllCats("id"), 0), "selected", "") %>><%= Server.HTMLEncode(rsAllCats("name")) %></option>
                    <% rsAllCats.MoveNext
                    Loop
                    rsAllCats.Close
                End If %>
            </select>

            <button type="submit" class="btn btn-primary" style="height:40px; padding:0 18px; font-weight:700; border-radius:8px;">Search</button>
            <% If searchKeyword <> "" Then %>
                <a href="books.asp<%= IIf(categoryId > 0, "?category_id=" & categoryId, "") %>" class="btn btn-secondary" style="height:40px; line-height:38px; padding:0 14px; border-radius:8px; text-decoration:none;">Clear</a>
            <% End If %>
        </form>

        <div style="display:flex; align-items:center; gap:10px; font-size:13px; color:#475569;">
            <span>Show:</span>
            <a href="books.asp?page=1&page_size=25&category_id=<%= categoryId %>&q=<%= Server.URLEncode(searchKeyword) %>" class="size-pill <%= IIf(pageSize = 25, "active", "") %>">25</a>
            <a href="books.asp?page=1&page_size=50&category_id=<%= categoryId %>&q=<%= Server.URLEncode(searchKeyword) %>" class="size-pill <%= IIf(pageSize = 50, "active", "") %>">50</a>
            <a href="books.asp?page=1&page_size=100&category_id=<%= categoryId %>&q=<%= Server.URLEncode(searchKeyword) %>" class="size-pill <%= IIf(pageSize = 100, "active", "") %>">100</a>
        </div>
    </div>

    <!-- Fit-To-Screen Table -->
    <div class="table-responsive">
        <table class="custom-table">
            <colgroup>
                <col style="width: 46px;">
                <col style="width: 31%;">
                <col style="width: 15%;">
                <col style="width: 12%;">
                <col style="width: 10%;">
                <col style="width: 7%;">
                <col style="width: 11%;">
                <col style="width: 128px;">
            </colgroup>
            <thead>
                <tr>
                    <th style="text-align:center;">ID</th>
                    <th>Title <span class="th-sub">(ಪುಸ್ತಕದ ಹೆಸರು)</span></th>
                    <th>Author <span class="th-sub">(ಲೇಖಕರು)</span></th>
                    <th>Category</th>
                    <th>Acc / ISBN</th>
                    <th>Price (&#8377;)</th>
                    <th>Status</th>
                    <th style="text-align:center;">Action</th>
                </tr>
            </thead>
            <tbody id="booksTableBody">
                <%
                Dim hasBooks, currentRowsLoaded
                hasBooks = False
                currentRowsLoaded = 0

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
                        <td colspan="8" style="text-align:center; color:#94a3b8; padding:36px;">
                            <% If dbErrorMsg <> "" Then %>
                                <span style="color:#e74c3c; font-weight:700;">Query Error: <%= Server.HTMLEncode(dbErrorMsg) %></span>
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
                        currentRowsLoaded = currentRowsLoaded + 1
                        RenderBookRow rs
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

    <!-- Lazy Load & Pagination Bottom Bar -->
    <% If totalBooksCount > 0 Then %>
        <div id="loadMoreSection" style="margin:28px 0 10px 0; display:flex; flex-direction:column; align-items:center; gap:14px;">
            <% If page < totalPages Then %>
                <button id="btnLoadMore" onclick="loadMoreBooks()" class="btn" style="background:#0077f6; color:#ffffff; font-size:14px; font-weight:700; padding:11px 26px; border-radius:10px; border:none; cursor:pointer; box-shadow:0 4px 14px rgba(0,119,246,0.25); display:inline-flex; align-items:center; gap:8px;">
                    <span>📥 Load More Books (<span id="countLoaded"><%= offsetVal + currentRowsLoaded %></span> of <%= totalBooksCount %>)</span>
                </button>
                <div id="loadingIndicator" style="display:none; color:#0077f6; font-size:14px; font-weight:600; padding:8px;">
                    <span style="display:inline-block; animation:spin 1s linear infinite;">⏳</span> Loading more books...
                </div>
            <% Else %>
                <div style="color:#64748b; font-size:13px; font-weight:700; background:#f1f5f9; padding:8px 20px; border-radius:20px;">
                    ✓ All <%= totalBooksCount %> Books Loaded
                </div>
            <% End If %>

            <!-- Page Navigation Links -->
            <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; justify-content:center;">
                <a href="books.asp?page=<%= page - 1 %>&page_size=<%= pageSize %>&category_id=<%= categoryId %>&q=<%= Server.URLEncode(searchKeyword) %>" class="page-btn <%= IIf(page <= 1, "disabled", "") %>">&laquo; Previous</a>

                <span style="font-size:13px; font-weight:700; color:#475569; padding:0 8px;">
                    Page <%= page %> of <%= totalPages %>
                </span>

                <a href="books.asp?page=<%= page + 1 %>&page_size=<%= pageSize %>&category_id=<%= categoryId %>&q=<%= Server.URLEncode(searchKeyword) %>" class="page-btn <%= IIf(page >= totalPages, "disabled", "") %>">Next &raquo;</a>
            </div>
        </div>
    <% End If %>
</div>

<script>
var currentPage = <%= page %>;
var totalPages = <%= totalPages %>;
var pageSize = <%= pageSize %>;
var totalBooks = <%= totalBooksCount %>;
var currentLoaded = <%= offsetVal + currentRowsLoaded %>;
var categoryId = <%= categoryId %>;
var searchQ = "<%= Replace(Replace(searchKeyword, "\", "\\"), """", "\""") %>";
var isLoading = false;

function loadMoreBooks() {
    if (isLoading || currentPage >= totalPages) return;
    isLoading = true;

    var btn = document.getElementById("btnLoadMore");
    var spinner = document.getElementById("loadingIndicator");
    if (btn) btn.style.display = "none";
    if (spinner) spinner.style.display = "block";

    var nextPage = currentPage + 1;
    var url = "books.asp?page=" + nextPage + "&page_size=" + pageSize + "&category_id=" + categoryId + "&q=" + encodeURIComponent(searchQ) + "&ajax=1";

    fetch(url)
        .then(function(res) { return res.text(); })
        .then(function(html) {
            if (html.trim().length > 0) {
                var tbody = document.getElementById("booksTableBody");
                tbody.insertAdjacentHTML("beforeend", html);
                currentPage = nextPage;
                currentLoaded = Math.min(totalBooks, currentLoaded + pageSize);
                
                var countElem = document.getElementById("countLoaded");
                if (countElem) countElem.innerText = currentLoaded;
            }

            if (spinner) spinner.style.display = "none";

            if (currentPage >= totalPages) {
                if (btn) {
                    btn.outerHTML = '<div style="color:#64748b; font-size:13px; font-weight:700; background:#f1f5f9; padding:8px 20px; border-radius:20px;">✓ All ' + totalBooks + ' Books Loaded</div>';
                }
            } else {
                if (btn) btn.style.display = "inline-flex";
            }
            isLoading = false;
        })
        .catch(function(err) {
            console.error("Error loading more books:", err);
            if (spinner) spinner.style.display = "none";
            if (btn) btn.style.display = "inline-flex";
            isLoading = false;
        });
}

// Auto lazy-load when scrolling near the bottom of the table
window.addEventListener("scroll", function() {
    if (currentPage < totalPages && !isLoading) {
        var scrollPosition = window.innerHeight + window.pageYOffset;
        var threshold = document.documentElement.scrollHeight - 350;
        if (scrollPosition >= threshold) {
            loadMoreBooks();
        }
    }
});
</script>

<% RenderFooter %>
