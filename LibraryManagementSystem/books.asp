<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim searchKeyword, conn, rs, sql, statusFilter
Dim totalTitles, totalShelfCopies, totalLoanedCopies, totalInventoryCopies, totalOverdueCopies
Dim countInStock, countCirculation, countLoanedOut, countOutOfStock, countOverdue
Dim avgR, revCount, starStr, availCopies, borrowedCopies, overdueCount, totalBookCopies
Dim rowStatusType, statusBadgeHtml, stockBreakdownHtml, isAvailBool

searchKeyword = Trim(Request.QueryString("q"))
statusFilter = LCase(Trim(Request.QueryString("filter")))
If statusFilter = "" Then statusFilter = "all"

Set conn = GetConnection()

' --------------------------------------------------------------------
' 1. Calculate Comprehensive Inventory Summary Metrics
' --------------------------------------------------------------------
totalTitles = 0
totalShelfCopies = 0
totalLoanedCopies = 0
totalInventoryCopies = 0
totalOverdueCopies = 0
countInStock = 0
countCirculation = 0
countLoanedOut = 0
countOutOfStock = 0
countOverdue = 0

On Error Resume Next

Dim rsInvSummary
Set rsInvSummary = conn.Execute("SELECT COUNT(*) AS cnt_titles, COALESCE(SUM(copies_available), 0) AS sum_shelf FROM Books")
If Not rsInvSummary.EOF Then
    totalTitles = SafeInt(rsInvSummary("cnt_titles"), 0)
    totalShelfCopies = SafeInt(rsInvSummary("sum_shelf"), 0)
End If
rsInvSummary.Close

Dim rsBorrowSummary
Set rsBorrowSummary = conn.Execute("SELECT " & _
    "COUNT(CASE WHEN status IN ('Borrowed', 'Overdue') THEN 1 END) AS cnt_loaned, " & _
    "COUNT(CASE WHEN (status = 'Overdue' OR (status = 'Borrowed' AND due_date < CURRENT_DATE())) THEN 1 END) AS cnt_overdue " & _
    "FROM Borrowings")
If Not rsBorrowSummary.EOF Then
    totalLoanedCopies = SafeInt(rsBorrowSummary("cnt_loaned"), 0)
    totalOverdueCopies = SafeInt(rsBorrowSummary("cnt_overdue"), 0)
End If
rsBorrowSummary.Close

totalInventoryCopies = totalShelfCopies + totalLoanedCopies

' --------------------------------------------------------------------
' 2. Fetch Books Catalog with Full Inventory Breakdown
' --------------------------------------------------------------------
sql = "SELECT b.id, b.title, b.isbn, b.publisher, b.price, b.copies_available, b.is_available, " & _
      "a.name AS author_name, c.name AS category_name, " & _
      "(SELECT COALESCE(AVG(rating), 0) FROM Feedbacks WHERE book_id = b.id) AS avg_rating, " & _
      "(SELECT COUNT(*) FROM Feedbacks WHERE book_id = b.id) AS review_count, " & _
      "(SELECT COUNT(*) FROM Borrowings WHERE book_id = b.id AND status IN ('Borrowed', 'Overdue')) AS active_borrowed, " & _
      "(SELECT COUNT(*) FROM Borrowings WHERE book_id = b.id AND (status = 'Overdue' OR (status = 'Borrowed' AND due_date < CURRENT_DATE()))) AS overdue_count " & _
      "FROM Books b " & _
      "LEFT JOIN Authors a ON b.author_id = a.id " & _
      "LEFT JOIN Categories c ON b.category_id = c.id "

If searchKeyword <> "" Then
    sql = sql & "WHERE b.title LIKE " & SQLQuote("%" & searchKeyword & "%") & " OR b.isbn LIKE " & SQLQuote("%" & searchKeyword & "%") & " OR a.name LIKE " & SQLQuote("%" & searchKeyword & "%") & " "
End If

sql = sql & "ORDER BY b.id DESC"

Set rs = conn.Execute(sql)
On Error GoTo 0

RenderHeader IIf(IsStaff(), "Book Catalog & Inventory Management", "Browse & Request Books")
%>

<style>
/* Inventory Status Overview Cards */
.inv-kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}
.inv-kpi-card {
    background: #ffffff;
    border-radius: 8px;
    padding: 18px 20px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    display: flex;
    align-items: center;
    gap: 16px;
    border-left: 5px solid #3498db;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.inv-kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}
.inv-kpi-icon {
    width: 48px;
    height: 48px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    flex-shrink: 0;
}
.inv-kpi-value {
    font-size: 26px;
    font-weight: 800;
    color: #2c3e50;
    line-height: 1.1;
}
.inv-kpi-label {
    font-size: 13px;
    font-weight: 600;
    color: #7f8c8d;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-top: 4px;
}

/* Inventory Breakdown Pills */
.inv-breakdown-box {
    display: flex;
    flex-direction: column;
    gap: 4px;
    font-size: 12.5px;
}
.inv-pill {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 11.5px;
    font-weight: 600;
}
.inv-pill-shelf { background: #e8f8f5; color: #16a085; border: 1px solid #a3e4d7; }
.inv-pill-loaned { background: #fef5e7; color: #d35400; border: 1px solid #f8c471; }
.inv-pill-overdue { background: #fdedec; color: #c0392b; border: 1px solid #f5b7b1; }
.inv-pill-total { background: #ebf5fb; color: #2980b9; border: 1px solid #aed6f1; }

/* Filter Buttons */
.filter-btn-group {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}
.filter-btn {
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 600;
    border-radius: 4px;
    border: 1px solid #dcdde1;
    background: #ffffff;
    color: #4b6584;
    cursor: pointer;
    transition: all 0.2s ease;
}
.filter-btn:hover {
    background: #f1f2f6;
    color: #2f3542;
}
.filter-btn.active {
    background: #2980b9;
    color: #ffffff;
    border-color: #2980b9;
    box-shadow: 0 2px 6px rgba(41, 128, 185, 0.3);
}
</style>

<% If IsStaff() Then %>
<!-- 1. Real-Time Inventory KPI Metrics -->
<div class="inv-kpi-grid">
    <div class="inv-kpi-card" style="border-left-color: #3498db;">
        <div class="inv-kpi-icon" style="background: rgba(52, 152, 219, 0.12); color: #2980b9;">📚</div>
        <div>
            <div class="inv-kpi-value"><%= totalTitles %></div>
            <div class="inv-kpi-label">Catalog Titles</div>
        </div>
    </div>

    <div class="inv-kpi-card" style="border-left-color: #2ecc71;">
        <div class="inv-kpi-icon" style="background: rgba(46, 204, 113, 0.12); color: #27ae60;">📗</div>
        <div>
            <div class="inv-kpi-value"><%= totalShelfCopies %></div>
            <div class="inv-kpi-label">On-Shelf Available</div>
        </div>
    </div>

    <div class="inv-kpi-card" style="border-left-color: #f39c12;">
        <div class="inv-kpi-icon" style="background: rgba(243, 156, 18, 0.12); color: #d68910;">📙</div>
        <div>
            <div class="inv-kpi-value"><%= totalLoanedCopies %></div>
            <div class="inv-kpi-label">Out on Loan</div>
        </div>
    </div>

    <div class="inv-kpi-card" style="border-left-color: #8e44ad;">
        <div class="inv-kpi-icon" style="background: rgba(142, 68, 173, 0.12); color: #8e44ad;">📦</div>
        <div>
            <div class="inv-kpi-value"><%= totalInventoryCopies %></div>
            <div class="inv-kpi-label">Total Inventory Copies</div>
        </div>
    </div>

    <% If totalOverdueCopies > 0 Then %>
    <div class="inv-kpi-card" style="border-left-color: #e74c3c;">
        <div class="inv-kpi-icon" style="background: rgba(231, 76, 60, 0.12); color: #c0392b;">🚨</div>
        <div>
            <div class="inv-kpi-value" style="color: #c0392b;"><%= totalOverdueCopies %></div>
            <div class="inv-kpi-label" style="color: #c0392b;">Overdue Returns</div>
        </div>
    </div>
    <% End If %>
</div>
<% End If %>

<div class="card">
    <div class="card-header" style="flex-wrap:wrap; gap:12px;">
        <div>
            <h2 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Book Catalog & Inventory Status", "Browse & Request Books") %></h2>
            <p style="margin:4px 0 0 0; color:#7f8c8d; font-size:13px;">
                <%= IIf(IsStaff(), "Real-time physical inventory monitoring, shelf stock count, active circulation, and pricing.", "Browse available defense electronics and literature titles. Submit borrowing requests directly.") %>
            </p>
        </div>
        <div style="display:flex; gap:10px;">
            <% If IsStaff() Then %>
                <a href="books_add.asp" class="btn btn-success">+ Add New Book</a>
                <a href="analytics.asp" class="btn btn-primary" style="background:#8e44ad;">📊 Inventory Analytics</a>
            <% End If %>
        </div>
    </div>

    <!-- Inventory Filter & Search Bar -->
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #f1f1f1;">
        <!-- Status Filter Buttons -->
        <div class="filter-btn-group">
            <button type="button" class="filter-btn active" onclick="filterCatalog('all', this)" id="btn-all">
                All Books (<span id="count-all">0</span>)
            </button>
            <button type="button" class="filter-btn" onclick="filterCatalog('instock', this)" id="btn-instock">
                🟢 In Stock (<span id="count-instock">0</span>)
            </button>
            <button type="button" class="filter-btn" onclick="filterCatalog('circulation', this)" id="btn-circulation">
                🔄 In Circulation (<span id="count-circulation">0</span>)
            </button>
            <button type="button" class="filter-btn" onclick="filterCatalog('loaned', this)" id="btn-loaned">
                ⏳ Fully Loaned (<span id="count-loaned">0</span>)
            </button>
            <button type="button" class="filter-btn" onclick="filterCatalog('outofstock', this)" id="btn-outofstock">
                🔴 Out of Stock (<span id="count-outofstock">0</span>)
            </button>
        </div>

        <!-- Live Search Input -->
        <div style="min-width:280px; flex:1; max-width:360px;">
            <input type="text" id="catalogSearch" onkeyup="searchCatalog()" placeholder="🔍 Live search title, author, category, ISBN..." style="width:100%; padding:8px 12px; border:1px solid #ccc; border-radius:4px; font-size:13.5px;">
        </div>
    </div>

    <table id="catalogTable">
        <thead>
            <tr>
                <th>ID</th>
                <th>Title (ಪುಸ್ತಕದ ಹೆಸರು)</th>
                <th>Author & Category</th>
                <th>ISBN</th>
                <th>Price (&#8377;)</th>
                <th>Rating</th>
                <th style="min-width:180px;">Inventory Status</th>
                <th style="text-align:right;">Action</th>
            </tr>
        </thead>
        <tbody>
            <%
            If rs Is Nothing Or rs.State = 0 Or rs.EOF Then
            %>
                <tr>
                    <td colspan="8" style="text-align:center; color:#95a5a6; padding:30px;">No books found in the catalog.</td>
                </tr>
            <%
            Else
                Do While Not rs.EOF
                    avgR = Round(SafeFloat(rs("avg_rating"), 0.0), 1)
                    revCount = SafeInt(rs("review_count"), 0)
                    availCopies = SafeInt(rs("copies_available"), 0)
                    borrowedCopies = SafeInt(rs("active_borrowed"), 0)
                    overdueCount = SafeInt(rs("overdue_count"), 0)
                    totalBookCopies = availCopies + borrowedCopies
                    isAvailBool = CBool(rs("is_available"))

                    If avgR > 0 Then
                        starStr = "⭐ " & avgR & " <span style='color:#7f8c8d; font-size:11.5px;'>(" & revCount & ")</span>"
                    Else
                        starStr = "<span style='color:#95a5a6; font-size:12px;'>No reviews</span>"
                    End If

                    ' Determine Accurate Inventory Status & Category Tag
                    If totalBookCopies = 0 Then
                        rowStatusType = "outofstock"
                        statusBadgeHtml = "<span class='badge bg-danger' style='font-size:12px;'>🔴 Out of Stock</span>"
                    ElseIf availCopies = 0 And borrowedCopies > 0 Then
                        rowStatusType = "loaned"
                        statusBadgeHtml = "<span class='badge' style='background:#f39c12; color:#fff; font-size:12px;'>⏳ All Loaned Out</span>"
                    ElseIf availCopies > 0 And borrowedCopies > 0 Then
                        rowStatusType = "circulation"
                        statusBadgeHtml = "<span class='badge' style='background:#2980b9; color:#fff; font-size:12px;'>🔄 In Circulation</span>"
                    Else
                        rowStatusType = "instock"
                        statusBadgeHtml = "<span class='badge bg-success' style='font-size:12px;'>🟢 Available on Shelf</span>"
                    End If
            %>
                <tr class="book-row" data-status="<%= rowStatusType %>" data-overdue="<%= IIf(overdueCount > 0, "true", "false") %>">
                    <td><strong>#<%= rs("id") %></strong></td>
                    <td>
                        <strong style="font-size:14.5px; color:#2c3e50;"><%= CleanText(rs("title") & "") %></strong>
                        <% If rs("publisher") & "" <> "" Then %>
                            <div style="font-size:11.5px; color:#7f8c8d; margin-top:2px;">🏢 <%= CleanText(rs("publisher") & "") %></div>
                        <% End If %>
                    </td>
                    <td>
                        <div style="color:#34495e; font-weight:600;"><%= CleanText(rs("author_name") & "") %></div>
                        <div style="font-size:12px; color:#7f8c8d; margin-top:2px;">🏷️ <%= CleanText(rs("category_name") & "") %></div>
                    </td>
                    <td><code><%= CleanText(rs("isbn") & "") %></code></td>
                    <td><strong style="color:#27ae60; font-size:14px;">&#8377;<%= FormatNumber(SafeFloat(rs("price"), 0), 2) %></strong></td>
                    <td><%= starStr %></td>
                    <td>
                        <!-- Enhanced Comprehensive Inventory Status -->
                        <div style="margin-bottom:5px;">
                            <%= statusBadgeHtml %>
                        </div>
                        <div class="inv-breakdown-box">
                            <div style="display:flex; gap:6px; flex-wrap:wrap;">
                                <span class="inv-pill inv-pill-shelf" title="Physical copies on library shelf">
                                    📗 Shelf: <strong><%= availCopies %></strong>
                                </span>
                                <% If borrowedCopies > 0 Then %>
                                    <span class="inv-pill inv-pill-loaned" title="Copies currently borrowed by staff members">
                                        📙 Loaned: <strong><%= borrowedCopies %></strong>
                                    </span>
                                <% End If %>
                                <% If overdueCount > 0 Then %>
                                    <span class="inv-pill inv-pill-overdue" title="Copies past due return date">
                                        🚨 Overdue: <strong><%= overdueCount %></strong>
                                    </span>
                                <% End If %>
                                <span class="inv-pill inv-pill-total" title="Total physical library inventory">
                                    📦 Total: <strong><%= totalBookCopies %></strong>
                                </span>
                            </div>
                        </div>
                    </td>
                    <td style="text-align:right;">
                        <div style="display:inline-flex; gap:5px; flex-wrap:wrap; justify-content:flex-end;">
                            <% If HasPermission("feedback") Then %>
                                <a href="feedback.asp?book_id=<%= rs("id") %>" class="btn btn-warning" style="padding:4px 8px; font-size:12px;" title="Write Review">⭐ Review</a>
                            <% End If %>

                            <% If IsStaff() Then %>
                                <a href="books_edit.asp?id=<%= rs("id") %>" class="btn btn-primary" style="padding:4px 8px; font-size:12px;">Edit</a>
                                <a href="books_delete.asp?id=<%= rs("id") %>" class="btn btn-danger" style="padding:4px 8px; font-size:12px;" onclick="return confirm('Are you sure you want to delete this book?');">Delete</a>
                            <% End If %>

                            <% If availCopies > 0 Then %>
                                <% If HasPermission("requests") Then %>
                                    <a href="request_book.asp?id=<%= rs("id") %>" class="btn btn-success" style="padding:4px 8px; font-size:12px;">Request Book</a>
                                <% End If %>
                            <% ElseIf totalBookCopies > 0 Then %>
                                <% If HasPermission("requests") Then %>
                                    <a href="request_book.asp?id=<%= rs("id") %>" class="btn btn-warning" style="padding:4px 8px; font-size:12px; background:#e67e22;" title="Reserve next available copy">Reserve Copy</a>
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

    <div id="noMatchRow" style="display:none; text-align:center; padding:35px; color:#95a5a6; font-size:14.5px;">
        <span style="font-size:26px;">🔍</span><br>
        <strong>No catalog books found matching your search or inventory filter.</strong>
    </div>
</div>

<script>
// Real-Time Inventory & Search Filtering
var activeFilter = 'all';

function filterCatalog(status, btn) {
    activeFilter = status;
    document.querySelectorAll('.filter-btn').forEach(function(b) {
        b.classList.remove('active');
    });
    if (btn) btn.classList.add('active');
    applyFilters();
}

function searchCatalog() {
    applyFilters();
}

function applyFilters() {
    var query = document.getElementById('catalogSearch').value.toLowerCase().trim();
    var rows = document.querySelectorAll('.book-row');
    var visibleCount = 0;

    rows.forEach(function(row) {
        var rowStatus = row.getAttribute('data-status');
        var text = row.innerText.toLowerCase();

        var matchesStatus = false;
        if (activeFilter === 'all') {
            matchesStatus = true;
        } else if (activeFilter === 'instock') {
            matchesStatus = (rowStatus === 'instock');
        } else if (activeFilter === 'circulation') {
            matchesStatus = (rowStatus === 'circulation');
        } else if (activeFilter === 'loaned') {
            matchesStatus = (rowStatus === 'loaned');
        } else if (activeFilter === 'outofstock') {
            matchesStatus = (rowStatus === 'outofstock');
        }

        var matchesSearch = (query === '' || text.indexOf(query) !== -1);

        if (matchesStatus && matchesSearch) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    });

    var noMatch = document.getElementById('noMatchRow');
    if (noMatch) {
        noMatch.style.display = (visibleCount === 0) ? 'block' : 'none';
    }
}

// Compute Dynamic Filter Counts on Page Load
document.addEventListener('DOMContentLoaded', function() {
    var rows = document.querySelectorAll('.book-row');
    var cAll = rows.length;
    var cInStock = 0;
    var cCirculation = 0;
    var cLoaned = 0;
    var cOutOfStock = 0;

    rows.forEach(function(row) {
        var st = row.getAttribute('data-status');
        if (st === 'instock') cInStock++;
        else if (st === 'circulation') cCirculation++;
        else if (st === 'loaned') cLoaned++;
        else if (st === 'outofstock') cOutOfStock++;
    });

    if (document.getElementById('count-all')) document.getElementById('count-all').innerText = cAll;
    if (document.getElementById('count-instock')) document.getElementById('count-instock').innerText = cInStock;
    if (document.getElementById('count-circulation')) document.getElementById('count-circulation').innerText = cCirculation;
    if (document.getElementById('count-loaned')) document.getElementById('count-loaned').innerText = cLoaned;
    if (document.getElementById('count-outofstock')) document.getElementById('count-outofstock').innerText = cOutOfStock;
});
</script>

<% RenderFooter %>
