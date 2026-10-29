<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

' Restrict Analytics Access: Admin and Librarian only (staff excluded)
If Not IsStaff() Then
    Response.Redirect("dashboard.asp?msg=Access+restricted:+Analytics+is+only+available+to+Administrators+and+Librarians")
    Response.End
End If

Dim conn, rs
Set conn = GetConnection()

' --------------------------------------------------------------------
' 1. Calculate Core KPI Metrics
' --------------------------------------------------------------------
Dim totalTitles, totalCopiesAvailable, totalCurrentlyBorrowed, totalOverdue, totalReturned, totalBookValuation
Dim totalInventoryCopies, circulationRate

totalTitles = 0
totalCopiesAvailable = 0
totalCurrentlyBorrowed = 0
totalOverdue = 0
totalReturned = 0
totalBookValuation = 0.0

On Error Resume Next

' Total Unique Titles & Copies Available & Valuation
Set rs = conn.Execute("SELECT COUNT(*) AS cnt_titles, COALESCE(SUM(copies_available), 0) AS sum_available, COALESCE(SUM(copies_available * price), 0.0) AS total_val FROM Books")
If Not rs.EOF Then
    totalTitles = SafeInt(rs("cnt_titles"), 0)
    totalCopiesAvailable = SafeInt(rs("sum_available"), 0)
    totalBookValuation = SafeFloat(rs("total_val"), 0.0)
End If
rs.Close

' Active Borrowings & Overdue Count
Set rs = conn.Execute("SELECT " & _
                      "COUNT(CASE WHEN status = 'Borrowed' AND (due_date IS NULL OR due_date >= CURRENT_DATE()) THEN 1 END) AS cnt_active, " & _
                      "COUNT(CASE WHEN (status = 'Overdue' OR (status = 'Borrowed' AND due_date < CURRENT_DATE())) THEN 1 END) AS cnt_overdue, " & _
                      "COUNT(CASE WHEN status = 'Returned' THEN 1 END) AS cnt_returned " & _
                      "FROM Borrowings")
If Not rs.EOF Then
    totalCurrentlyBorrowed = SafeInt(rs("cnt_active"), 0)
    totalOverdue = SafeInt(rs("cnt_overdue"), 0)
    totalReturned = SafeInt(rs("cnt_returned"), 0)
End If
rs.Close

' Total Physical Inventory = Shelf Copies + Currently Out on Loan
totalInventoryCopies = totalCopiesAvailable + totalCurrentlyBorrowed + totalOverdue
If totalInventoryCopies > 0 Then
    circulationRate = Round(((totalCurrentlyBorrowed + totalOverdue) / totalInventoryCopies) * 100, 1)
Else
    circulationRate = 0
End If

' --------------------------------------------------------------------
' --------------------------------------------------------------------
' 2. Category Distribution & Stock (Stored into Arrays for Table & Chart)
' --------------------------------------------------------------------
Dim rsCatDist, catNames(), catCopies(), catLoans(), catTitles()
Dim catCount, ci
catCount = 0
ReDim catNames(50), catCopies(50), catLoans(50), catTitles(50)

Set rsCatDist = conn.Execute("SELECT c.id, c.name, COUNT(b.id) AS title_count, " & _
                             "COALESCE(SUM(b.copies_available), 0) AS copies_count, " & _
                             "(SELECT COUNT(*) FROM Borrowings br JOIN Books bk ON br.book_id = bk.id WHERE bk.category_id = c.id AND br.status IN ('Borrowed', 'Overdue')) AS active_loans " & _
                             "FROM Categories c " & _
                             "LEFT JOIN Books b ON c.id = b.category_id " & _
                             "GROUP BY c.id, c.name " & _
                             "ORDER BY copies_count DESC, title_count DESC")

If Not rsCatDist Is Nothing And rsCatDist.State = 1 Then
    Do While Not rsCatDist.EOF
        catNames(catCount) = CleanText(rsCatDist("name") & "")
        catTitles(catCount) = SafeInt(rsCatDist("title_count"), 0)
        catCopies(catCount) = SafeInt(rsCatDist("copies_count"), 0)
        catLoans(catCount) = SafeInt(rsCatDist("active_loans"), 0)
        catCount = catCount + 1
        rsCatDist.MoveNext
    Loop
    rsCatDist.Close
End If

' --------------------------------------------------------------------
' 3. Top Most In-Demand Books (Stored into Arrays for Table & Chart)
' --------------------------------------------------------------------
Dim rsTopBooks, topBookTitles(), topBookAuthors(), topBookPrices(), topBookLoans(), topBookLeft()
Dim topBookCount, bi
topBookCount = 0
ReDim topBookTitles(50), topBookAuthors(50), topBookPrices(50), topBookLoans(50), topBookLeft(50)

Set rsTopBooks = conn.Execute("SELECT b.id, b.title, b.price, a.name AS author_name, c.name AS category_name, " & _
                             "COUNT(br.id) AS borrow_count, b.copies_available " & _
                             "FROM Books b " & _
                             "LEFT JOIN Authors a ON b.author_id = a.id " & _
                             "LEFT JOIN Categories c ON b.category_id = c.id " & _
                             "LEFT JOIN Borrowings br ON b.id = br.book_id " & _
                             "GROUP BY b.id, b.title, b.price, a.name, c.name, b.copies_available " & _
                             "ORDER BY borrow_count DESC, b.title ASC LIMIT 6")

If Not rsTopBooks Is Nothing And rsTopBooks.State = 1 Then
    Do While Not rsTopBooks.EOF
        topBookTitles(topBookCount) = CleanText(rsTopBooks("title") & "")
        topBookAuthors(topBookCount) = CleanText(rsTopBooks("author_name") & "")
        topBookPrices(topBookCount) = SafeFloat(rsTopBooks("price"), 0)
        topBookLoans(topBookCount) = SafeInt(rsTopBooks("borrow_count"), 0)
        topBookLeft(topBookCount) = SafeInt(rsTopBooks("copies_available"), 0)
        topBookCount = topBookCount + 1
        rsTopBooks.MoveNext
    Loop
    rsTopBooks.Close
End If

' --------------------------------------------------------------------
' 4. Overdue Borrowings Detailed Log
' --------------------------------------------------------------------
Dim rsOverdueLog
Set rsOverdueLog = conn.Execute("SELECT b.id, bk.title, u.username, u.staff_number, " & _
                               "b.borrow_date, b.due_date, DATEDIFF(CURRENT_DATE(), b.due_date) AS days_late " & _
                               "FROM Borrowings b " & _
                               "JOIN Books bk ON b.book_id = bk.id " & _
                               "JOIN Users u ON b.user_id = u.id " & _
                               "WHERE b.status = 'Overdue' OR (b.status = 'Borrowed' AND b.due_date < CURRENT_DATE()) " & _
                               "ORDER BY days_late DESC")

' --------------------------------------------------------------------
' 5. Top Staff Readers
' --------------------------------------------------------------------
Dim rsTopBorrowers
Set rsTopBorrowers = conn.Execute("SELECT u.id, u.username, u.role, u.staff_number, " & _
                                 "COUNT(CASE WHEN b.status IN ('Borrowed', 'Overdue') THEN 1 END) AS current_held, " & _
                                 "COUNT(CASE WHEN b.status = 'Returned' THEN 1 END) AS total_returned " & _
                                 "FROM Users u " & _
                                 "JOIN Borrowings b ON u.id = b.user_id " & _
                                 "GROUP BY u.id, u.username, u.role, u.staff_number " & _
                                 "ORDER BY current_held DESC, total_returned DESC LIMIT 5")

On Error GoTo 0

RenderHeader "Library Analytics & Inventory Intelligence"
%>

<!-- Header with Action Buttons -->
<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:24px;">
    <div>
        <h1 style="margin:0; font-size:26px; color:#2c3e50; font-weight:800; letter-spacing:-0.02em;">
            📊 Library Analytics & Reports
        </h1>
        <p style="margin:6px 0 0 0; color:#7f8c8d; font-size:14px;">
            Comprehensive live overview of catalog volume, active circulations, shelf availability, and overdue audits.
        </p>
    </div>
    <div style="display:flex; gap:10px;">
        <button onclick="window.print()" class="btn btn-secondary" style="display:flex; align-items:center; gap:6px; padding:9px 16px;">
            🖨️ Print / Export PDF
        </button>
        <a href="borrowings.asp?filter=overdue" class="btn btn-danger" style="background:#e74c3c; display:flex; align-items:center; gap:6px; padding:9px 16px; font-weight:bold;">
            🚨 View Overdue Queue (<%= totalOverdue %>)
        </a>
    </div>
</div>

<!-- ==================================================================== -->
<!-- 1. Primary Metrics Dashboard Cards -->
<!-- ==================================================================== -->
<div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap:18px; margin-bottom:28px;">
    <!-- Total Titles -->
    <div class="card" style="margin:0; border-top: 4px solid #3498db; padding:20px;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <span style="font-size:12px; font-weight:700; color:#7f8c8d; text-transform:uppercase; letter-spacing:0.05em;">Book Titles</span>
                <h2 style="font-size:32px; margin:8px 0 4px 0; color:#2c3e50; font-weight:800;"><%= totalTitles %></h2>
                <span style="font-size:13px; color:#3498db; font-weight:600;">Unique Catalog Titles</span>
            </div>
            <span style="font-size:32px;">📚</span>
        </div>
    </div>

    <!-- Total Physical Copies -->
    <div class="card" style="margin:0; border-top: 4px solid #2ecc71; padding:20px;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <span style="font-size:12px; font-weight:700; color:#7f8c8d; text-transform:uppercase; letter-spacing:0.05em;">Total Physical Books</span>
                <h2 style="font-size:32px; margin:8px 0 4px 0; color:#27ae60; font-weight:800;"><%= totalInventoryCopies %></h2>
                <span style="font-size:13px; color:#27ae60; font-weight:600;">📦 <%= totalCopiesAvailable %> On Shelves Available</span>
            </div>
            <span style="font-size:32px;">📖</span>
        </div>
    </div>

    <!-- Currently Borrowed -->
    <div class="card" style="margin:0; border-top: 4px solid #f39c12; padding:20px;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <span style="font-size:12px; font-weight:700; color:#7f8c8d; text-transform:uppercase; letter-spacing:0.05em;">Currently Borrowed</span>
                <h2 style="font-size:32px; margin:8px 0 4px 0; color:#e67e22; font-weight:800;"><%= totalCurrentlyBorrowed %></h2>
                <span style="font-size:13px; color:#e67e22; font-weight:600;">🔄 Active Circulations</span>
            </div>
            <span style="font-size:32px;">🤝</span>
        </div>
    </div>

    <!-- Overdue (Alert Card) -->
    <div class="card" style="margin:0; border-top: 4px solid #e74c3c; padding:20px; background:<%= IIf(totalOverdue > 0, "#fdf3f2", "#ffffff") %>;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <span style="font-size:12px; font-weight:700; color:#c0392b; text-transform:uppercase; letter-spacing:0.05em;">Overdue (Day 16+)</span>
                <h2 style="font-size:32px; margin:8px 0 4px 0; color:#e74c3c; font-weight:800;"><%= totalOverdue %></h2>
                <span style="font-size:13px; color:#c0392b; font-weight:600;">
                    <%= IIf(totalOverdue > 0, "🚨 Requires Immediate Follow-Up", "✅ All loans within 15-day limit") %>
                </span>
            </div>
            <span style="font-size:32px;">⏰</span>
        </div>
    </div>

    <!-- Estimated Asset Value -->
    <div class="card" style="margin:0; border-top: 4px solid #9b59b6; padding:20px;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <span style="font-size:12px; font-weight:700; color:#7f8c8d; text-transform:uppercase; letter-spacing:0.05em;">Shelf Asset Value</span>
                <h2 style="font-size:30px; margin:8px 0 4px 0; color:#8e44ad; font-weight:800;">&#8377;<%= FormatNumber(totalBookValuation, 2) %></h2>
                <span style="font-size:13px; color:#8e44ad; font-weight:600;">💰 Total Shelf Stock Valuation</span>
            </div>
            <span style="font-size:32px;">🏷️</span>
        </div>
    </div>
</div>

<!-- ==================================================================== -->
<!-- 2. Visual Analytics Charts: Pie Chart & Bar Chart -->
<!-- ==================================================================== -->
<div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap:22px; margin-bottom:28px;">
    <!-- Chart 1: System Pie / Doughnut Chart -->
    <div class="card" style="margin:0; padding:22px; display:flex; flex-direction:column;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid #f1f5f9; padding-bottom:10px;">
            <div>
                <h3 style="margin:0; color:#0f172a; font-size:17px; font-weight:700;">🍩 System Inventory &amp; Status Breakdown</h3>
                <p style="margin:3px 0 0 0; color:#64748b; font-size:12.5px;">Live distribution of physical book copies by loan status</p>
            </div>
            <span class="status-pill status-pill-blue" style="font-size:12px; padding:4px 10px;"><%= circulationRate %>% Circulating</span>
        </div>
        <div style="position:relative; height:290px; width:100%; display:flex; align-items:center; justify-content:center;">
            <canvas id="systemPieChart"></canvas>
        </div>
    </div>

    <!-- Chart 2: Category Volumes & Circulation Bar Chart -->
    <div class="card" style="margin:0; padding:22px; display:flex; flex-direction:column;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid #f1f5f9; padding-bottom:10px;">
            <div>
                <h3 style="margin:0; color:#0f172a; font-size:17px; font-weight:700;">📊 Category Volumes &amp; Circulation</h3>
                <p style="margin:3px 0 0 0; color:#64748b; font-size:12.5px;">Shelf stock available vs active loans per category</p>
            </div>
            <a href="categories.asp" class="btn btn-secondary" style="padding:4px 10px; font-size:12px;">Categories ➔</a>
        </div>
        <div style="position:relative; height:290px; width:100%;">
            <canvas id="categoryBarChart"></canvas>
        </div>
    </div>
</div>

<!-- Circulation Progress Bar -->
<div class="card" style="margin-bottom:28px;">
    <div class="card-header" style="margin-bottom:15px; border-bottom:1px solid #eee;">
        <div>
            <h3 style="margin:0; color:#2c3e50; font-size:16px;">Shelf Availability vs. Active Circulation Ratio</h3>
            <p style="margin:4px 0 0 0; color:#7f8c8d; font-size:12.5px;">Proportion of catalog inventory on physical shelves vs out with members.</p>
        </div>
        <strong style="color:#2980b9; font-size:15px;"><%= circulationRate %>% Circulation Rate</strong>
    </div>

    <!-- Multi-Color Distribution Bar -->
    <div style="height:20px; width:100%; background:#ecf0f1; border-radius:10px; overflow:hidden; display:flex; margin-bottom:12px;">
        <%
        Dim pctAvail, pctBorrowed, pctOverdue
        If totalInventoryCopies > 0 Then
            pctAvail = Round((totalCopiesAvailable / totalInventoryCopies) * 100, 1)
            pctBorrowed = Round((totalCurrentlyBorrowed / totalInventoryCopies) * 100, 1)
            pctOverdue = Round((totalOverdue / totalInventoryCopies) * 100, 1)
        Else
            pctAvail = 100
            pctBorrowed = 0
            pctOverdue = 0
        End If
        %>
        <div style="width:<%= pctAvail %>%; background:#10b981;" title="Available on Shelves: <%= totalCopiesAvailable %> (<%= pctAvail %>%)"></div>
        <div style="width:<%= pctBorrowed %>%; background:#f59e0b;" title="Active Borrowings: <%= totalCurrentlyBorrowed %> (<%= pctBorrowed %>%)"></div>
        <div style="width:<%= pctOverdue %>%; background:#ef4444;" title="Overdue Loans: <%= totalOverdue %> (<%= pctOverdue %>%)"></div>
    </div>

    <!-- Legend -->
    <div style="display:flex; gap:20px; flex-wrap:wrap; font-size:13px; color:#555;">
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="display:inline-block; width:13px; height:13px; background:#10b981; border-radius:3px;"></span>
            <span><strong><%= totalCopiesAvailable %> Available on Shelves</strong> (<%= pctAvail %>%)</span>
        </div>
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="display:inline-block; width:13px; height:13px; background:#f59e0b; border-radius:3px;"></span>
            <span><strong><%= totalCurrentlyBorrowed %> Active Loans</strong> (<%= pctBorrowed %>%)</span>
        </div>
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="display:inline-block; width:13px; height:13px; background:#ef4444; border-radius:3px;"></span>
            <span><strong><%= totalOverdue %> Overdue</strong> (<%= pctOverdue %>%)</span>
        </div>
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="display:inline-block; width:13px; height:13px; background:#3b82f6; border-radius:3px;"></span>
            <span><strong><%= totalReturned %> Returned Historical</strong></span>
        </div>
    </div>
</div>

<style>
/* Analytics Table Baseline & Overflow Alignment */
.analytics-card {
    margin: 0;
    overflow: hidden;
    display: flex;
    flex-direction: column;
}
.analytics-tbl-wrap {
    width: 100%;
    overflow-x: auto;
}
.analytics-tbl {
    width: 100%;
    border-collapse: collapse;
}
.analytics-tbl th,
.analytics-tbl td {
    padding: 10px 12px;
    vertical-align: middle !important;
    border-bottom: 1px solid #eee;
}
.analytics-tbl th {
    background: #f8f9fa;
    color: #2c3e50;
    font-weight: 600;
    font-size: 13px;
    white-space: nowrap;
}
.col-nowrap {
    white-space: nowrap !important;
}
</style>

<!-- ==================================================================== -->
<!-- 3. Overdue Audits (If Any Exist) -->
<!-- ==================================================================== -->
<% If totalOverdue > 0 Then %>
<div class="card" style="border-left:6px solid #e74c3c; margin-bottom:28px;">
    <div class="card-header" style="border-bottom:1px solid #fce4e4;">
        <div>
            <h3 style="margin:0; color:#c0392b; font-size:18px;">🚨 Overdue Loans Requiring Immediate Return</h3>
            <p style="margin:4px 0 0 0; color:#e74c3c; font-size:13px;">Books past the 15-day return threshold. Staff must contact borrowers for return.</p>
        </div>
        <a href="borrowings.asp?filter=overdue" class="btn btn-danger" style="padding:6px 14px; font-size:13px;">Manage in Borrowings ➔</a>
    </div>

    <div class="analytics-tbl-wrap">
        <table class="analytics-tbl">
            <thead>
                <tr>
                    <th>#ID</th>
                    <th>Overdue Book Title</th>
                    <th>Borrower (Staff)</th>
                    <th>Staff Number</th>
                    <th>Borrowed On</th>
                    <th>Due Date (Day 15)</th>
                    <th style="text-align:center;">Days Late</th>
                    <th style="text-align:center; width:125px;">Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                Do While Not rsOverdueLog.EOF
                %>
                <tr style="background:#fff7f7;">
                    <td>#<%= rsOverdueLog("id") %></td>
                    <td><strong style="color:#c0392b;"><%= CleanText(rsOverdueLog("title") & "") %></strong></td>
                    <td><strong><%= CleanText(rsOverdueLog("username") & "") %></strong></td>
                    <td><code><%= CleanText(rsOverdueLog("staff_number") & "") %></code></td>
                    <td class="col-nowrap"><%= FormatDateTime(rsOverdueLog("borrow_date"), 2) %></td>
                    <td class="col-nowrap"><strong style="color:#e74c3c;"><%= FormatDateTime(rsOverdueLog("due_date"), 2) %></strong></td>
                    <td style="text-align:center;" class="col-nowrap"><span class="status-pill status-pill-red">+<%= rsOverdueLog("days_late") %> Days Late</span></td>
                    <td style="text-align:center;">
                        <div class="action-stack" style="width:115px; margin:0 auto;">
                            <a href="borrowings.asp?filter=overdue" class="action-btn action-btn-edit">Process Return</a>
                        </div>
                    </td>
                </tr>
                <%
                    rsOverdueLog.MoveNext
                Loop
                rsOverdueLog.Close
                %>
            </tbody>
        </table>
    </div>
</div>
<% End If %>

<!-- ==================================================================== -->
<!-- 4. Two-Column Intelligence Section: Categories & Top Books -->
<!-- ==================================================================== -->
<div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap:24px; margin-bottom:28px;">
    
    <!-- Category Distribution Table -->
    <div class="card analytics-card">
        <div class="card-header">
            <h3 style="margin:0; color:#2c3e50; font-size:18px;">📁 Category Breakdown & Volumes</h3>
            <a href="categories.asp" class="btn btn-secondary" style="padding:4px 10px; font-size:12px;">Categories ➔</a>
        </div>
        <div class="analytics-tbl-wrap">
            <table class="analytics-tbl">
                <thead>
                    <tr>
                        <th>Category Name</th>
                        <th style="text-align:center;">Titles</th>
                        <th style="text-align:center;">Shelf Stock</th>
                        <th style="text-align:center;">Active Loans</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If catCount > 0 Then
                        For ci = 0 To catCount - 1
                    %>
                    <tr>
                        <td><strong style="color:#2c3e50;"><%= catNames(ci) %></strong></td>
                        <td style="text-align:center;" class="col-nowrap"><span class="status-pill status-pill-blue" style="font-size:12px; padding:3px 12px;"><%= catTitles(ci) %> Titles</span></td>
                        <td style="text-align:center;" class="col-nowrap"><strong><%= catCopies(ci) %> copies</strong></td>
                        <td style="text-align:center;" class="col-nowrap">
                            <% If catLoans(ci) > 0 Then %>
                                <span class="status-pill status-pill-amber" style="font-size:12px; padding:3px 12px;"><%= catLoans(ci) %> on loan</span>
                            <% Else %>
                                <span style="color:#95a5a6; font-size:12px;">0 on loan</span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        Next
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Top In-Demand Books -->
    <div class="card analytics-card">
        <div class="card-header">
            <h3 style="margin:0; color:#2c3e50; font-size:18px;">🔥 Most Popular / In-Demand Books</h3>
            <a href="books.asp" class="btn btn-secondary" style="padding:4px 10px; font-size:12px;">Catalog ➔</a>
        </div>
        <div class="analytics-tbl-wrap">
            <table class="analytics-tbl">
                <thead>
                    <tr>
                        <th>Book & Author</th>
                        <th style="text-align:right;">Price</th>
                        <th style="text-align:center;">Times Loaned</th>
                        <th style="text-align:center;">Shelf Stock</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If topBookCount > 0 Then
                        For bi = 0 To topBookCount - 1
                    %>
                    <tr>
                        <td>
                            <strong style="color:#2c3e50; font-size:13.5px; display:block; line-height:1.3;"><%= topBookTitles(bi) %></strong>
                            <span style="color:#7f8c8d; font-size:12px; display:block; margin-top:3px;">✍️ <%= topBookAuthors(bi) %></span>
                        </td>
                        <td style="text-align:right;" class="col-nowrap"><strong style="color:#27ae60;">&#8377;<%= FormatNumber(topBookPrices(bi), 2) %></strong></td>
                        <td style="text-align:center;" class="col-nowrap"><span class="status-pill status-pill-amber" style="font-size:12px; padding:3px 12px;">⭐ <%= topBookLoans(bi) %> Loans</span></td>
                        <td style="text-align:center;" class="col-nowrap">
                            <% If topBookLeft(bi) > 0 Then %>
                                <span class="status-pill status-pill-green" style="font-size:12px; padding:3px 12px;"><%= topBookLeft(bi) %> Left</span>
                            <% Else %>
                                <span class="status-pill status-pill-red" style="font-size:12px; padding:3px 12px;">Out of stock</span>
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        Next
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- ==================================================================== -->
<!-- 5. Top Readers & Borrowing Activity Leaders -->
<!-- ==================================================================== -->
<div class="card">
    <div class="card-header">
        <div>
            <h3 style="margin:0; color:#2c3e50; font-size:18px;">👥 Active Borrowers & Staff Activity Leaders</h3>
            <p style="margin:4px 0 0 0; color:#7f8c8d; font-size:13px;">Team members with active loans and successful historical returns.</p>
        </div>
        <a href="borrowings.asp" class="btn btn-secondary" style="padding:5px 12px; font-size:12.5px;">Borrowing Log ➔</a>
    </div>

    <div class="analytics-tbl-wrap">
        <table class="analytics-tbl">
            <thead>
                <tr>
                    <th>Member Name</th>
                    <th style="text-align:center;">Role</th>
                    <th>Staff Number</th>
                    <th style="text-align:center;">Currently Holding</th>
                    <th style="text-align:center;">Successfully Returned</th>
                    <th style="text-align:center;">Reading Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                If Not rsTopBorrowers Is Nothing And rsTopBorrowers.State = 1 Then
                    Do While Not rsTopBorrowers.EOF
                        Dim curHeld
                        curHeld = SafeInt(rsTopBorrowers("current_held"), 0)
                %>
                <tr>
                    <td><strong><%= CleanText(rsTopBorrowers("username") & "") %></strong></td>
                    <td style="text-align:center;">
                        <% If LCase(Trim(rsTopBorrowers("role") & "")) = "admin" Then %>
                            <span class="status-pill status-pill-red" style="font-size:12px; padding:3px 12px;">Admin</span>
                        <% ElseIf LCase(Trim(rsTopBorrowers("role") & "")) = "librarian" Then %>
                            <span class="status-pill status-pill-purple" style="font-size:12px; padding:3px 12px;">Librarian</span>
                        <% Else %>
                            <span class="status-pill status-pill-blue" style="font-size:12px; padding:3px 12px;">Staff</span>
                        <% End If %>
                    </td>
                    <td><code><%= CleanText(rsTopBorrowers("staff_number") & "") %></code></td>
                    <td style="text-align:center;" class="col-nowrap">
                        <% If curHeld > 0 Then %>
                            <span class="status-pill status-pill-amber" style="font-size:12px; padding:3px 12px;"><%= curHeld %> Active</span>
                        <% Else %>
                            <span style="color:#95a5a6; font-size:12.5px;">0 Active</span>
                        <% End If %>
                    </td>
                    <td style="text-align:center;" class="col-nowrap"><strong style="color:#27ae60;"><%= rsTopBorrowers("total_returned") %> Books Returned</strong></td>
                    <td style="text-align:center;" class="col-nowrap">
                        <% If curHeld > 0 Then %>
                            <span style="color:#2980b9; font-weight:600;">📖 Actively Reading</span>
                        <% Else %>
                            <span style="color:#27ae60; font-weight:600;">✅ Clean Record</span>
                        <% End If %>
                    </td>
                </tr>
                <%
                        rsTopBorrowers.MoveNext
                    Loop
                    rsTopBorrowers.Close
                End If

                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>
</div>

<!-- Chart.js Library and Initialization -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script>
document.addEventListener('DOMContentLoaded', function() {
    // 1. System Inventory & Status Breakdown (Pie / Doughnut Chart)
    var ctxPie = document.getElementById('systemPieChart');
    if (ctxPie) {
        new Chart(ctxPie, {
            type: 'doughnut',
            data: {
                labels: ['Available on Shelves', 'Active Borrowings', 'Overdue Loans', 'Returned Records'],
                datasets: [{
                    data: [<%= totalCopiesAvailable %>, <%= totalCurrentlyBorrowed %>, <%= totalOverdue %>, <%= totalReturned %>],
                    backgroundColor: ['#10b981', '#f59e0b', '#ef4444', '#3b82f6'],
                    borderWidth: 2,
                    borderColor: '#ffffff',
                    hoverOffset: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '58%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            boxWidth: 14,
                            padding: 14,
                            font: { family: "'Plus Jakarta Sans', sans-serif", size: 12, weight: '600' },
                            color: '#475569'
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                var total = context.dataset.data.reduce(function(a, b) { return a + b; }, 0);
                                var val = context.parsed;
                                var pct = total > 0 ? Math.round((val / total) * 100) : 0;
                                return ' ' + context.label + ': ' + val + ' (' + pct + '%)';
                            }
                        }
                    }
                }
            }
        });
    }

    // 2. Category Volumes & Circulation (Bar Chart)
    var ctxBar = document.getElementById('categoryBarChart');
    if (ctxBar) {
        var catLabels = [<% For ci = 0 To catCount - 1: Response.Write("""" & Replace(catNames(ci), """", "\""") & """" & IIf(ci < catCount - 1, ", ", "")): Next %>];
        var catShelfStock = [<% For ci = 0 To catCount - 1: Response.Write(catCopies(ci) & IIf(ci < catCount - 1, ", ", "")): Next %>];
        var catActiveLoans = [<% For ci = 0 To catCount - 1: Response.Write(catLoans(ci) & IIf(ci < catCount - 1, ", ", "")): Next %>];

        new Chart(ctxBar, {
            type: 'bar',
            data: {
                labels: catLabels,
                datasets: [
                    {
                        label: 'Shelf Stock (Available)',
                        data: catShelfStock,
                        backgroundColor: '#0284c7',
                        borderRadius: 6,
                        maxBarThickness: 32
                    },
                    {
                        label: 'Active Loans (On Loan)',
                        data: catActiveLoans,
                        backgroundColor: '#f59e0b',
                        borderRadius: 6,
                        maxBarThickness: 32
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: {
                            font: { family: "'Plus Jakarta Sans', sans-serif", size: 11, weight: '600' },
                            color: '#64748b'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: '#f1f5f9' },
                        ticks: {
                            precision: 0,
                            font: { family: "'Plus Jakarta Sans', sans-serif", size: 11 },
                            color: '#64748b'
                        }
                    }
                },
                plugins: {
                    legend: {
                        position: 'top',
                        align: 'end',
                        labels: {
                            boxWidth: 12,
                            padding: 12,
                            font: { family: "'Plus Jakarta Sans', sans-serif", size: 12, weight: '600' },
                            color: '#475569'
                        }
                    }
                }
            }
        });
    }
});
</script>

<% RenderFooter %>
