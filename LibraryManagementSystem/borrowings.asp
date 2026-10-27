<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim conn, errorMessage, successMessage
Dim actionType, bookId, userId, days, todayStr, dueStr, insertSql, borrowId, retDateStr, currentFilter
Dim rsBooks, rsUsers, rsList, sqlList, currentUserId
Dim bStatus, bBorrowDate, bDueDate, bReturnDate, isOverdueRed, rowStyle, fmtBorrowDate, fmtDueDate, fmtReturnDate, rowStatusType
Dim bAuthor, bCopies, uName, uEmail, uStaffNo, uInternalNo, uRole


errorMessage = ""
successMessage = ""
currentFilter = LCase(Trim(Request.QueryString("filter") & ""))
If currentFilter = "" Then currentFilter = "all"
Set conn = GetConnection()

' --------------------------------------------------------------------
' Auto-Update Overdue Status (Mark Overdue if CURRENT_DATE > due_date)
' --------------------------------------------------------------------
On Error Resume Next
conn.Execute("UPDATE Borrowings SET status = 'Overdue' WHERE status = 'Borrowed' AND due_date < CURRENT_DATE()")
On Error GoTo 0

' Handle Form Action (Issue Book / Return Book) - Staff Only
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not IsStaff() Then
        errorMessage = "Only library staff (Admin/Librarian) can issue or process book returns."
    Else
        actionType = Request.Form("action_type")

        If actionType = "issue" Then
            bookId = SafeInt(Request.Form("book_id"), 0)
            userId = SafeInt(Request.Form("user_id"), 0)
            days = SafeInt(Request.Form("days"), 15)

            If bookId = 0 Or userId = 0 Then
                errorMessage = "Please select both a Book and a Staff Member."
            Else
                todayStr = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
                dueStr = Year(DateAdd("d", days, Now)) & "-" & Right("0" & Month(DateAdd("d", days, Now)), 2) & "-" & Right("0" & Day(DateAdd("d", days, Now)), 2)

                insertSql = "INSERT INTO Borrowings (book_id, user_id, borrow_date, due_date, status) VALUES (" & _
                            bookId & ", " & userId & ", " & SQLQuote(todayStr) & ", " & SQLQuote(dueStr) & ", 'Borrowed')"

                On Error Resume Next
                conn.Execute(insertSql)
                ' Decrease copy count
                conn.Execute("UPDATE Books SET copies_available = copies_available - 1 WHERE id = " & bookId & " AND copies_available > 0")
                If Err.Number <> 0 Then
                    errorMessage = "Issue Book Error: " & Err.Description
                Else
                    successMessage = "Book issued successfully for " & days & " days."
                End If
                On Error GoTo 0
            End If

        ElseIf actionType = "return" Then
            borrowId = SafeInt(Request.Form("borrow_id"), 0)
            If borrowId > 0 Then
                retDateStr = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)

                On Error Resume Next
                ' Update borrowing status
                conn.Execute("UPDATE Borrowings SET status = 'Returned', return_date = " & SQLQuote(retDateStr) & " WHERE id = " & borrowId)
                
                ' Increment available book copy count
                conn.Execute("UPDATE Books SET copies_available = copies_available + 1 WHERE id = (SELECT book_id FROM Borrowings WHERE id = " & borrowId & ")")
                If Err.Number <> 0 Then
                    errorMessage = "Return Book Error: " & Err.Description
                Else
                    successMessage = "Book returned successfully."
                End If
                On Error GoTo 0
            End If
        End If
    End If
End If

' Fetch Available Books & Users for Select Dropdowns (Staff view)
If IsStaff() Then
    Set rsBooks = conn.Execute("SELECT b.id, b.title, b.isbn, b.copies_available, COALESCE(a.name, 'Various') AS author_name FROM Books b LEFT JOIN Authors a ON b.author_id = a.id WHERE b.copies_available > 0 ORDER BY b.title ASC")
    Set rsUsers = conn.Execute("SELECT id, username, email, COALESCE(staff_number, '') AS staff_number, role FROM Users ORDER BY username ASC")
End If

' Fetch Borrowing Records List
If IsStaff() Then
    sqlList = "SELECT b.id, bk.title, u.username, b.borrow_date, b.due_date, b.return_date, b.status " & _
              "FROM Borrowings b " & _
              "LEFT JOIN Books bk ON b.book_id = bk.id " & _
              "LEFT JOIN Users u ON b.user_id = u.id " & _
              "ORDER BY b.id DESC"
Else
    currentUserId = SafeInt(Session("UserId"), 0)
    sqlList = "SELECT b.id, bk.title, u.username, b.borrow_date, b.due_date, b.return_date, b.status " & _
              "FROM Borrowings b " & _
              "LEFT JOIN Books bk ON b.book_id = bk.id " & _
              "LEFT JOIN Users u ON b.user_id = u.id " & _
              "WHERE b.user_id = " & currentUserId & " " & _
              "ORDER BY b.id DESC"
End If

Set rsList = conn.Execute(sqlList)

RenderHeader IIf(IsStaff(), "Book Borrowing Management", "My Borrowed Books")
%>

<style>
/* Modern Searchable Dropdown / Combobox Styles */
.searchable-select-wrap {
    position: relative;
    width: 100%;
}
.searchable-input-box {
    position: relative;
    display: flex;
    align-items: center;
}
.searchable-input-box input.search-input {
    width: 100%;
    height: 42px;
    padding: 8px 36px 8px 12px;
    border: 1px solid #ced4da;
    border-radius: 5px;
    font-size: 13.5px;
    background: #fff;
    color: #2c3e50;
    transition: all 0.2s ease;
}
.searchable-input-box input.search-input:focus {
    border-color: #3498db;
    box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.18);
    outline: none;
}
.searchable-input-box .clear-btn {
    position: absolute;
    right: 28px;
    background: transparent;
    border: none;
    cursor: pointer;
    font-size: 13px;
    color: #95a5a6;
    padding: 4px 6px;
    display: none;
    line-height: 1;
}
.searchable-input-box .clear-btn:hover {
    color: #e74c3c;
}
.searchable-input-box .arrow-icon {
    position: absolute;
    right: 11px;
    pointer-events: none;
    font-size: 10px;
    color: #7f8c8d;
    transition: transform 0.2s ease;
}
.searchable-select-wrap.open .arrow-icon {
    transform: rotate(180deg);
}
.searchable-dropdown-list {
    position: absolute;
    top: calc(100% + 4px);
    left: 0;
    right: 0;
    background: #ffffff;
    border: 1px solid #cbd5e0;
    border-radius: 6px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.12);
    max-height: 250px;
    overflow-y: auto;
    z-index: 9999;
    display: none;
}
.searchable-dropdown-list.show {
    display: block;
}
.searchable-item {
    padding: 10px 12px;
    cursor: pointer;
    border-bottom: 1px solid #f1f5f9;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 8px;
    transition: background 0.15s ease;
}
.searchable-item:last-child {
    border-bottom: none;
}
.searchable-item:hover {
    background: #f0f4f8;
}
.searchable-item.selected {
    background: #e8f4fd;
    border-left: 3px solid #3498db;
}
.searchable-item-primary {
    font-weight: 600;
    color: #2c3e50;
    font-size: 13px;
    line-height: 1.3;
}
.searchable-item-sub {
    font-size: 11.5px;
    color: #7f8c8d;
    margin-top: 2px;
}
.searchable-item-badge {
    font-size: 11px;
    font-weight: 600;
    padding: 3px 8px;
    border-radius: 12px;
    white-space: nowrap;
}
.badge-stock {
    background: #e8f8f5;
    color: #16a085;
    border: 1px solid #a3e4d7;
}
.badge-staff {
    background: #ebf5fb;
    color: #2980b9;
    border: 1px solid #aed6f1;
}
.searchable-empty {
    padding: 16px;
    text-align: center;
    color: #95a5a6;
    font-size: 13px;
}
</style>

<% If IsStaff() Then %>
    <div class="card">
        <div class="card-header">
            <h2 style="margin:0; color:#2c3e50;">Issue / Borrow Book</h2>
        </div>

        <% If errorMessage <> "" Then %>
            <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
        <% End If %>

        <% If successMessage <> "" Then %>
            <div class="alert alert-info"><%= CleanText(successMessage) %></div>
        <% End If %>

        <form action="borrowings.asp" method="POST" onsubmit="return validateIssueForm(event);" style="display:grid; grid-template-columns: 2.2fr 2.2fr 1fr 1.1fr; gap:16px; align-items:end;">
            <input type="hidden" name="action_type" value="issue">

            <!-- Searchable Book Selector -->
            <div class="form-group" style="margin-bottom:0;">
                <label for="bookSearchInput" style="display:block; margin-bottom:6px; font-weight:600; color:#2c3e50;">
                    📚 Select Book (Search Title / Author)
                </label>
                <div class="searchable-select-wrap" id="bookSearchableWrap">
                    <input type="hidden" id="book_id" name="book_id" value="">
                    <div class="searchable-input-box">
                        <input type="text" id="bookSearchInput" class="search-input" placeholder="🔍 Type to search book or author..." autocomplete="off">
                        <button type="button" class="clear-btn" id="bookClearBtn" title="Clear selection">✕</button>
                        <span class="arrow-icon">▼</span>
                    </div>
                    <div class="searchable-dropdown-list" id="bookDropdown">
                        <%
                        If Not rsBooks Is Nothing And rsBooks.State = 1 Then
                            Do While Not rsBooks.EOF
                                bAuthor = CleanText(rsBooks("author_name") & "")
                                bCopies = rsBooks("copies_available")
                        %>
                            <div class="searchable-item" 
                                 data-id="<%= rsBooks("id") %>" 
                                 data-title="<%= CleanText(rsBooks("title") & "") %>" 
                                 data-search="<%= LCase(CleanText(rsBooks("title") & " " & bAuthor & " " & rsBooks("isbn"))) %>">
                                <div>
                                    <div class="searchable-item-primary"><%= CleanText(rsBooks("title") & "") %></div>
                                    <div class="searchable-item-sub"><%= bAuthor %> &bull; ISBN: <%= CleanText(rsBooks("isbn") & "") %></div>
                                </div>
                                <span class="searchable-item-badge badge-stock"><%= bCopies %> left</span>
                            </div>
                        <%
                                rsBooks.MoveNext
                            Loop
                        End If
                        %>
                        <div class="searchable-empty" id="bookNoMatches" style="display:none;">
                            No books found matching search.
                        </div>
                    </div>
                </div>
            </div>

            <!-- Searchable Staff Selector -->
            <div class="form-group" style="margin-bottom:0;">
                <label for="userSearchInput" style="display:block; margin-bottom:6px; font-weight:600; color:#2c3e50;">
                    👤 Select Staff Member (Search Name / ID)
                </label>
                <div class="searchable-select-wrap" id="userSearchableWrap">
                    <input type="hidden" id="user_id" name="user_id" value="">
                    <div class="searchable-input-box">
                        <input type="text" id="userSearchInput" class="search-input" placeholder="🔍 Type to search staff name, ID, or email..." autocomplete="off">
                        <button type="button" class="clear-btn" id="userClearBtn" title="Clear selection">✕</button>
                        <span class="arrow-icon">▼</span>
                    </div>
                    <div class="searchable-dropdown-list" id="userDropdown">
                        <%
                        If Not rsUsers Is Nothing And rsUsers.State = 1 Then
                            Do While Not rsUsers.EOF
                                uName = CleanText(rsUsers("username") & "")
                                uEmail = CleanText(rsUsers("email") & "")
                                uStaffNo = CleanText(rsUsers("staff_number") & "")
                                uRole = CleanText(rsUsers("role") & "")
                        %>
                            <div class="searchable-item" 
                                 data-id="<%= rsUsers("id") %>" 
                                 data-title="<%= uName %>" 
                                 data-search="<%= LCase(uName & " " & uEmail & " " & uStaffNo & " " & uRole) %>">
                                <div>
                                    <div class="searchable-item-primary"><%= uName %></div>
                                    <div class="searchable-item-sub"><%= uEmail %> &bull; Staff No: <%= uStaffNo %></div>
                                </div>
                                <span class="searchable-item-badge badge-staff"><%= uRole %></span>
                            </div>
                        <%
                                rsUsers.MoveNext
                            Loop
                        End If
                        %>
                        <div class="searchable-empty" id="userNoMatches" style="display:none;">
                            No staff members found matching search.
                        </div>
                    </div>
                </div>
            </div>

            <div class="form-group" style="margin-bottom:0;">
                <label for="days" style="display:block; margin-bottom:6px; font-weight:600; color:#2c3e50;">Days Due (Standard: 15)</label>
                <input type="number" id="days" name="days" value="15" min="1" required style="height:42px;">
            </div>

            <div>
                <button type="submit" class="btn btn-success" style="width:100%; height:42px; font-weight:600; font-size:14px; display:flex; align-items:center; justify-content:center; gap:6px;">
                    <span>✓</span> Issue Book
                </button>
            </div>
        </form>
    </div>
<% End If %>

<div class="card">
    <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <div>
            <h3 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Active & Past Borrowing Log", "My Borrowed Books History") %></h3>
            <% If IsStaff() Then %>
                <p style="margin:4px 0 0 0; color:#e74c3c; font-size:13px; font-weight:600;">🚨 Note: Borrowings reaching day 16 (past 15-day return limit) are highlighted in bright red.</p>
            <% End If %>
        </div>
        <div style="display:flex; gap:10px; align-items:center;">
            <button type="button" id="btn-overdue-header" class="btn btn-danger" onclick="filterBorrowings('overdue', document.getElementById('btn-overdue'))" style="background:#e74c3c; color:white; font-weight:bold; padding:8px 16px; border-radius:4px; font-size:13.5px; box-shadow:0 2px 4px rgba(231,76,60,0.3); cursor:pointer;">
                🚨 Filter Overdue Borrowings (<span id="header-overdue-count">0</span>)
            </button>
            <button type="button" id="btn-reset-header" class="btn btn-secondary" onclick="filterBorrowings('all', document.getElementById('btn-all'))" style="display:none; padding:8px 14px; font-size:13.5px; cursor:pointer;">
                Show All (<span id="header-all-count">0</span>)
            </button>
        </div>
    </div>

    <!-- Filter Buttons & Quick Search Toolbar -->
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:15px; padding-bottom:12px; border-bottom:1px solid #f1f1f1;">
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <button type="button" class="btn filter-btn <%= IIf(currentFilter="all", "btn-primary", "btn-secondary") %>" onclick="filterBorrowings('all', this)" id="btn-all">
                All (<span id="count-all">0</span>)
            </button>
            <button type="button" class="btn filter-btn <%= IIf(currentFilter="overdue", "btn-danger", "btn-secondary") %>" style="<%= IIf(currentFilter="overdue", "background:#e74c3c;color:#fff;font-weight:bold;", "") %>" onclick="filterBorrowings('overdue', this)" id="btn-overdue">
                🚨 Overdue (<span id="count-overdue">0</span>)
            </button>
            <button type="button" class="btn filter-btn <%= IIf(currentFilter="borrowed", "btn-warning", "btn-secondary") %>" onclick="filterBorrowings('borrowed', this)" id="btn-borrowed">
                Borrowed (<span id="count-borrowed">0</span>)
            </button>
            <button type="button" class="btn filter-btn <%= IIf(currentFilter="returned", "btn-success", "btn-secondary") %>" onclick="filterBorrowings('returned', this)" id="btn-returned">
                Returned (<span id="count-returned">0</span>)
            </button>
        </div>

        <div style="min-width:240px; flex:1; max-width:320px;">
            <input type="text" id="borrowSearch" onkeyup="searchBorrowings()" placeholder="🔍 Search title, staff, or #ID..." style="width:100%; padding:7px 12px; border:1px solid #ccc; border-radius:4px; font-size:13.5px;">
        </div>
    </div>

    <table id="borrowingsTable">
        <thead>
            <tr>
                <th>ID</th>
                <th>Book Title</th>
                <% If IsStaff() Then %>
                    <th>Staff Name</th>
                <% End If %>
                <th>Borrowed On</th>
                <th>Due On (Day 15)</th>
                <th>Returned On</th>
                <th>Status</th>
                <% If IsStaff() Then %>
                    <th>Action</th>
                <% End If %>
            </tr>
        </thead>
        <tbody>
            <%
            If rsList Is Nothing Or rsList.State = 0 Or rsList.EOF Then
            %>
                <tr>
                    <td colspan="<%= IIf(IsStaff(), 8, 6) %>" style="text-align:center; color:#95a5a6;">No borrowing transactions found.</td>
                </tr>
            <%
            Else
                Do While Not rsList.EOF
                    bStatus = rsList("status") & ""
                    bBorrowDate = rsList("borrow_date")
                    bDueDate = rsList("due_date")
                    bReturnDate = rsList("return_date")
                    
                    isOverdueRed = False
                    If bStatus <> "Returned" And Not IsNull(bDueDate) And IsDate(bDueDate) Then
                        If Date > CDate(bDueDate) Then
                            isOverdueRed = True
                        End If
                    End If
                    
                    If isOverdueRed Then
                        rowStatusType = "overdue"
                    ElseIf bStatus = "Returned" Then
                        rowStatusType = "returned"
                    Else
                        rowStatusType = "borrowed"
                    End If

                    If isOverdueRed Then
                        rowStyle = "style='background-color:#fce4e4; color:#721c24; border-left: 6px solid #e74c3c; font-weight:bold;'"
                    Else
                        rowStyle = ""
                    End If

                    If Not IsNull(bBorrowDate) And IsDate(bBorrowDate) Then
                        fmtBorrowDate = FormatDateTime(bBorrowDate, 2)
                    Else
                        fmtBorrowDate = "-"
                    End If

                    If Not IsNull(bDueDate) And IsDate(bDueDate) Then
                        fmtDueDate = FormatDateTime(bDueDate, 2)
                    Else
                        fmtDueDate = "-"
                    End If

                    If Not IsNull(bReturnDate) And IsDate(bReturnDate) Then
                        fmtReturnDate = FormatDateTime(bReturnDate, 2)
                    Else
                        fmtReturnDate = "-"
                    End If
            %>
                <tr <%= rowStyle %> data-status="<%= rowStatusType %>" class="borrow-row">
                    <td>#<%= rsList("id") %></td>
                    <td><strong style="color:#2c3e50;"><%= CleanText(rsList("title") & "") %></strong></td>
                    <% If IsStaff() Then %>
                        <td><%= CleanText(rsList("username") & "") %></td>
                    <% End If %>
                    <td><%= fmtBorrowDate %></td>
                    <td><%= fmtDueDate %></td>
                    <td>
                        <% If fmtReturnDate <> "-" Then %>
                            <%= fmtReturnDate %>
                        <% Else %>
                            <span style="color:#95a5a6;">-</span>
                        <% End If %>
                    </td>
                    <td>
                        <% If isOverdueRed Then %>
                            <span class="status-pill status-pill-red">🚨 Overdue</span>
                        <% ElseIf bStatus = "Returned" Then %>
                            <span class="status-pill status-pill-green">Returned</span>
                        <% Else %>
                            <span class="status-pill status-pill-amber">Borrowed</span>
                        <% End If %>
                    </td>
                    <% If IsStaff() Then %>
                        <td>
                            <% If bStatus <> "Returned" Then %>
                                <div class="action-stack" style="width:115px;">
                                    <form action="borrowings.asp" method="POST" style="margin:0;">
                                        <input type="hidden" name="action_type" value="return">
                                        <input type="hidden" name="borrow_id" value="<%= rsList("id") %>">
                                        <button type="submit" class="action-btn action-btn-request">Return Book</button>
                                    </form>
                                </div>
                            <% Else %>
                                <span class="action-text-muted">Completed</span>
                            <% End If %>
                        </td>
                    <% End If %>
                </tr>
            <%
                    rsList.MoveNext
                Loop
                rsList.Close
            End If

            conn.Close
            Set conn = Nothing
            %>
        </tbody>
    </table>

    <div id="noMatchRow" style="display:none; text-align:center; padding:30px; color:#95a5a6; font-size:14.5px;">
        <span style="font-size:24px;">🔍</span><br>
        <strong>No borrowings found matching the selected filter.</strong>
    </div>
</div>

<script>
function filterBorrowings(status, btn) {
    document.querySelectorAll('.filter-btn').forEach(function(b) {
        b.className = 'btn btn-secondary filter-btn';
        b.style.fontWeight = 'normal';
        b.style.background = '';
        b.style.color = '';
    });

    if (btn) {
        if (status === 'overdue') {
            btn.className = 'btn btn-danger filter-btn';
            btn.style.background = '#e74c3c';
            btn.style.color = '#fff';
            btn.style.fontWeight = 'bold';
        } else if (status === 'borrowed') {
            btn.className = 'btn btn-warning filter-btn';
            btn.style.fontWeight = 'bold';
        } else if (status === 'returned') {
            btn.className = 'btn btn-success filter-btn';
            btn.style.fontWeight = 'bold';
        } else {
            btn.className = 'btn btn-primary filter-btn';
            btn.style.fontWeight = 'bold';
        }
    }

    var btnReset = document.getElementById('btn-reset-header');
    if (btnReset) {
        btnReset.style.display = (status !== 'all') ? 'inline-block' : 'none';
    }

    var rows = document.querySelectorAll('.borrow-row');
    var visible = 0;
    rows.forEach(function(r) {
        var rStatus = r.getAttribute('data-status');
        if (status === 'all' || rStatus === status) {
            r.style.display = '';
            visible++;
        } else {
            r.style.display = 'none';
        }
    });

    var noMatch = document.getElementById('noMatchRow');
    if (noMatch) {
        noMatch.style.display = (visible === 0 && rows.length > 0) ? 'block' : 'none';
    }
}

function searchBorrowings() {
    var query = document.getElementById('borrowSearch').value.toLowerCase();
    var rows = document.querySelectorAll('.borrow-row');
    var visible = 0;
    rows.forEach(function(r) {
        var text = r.innerText.toLowerCase();
        if (text.indexOf(query) > -1) {
            r.style.display = '';
            visible++;
        } else {
            r.style.display = 'none';
        }
    });

    var noMatch = document.getElementById('noMatchRow');
    if (noMatch) {
        noMatch.style.display = (visible === 0 && rows.length > 0) ? 'block' : 'none';
    }
}

// --- Searchable Combobox Component Logic ---
function initSearchableSelect(wrapId, inputId, hiddenId, clearBtnId, dropdownId, noMatchesId) {
    var wrap = document.getElementById(wrapId);
    var input = document.getElementById(inputId);
    var hidden = document.getElementById(hiddenId);
    var clearBtn = document.getElementById(clearBtnId);
    var dropdown = document.getElementById(dropdownId);
    var noMatches = document.getElementById(noMatchesId);
    if (!wrap || !input || !hidden || !dropdown) return;

    var items = dropdown.querySelectorAll('.searchable-item');

    function openDropdown() {
        document.querySelectorAll('.searchable-dropdown-list').forEach(function(d) {
            if (d !== dropdown) {
                d.classList.remove('show');
                if (d.parentElement) d.parentElement.classList.remove('open');
            }
        });
        dropdown.classList.add('show');
        wrap.classList.add('open');
        filterItems();
    }

    function filterItems() {
        var query = input.value.trim().toLowerCase();
        var visibleCount = 0;
        items.forEach(function(item) {
            var searchData = item.getAttribute('data-search') || '';
            if (!query || searchData.indexOf(query) > -1) {
                item.style.display = 'flex';
                visibleCount++;
            } else {
                item.style.display = 'none';
            }
        });
        if (noMatches) {
            noMatches.style.display = (visibleCount === 0) ? 'block' : 'none';
        }
    }

    input.addEventListener('focus', function() {
        openDropdown();
    });

    input.addEventListener('click', function() {
        openDropdown();
    });

    input.addEventListener('input', function() {
        hidden.value = '';
        items.forEach(function(it) { it.classList.remove('selected'); });
        if (clearBtn) clearBtn.style.display = input.value ? 'block' : 'none';
        openDropdown();
        filterItems();
    });

    items.forEach(function(item) {
        item.addEventListener('click', function(e) {
            e.stopPropagation();
            var id = item.getAttribute('data-id');
            var title = item.getAttribute('data-title');
            hidden.value = id;
            input.value = title;
            items.forEach(function(it) { it.classList.remove('selected'); });
            item.classList.add('selected');
            if (clearBtn) clearBtn.style.display = 'block';
            dropdown.classList.remove('show');
            wrap.classList.remove('open');
            input.style.borderColor = '#2ecc71';
            setTimeout(function() { input.style.borderColor = ''; }, 1200);
        });
    });

    if (clearBtn) {
        clearBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            hidden.value = '';
            input.value = '';
            clearBtn.style.display = 'none';
            items.forEach(function(it) { it.classList.remove('selected'); });
            filterItems();
            input.focus();
        });
    }
}

// Global click outside listener
document.addEventListener('click', function(e) {
    if (!e.target.closest('.searchable-select-wrap')) {
        document.querySelectorAll('.searchable-dropdown-list').forEach(function(d) {
            d.classList.remove('show');
        });
        document.querySelectorAll('.searchable-select-wrap').forEach(function(w) {
            w.classList.remove('open');
            var inp = w.querySelector('.search-input');
            var hid = w.querySelector('input[type="hidden"]');
            var clr = w.querySelector('.clear-btn');
            if (inp && hid && !hid.value) {
                inp.value = '';
                if (clr) clr.style.display = 'none';
            }
        });
    }
});

function validateIssueForm(e) {
    var bId = document.getElementById('book_id').value;
    var uId = document.getElementById('user_id').value;
    if (!bId || bId === '' || bId === '0') {
        alert('⚠️ Please search and choose a Book from the dropdown list before issuing.');
        var bInp = document.getElementById('bookSearchInput');
        if (bInp) { bInp.focus(); bInp.click(); }
        if (e && e.preventDefault) e.preventDefault();
        return false;
    }
    if (!uId || uId === '' || uId === '0') {
        alert('⚠️ Please search and choose a Staff Member from the dropdown list before issuing.');
        var uInp = document.getElementById('userSearchInput');
        if (uInp) { uInp.focus(); uInp.click(); }
        if (e && e.preventDefault) e.preventDefault();
        return false;
    }
    return true;
}

document.addEventListener('DOMContentLoaded', function() {
    initSearchableSelect('bookSearchableWrap', 'bookSearchInput', 'book_id', 'bookClearBtn', 'bookDropdown', 'bookNoMatches');
    initSearchableSelect('userSearchableWrap', 'userSearchInput', 'user_id', 'userClearBtn', 'userDropdown', 'userNoMatches');

    var rows = document.querySelectorAll('.borrow-row');
    var allCount = rows.length;
    var overdueCount = 0;
    var borrowedCount = 0;
    var returnedCount = 0;

    rows.forEach(function(r) {
        var s = r.getAttribute('data-status');
        if (s === 'overdue') overdueCount++;
        else if (s === 'borrowed') borrowedCount++;
        else if (s === 'returned') returnedCount++;
    });

    if (document.getElementById('count-all')) document.getElementById('count-all').innerText = allCount;
    if (document.getElementById('count-overdue')) document.getElementById('count-overdue').innerText = overdueCount;
    if (document.getElementById('count-borrowed')) document.getElementById('count-borrowed').innerText = borrowedCount;
    if (document.getElementById('count-returned')) document.getElementById('count-returned').innerText = returnedCount;

    if (document.getElementById('header-overdue-count')) document.getElementById('header-overdue-count').innerText = overdueCount;
    if (document.getElementById('header-all-count')) document.getElementById('header-all-count').innerText = allCount;

    <% If currentFilter <> "all" Then %>
        var activeBtn = document.getElementById('btn-<%= currentFilter %>');
        if (activeBtn) {
            filterBorrowings('<%= currentFilter %>', activeBtn);
        }
    <% End If %>
});
</script>

<% RenderFooter %>
