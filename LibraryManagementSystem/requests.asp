<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim conn, errorMessage, successMessage, msg
Dim rsExpired, expReqId, expBookId
Dim actionType, requestId, rsReq, reqBookId, reqUserId, reqStart, reqEnd, insertBorrowSql
Dim rsList, sqlList, currentUserId
Dim reqStatus, approvedAt, hoursRemaining, hoursPassed

errorMessage = ""
successMessage = ""
msg = Request.QueryString("msg")
If msg <> "" Then successMessage = msg

Set conn = GetConnection()

' --------------------------------------------------------------------
' Auto-Cleanup: Expire approved requests older than 48 hours & restore book copy
' --------------------------------------------------------------------
On Error Resume Next
Set rsExpired = conn.Execute("SELECT id, book_id FROM BookRequests WHERE status = 'Approved' AND approved_at IS NOT NULL AND approved_at < NOW() - INTERVAL 48 HOUR")
If Not (rsExpired Is Nothing Or rsExpired.State = 0) Then
    Do While Not rsExpired.EOF
        expReqId = rsExpired("id")
        expBookId = rsExpired("book_id")
        
        ' 1. Restore book inventory copy
        conn.Execute("UPDATE Books SET copies_available = copies_available + 1 WHERE id = " & expBookId)
        ' 2. Mark request as Expired
        conn.Execute("UPDATE BookRequests SET status = 'Expired' WHERE id = " & expReqId)
        
        rsExpired.MoveNext
    Loop
    rsExpired.Close
End If
On Error GoTo 0

' --------------------------------------------------------------------
' Handle Actions (Approve / Reject / Mark Collected) - Staff / Admin / Librarian
' --------------------------------------------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not IsStaff() Then
        errorMessage = "Only library staff can approve, reject, or process collection for book requests."
    Else
        actionType = Request.Form("action_type")
        requestId = SafeInt(Request.Form("request_id"), 0)

        If requestId > 0 Then
            Set rsReq = conn.Execute("SELECT * FROM BookRequests WHERE id = " & requestId)

            If rsReq.EOF Then
                errorMessage = "Request not found."
            Else
                reqBookId = rsReq("book_id")
                reqUserId = rsReq("user_id")
                reqStart = rsReq("start_date")
                reqEnd = rsReq("end_date")

                If actionType = "approve" Then
                    On Error Resume Next
                    ' 1. Decrement available book copy (reserve for 48h)
                    conn.Execute("UPDATE Books SET copies_available = copies_available - 1 WHERE id = " & reqBookId & " AND copies_available > 0")

                    ' 2. Update Request status to Approved with timestamp
                    conn.Execute("UPDATE BookRequests SET status = 'Approved', approved_at = NOW() WHERE id = " & requestId)

                    If Err.Number <> 0 Then
                        errorMessage = "Error approving request: " & Err.Description
                    Else
                        successMessage = "Book request approved! Staff member has 48 hours to collect the book."
                    End If
                    On Error GoTo 0

                ElseIf actionType = "collect" Then
                    On Error Resume Next
                    ' 1. Create Borrowing record upon collection
                    insertBorrowSql = "INSERT INTO Borrowings (book_id, user_id, borrow_date, due_date, status) VALUES (" & _
                                      reqBookId & ", " & reqUserId & ", " & FormatSQLDate(reqStart) & ", " & FormatSQLDate(reqEnd) & ", 'Borrowed')"
                    conn.Execute(insertBorrowSql)

                    ' 2. Update Request status to Collected
                    conn.Execute("UPDATE BookRequests SET status = 'Collected' WHERE id = " & requestId)

                    If Err.Number <> 0 Then
                        errorMessage = "Error processing collection: " & Err.Description
                    Else
                        successMessage = "Book successfully collected! Borrowing transaction issued."
                    End If
                    On Error GoTo 0

                ElseIf actionType = "reject" Then
                    On Error Resume Next
                    conn.Execute("UPDATE BookRequests SET status = 'Rejected' WHERE id = " & requestId)
                    If Err.Number <> 0 Then
                        errorMessage = "Error rejecting request: " & Err.Description
                    Else
                        successMessage = "Book request has been rejected."
                    End If
                    On Error GoTo 0
                End If
            End If
            rsReq.Close
        End If
    End If
End If

' --------------------------------------------------------------------
' Compute Summary Cards Statistics (Directly from Real Database Data)
' --------------------------------------------------------------------
Dim totalReqCount, pendingCount, readyCount, expiredCount
totalReqCount = 0
pendingCount = 0
readyCount = 0
expiredCount = 0

Dim rsStats, statsFilter
statsFilter = ""
If Not IsStaff() Then
    statsFilter = "WHERE user_id = " & SafeInt(Session("UserId"), 0)
End If

On Error Resume Next
Set rsStats = conn.Execute("SELECT COUNT(*) AS total_cnt, " & _
                           "SUM(CASE WHEN status = 'Pending' THEN 1 ELSE 0 END) AS pending_cnt, " & _
                           "SUM(CASE WHEN status = 'Approved' THEN 1 ELSE 0 END) AS ready_cnt, " & _
                           "SUM(CASE WHEN status = 'Expired' THEN 1 ELSE 0 END) AS expired_cnt " & _
                           "FROM BookRequests " & statsFilter)
If Not (rsStats Is Nothing Or rsStats.State = 0) Then
    If Not rsStats.EOF Then
        totalReqCount = SafeInt(rsStats("total_cnt"), 0)
        pendingCount = SafeInt(rsStats("pending_cnt"), 0)
        readyCount = SafeInt(rsStats("ready_cnt"), 0)
        expiredCount = SafeInt(rsStats("expired_cnt"), 0)
    End If
    rsStats.Close
    Set rsStats = Nothing
End If
On Error GoTo 0

' Fetch Requests List with all Metadata Fields
sqlList = "SELECT br.id AS req_id, br.book_id, " & _
          "COALESCE(br.book_name, bk.title) AS book_name, " & _
          "COALESCE(br.author_name, a.name, 'Unknown Author') AS author_name, " & _
          "COALESCE(br.price, bk.price, 0.00) AS price, " & _
          "COALESCE(br.staff_name, u.username) AS staff_name, " & _
          "COALESCE(br.staff_number, u.staff_number, '9876543210') AS staff_number, " & _
          "COALESCE(br.staff_internal_number, u.staff_internal_number, 'BEL-EXT-101') AS staff_internal_number, " & _
          "br.start_date, br.end_date, br.status, br.approved_at, br.created_at " & _
          "FROM BookRequests br " & _
          "LEFT JOIN Books bk ON br.book_id = bk.id " & _
          "LEFT JOIN Authors a ON bk.author_id = a.id " & _
          "LEFT JOIN Users u ON br.user_id = u.id "

If Not IsStaff() Then
    currentUserId = SafeInt(Session("UserId"), 0)
    sqlList = sqlList & "WHERE br.user_id = " & currentUserId & " "
End If

sqlList = sqlList & "ORDER BY br.id DESC"

Set rsList = conn.Execute(sqlList)

RenderHeader IIf(IsStaff(), "Staff Book Requests & Collection Approval", "My Book Requests")
%>

<style>
/* Enterprise Summary Metric Cards */
.summary-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 18px;
    margin-bottom: 24px;
}
.summary-card {
    display: flex;
    align-items: center;
    gap: 16px;
    padding: 18px 20px;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.summary-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}
.sum-blue { background-color: #f0f7ff; border-color: #d0e7ff; }
.sum-amber { background-color: #fffbeb; border-color: #fef3c7; }
.sum-green { background-color: #f0fdf4; border-color: #dcfce7; }
.sum-red { background-color: #fef2f2; border-color: #fee2e2; }

.sum-icon-box {
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.sum-icon-blue { background-color: #dbeafe; color: #0284c7; }
.sum-icon-amber { background-color: #fef3c7; color: #d97706; }
.sum-icon-green { background-color: #dcfce7; color: #16a34a; }
.sum-icon-red { background-color: #fee2e2; color: #dc2626; }

.sum-count {
    font-size: 28px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1;
    margin-bottom: 4px;
}
.sum-label {
    font-size: 13.5px;
    font-weight: 700;
    color: #334155;
    line-height: 1.2;
}
.sum-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 2px;
}

/* Toolbar: Search & Filters */
.filter-toolbar {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 22px;
    flex-wrap: wrap;
}
.search-input-wrap {
    flex: 1;
    min-width: 280px;
    position: relative;
    display: flex;
    align-items: center;
}
.search-icon {
    position: absolute;
    left: 14px;
    color: #94a3b8;
    pointer-events: none;
}
.filter-search-input {
    width: 100%;
    height: 44px;
    padding: 0 16px 0 42px;
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    font-size: 14px;
    color: #0f172a;
    background: #ffffff;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
}
.filter-search-input:focus {
    outline: none;
    border-color: #0077f6;
    box-shadow: 0 0 0 3px rgba(0, 119, 246, 0.12);
}
.filter-select {
    height: 44px;
    padding: 0 32px 0 14px;
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    font-size: 13.5px;
    font-weight: 600;
    color: #334155;
    background: #ffffff;
    cursor: pointer;
    min-width: 160px;
    transition: border-color 0.15s ease;
}
.filter-select:focus {
    outline: none;
    border-color: #0077f6;
}

/* Requests Table Styling */
.requests-table-wrap {
    overflow-x: auto;
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
}
.requests-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
    margin: 0;
}
.requests-table th {
    background-color: #f8fafc;
    color: #475569;
    font-size: 13px;
    font-weight: 700;
    padding: 14px 16px;
    border-bottom: 1.5px solid #e2e8f0;
    white-space: nowrap;
    text-align: left;
}
.requests-table td {
    padding: 14px 16px;
    border-bottom: 1px solid #f1f5f9;
    vertical-align: middle !important;
}
.requests-table tr:hover {
    background-color: #fcfdfe;
}

/* Side-by-Side Action Buttons */
.action-cell {
    white-space: nowrap;
}
.action-btn-row {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}
.btn-table-approve {
    background-color: #16a34a;
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    transition: background-color 0.15s ease, transform 0.1s ease;
}
.btn-table-approve:hover {
    background-color: #15803d;
    transform: translateY(-1px);
}
.btn-table-reject {
    background-color: #ef4444;
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    transition: background-color 0.15s ease, transform 0.1s ease;
}
.btn-table-reject:hover {
    background-color: #dc2626;
    transform: translateY(-1px);
}
.btn-table-collect {
    background-color: #16a34a;
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
}
.action-muted-label {
    color: #94a3b8;
    font-size: 13px;
    font-weight: 600;
}

/* Pagination Footer */
.table-footer-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 16px 4px 4px 4px;
    font-size: 13px;
    color: #64748b;
    flex-wrap: wrap;
    gap: 12px;
}
.pagination-controls {
    display: inline-flex;
    align-items: center;
    gap: 4px;
}
.page-btn {
    width: 32px;
    height: 32px;
    border-radius: 6px;
    border: 1px solid #e2e8f0;
    background: #ffffff;
    color: #475569;
    font-size: 13px;
    font-weight: 600;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    transition: all 0.15s ease;
}
.page-btn:hover {
    background: #f1f5f9;
    color: #0f172a;
}
.page-btn.active {
    background: #0077f6;
    color: #ffffff;
    border-color: #0077f6;
}

@media (max-width: 1024px) {
    .summary-grid {
        grid-template-columns: repeat(2, 1fr);
    }
}
@media (max-width: 640px) {
    .summary-grid {
        grid-template-columns: 1fr;
    }
}
</style>

<div class="card">
    <!-- Page Header matching screenshot -->
    <div class="card-header">
        <div>
            <h1 style="margin:0; font-size:24px; font-weight:800; color:#0f172a;"><%= IIf(IsStaff(), "Staff Book Requests & Collection Approval", "My Book Requests") %></h1>
            <p style="margin:5px 0 0 0; color:#64748b; font-size:13.5px;">Approved requests must be collected within <strong>48 hours</strong>; otherwise they auto-expire and return to inventory.</p>
        </div>
        <a href="request_book.asp" class="btn btn-success">+ Request a Book</a>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
    <% End If %>

    <% If successMessage <> "" Then %>
        <div class="alert alert-info"><%= CleanText(successMessage) %></div>
    <% End If %>

    <!-- 4 Summary Metric Cards (Directly computed from BookRequests data) -->
    <div class="summary-grid">
        <div class="summary-card sum-blue">
            <div class="sum-icon-box sum-icon-blue">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
            </div>
            <div>
                <div class="sum-count"><%= totalReqCount %></div>
                <div class="sum-label">Total Requests</div>
                <div class="sum-sub">All time</div>
            </div>
        </div>

        <div class="summary-card sum-amber">
            <div class="sum-icon-box sum-icon-amber">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            </div>
            <div>
                <div class="sum-count" style="color:#d97706;"><%= pendingCount %></div>
                <div class="sum-label" style="color:#b45309;">Pending Approval</div>
                <div class="sum-sub">Awaiting action</div>
            </div>
        </div>

        <div class="summary-card sum-green">
            <div class="sum-icon-box sum-icon-green">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            </div>
            <div>
                <div class="sum-count" style="color:#16a34a;"><%= readyCount %></div>
                <div class="sum-label" style="color:#15803d;">Ready for Collection</div>
                <div class="sum-sub">Approved (within 48h)</div>
            </div>
        </div>

        <div class="summary-card sum-red">
            <div class="sum-icon-box sum-icon-red">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
            </div>
            <div>
                <div class="sum-count" style="color:#dc2626;"><%= expiredCount %></div>
                <div class="sum-label" style="color:#b91c1c;">Expired</div>
                <div class="sum-sub">Auto-expired</div>
            </div>
        </div>
    </div>

    <!-- Toolbar: Search & Filter Controls -->
    <div class="filter-toolbar">
        <div class="search-input-wrap">
            <svg class="search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            <input type="text" id="requestsSearchInput" class="filter-search-input" placeholder="Search staff name, book title, or request ID..." onkeyup="filterRequestsTable()">
        </div>
        <select id="requestsStatusSelect" class="filter-select" onchange="filterRequestsTable()">
            <option value="">All Statuses</option>
            <option value="pending">Pending Approval</option>
            <option value="approved">Ready for Collection (Approved)</option>
            <option value="collected">Collected / Issued</option>
            <option value="rejected">Rejected</option>
            <option value="expired">Expired</option>
        </select>
        <select id="requestsDateSelect" class="filter-select" onchange="filterRequestsTable()">
            <option value="">Select date range</option>
            <option value="all">All Dates</option>
            <option value="today">Today</option>
            <option value="week">Past 7 Days</option>
            <option value="month">Past 30 Days</option>
        </select>
    </div>

    <!-- Requests Table with exact columns and styling -->
    <div class="requests-table-wrap">
        <table class="requests-table" id="requestsTable">
            <thead>
                <tr>
                    <th>Request ID<br><span style="font-weight:500; font-size:11.5px; color:#94a3b8;">Book ID</span></th>
                    <th>Staff Name</th>
                    <th>Staff Number</th>
                    <th>Book Title</th>
                    <th>Author</th>
                    <th>Price (&#8377;)</th>
                    <th>Borrow Dates</th>
                    <th>Status &amp; 48h Window</th>
                    <% If IsStaff() Then %>
                        <th style="min-width:140px;">Action</th>
                    <% End If %>
                </tr>
            </thead>
            <tbody>
                <%
                Dim rowCount
                rowCount = 0
                If rsList Is Nothing Or rsList.State = 0 Or rsList.EOF Then
                %>
                    <tr id="noRowsFound">
                        <td colspan="<%= IIf(IsStaff(), 9, 8) %>" style="text-align:center; color:#94a3b8; padding:32px 16px;">No book requests found.</td>
                    </tr>
                <%
                Else
                    Do While Not rsList.EOF
                        rowCount = rowCount + 1
                        reqStatus = rsList("status")
                        approvedAt = rsList("approved_at")
                        hoursRemaining = 48
                        
                        If reqStatus = "Approved" And Not IsNull(approvedAt) Then
                            hoursPassed = DateDiff("h", approvedAt, Now())
                            hoursRemaining = 48 - hoursPassed
                            If hoursRemaining < 0 Then hoursRemaining = 0
                        End If

                        ' Extract transliteration if book title has parentheses
                        Dim rawTitle, titleMain, titleSub, parenPos
                        rawTitle = CleanText(rsList("book_name") & "")
                        parenPos = InStr(rawTitle, "(")
                        If parenPos > 1 Then
                            titleMain = Trim(Left(rawTitle, parenPos - 1))
                            titleSub = Trim(Mid(rawTitle, parenPos))
                        Else
                            titleMain = rawTitle
                            titleSub = ""
                        End If

                        ' Author transliteration
                        Dim rawAuthor, authorMain, authorSub, aParenPos
                        rawAuthor = CleanText(rsList("author_name") & "")
                        aParenPos = InStr(rawAuthor, "(")
                        If aParenPos > 1 Then
                            authorMain = Trim(Left(rawAuthor, aParenPos - 1))
                            authorSub = Trim(Mid(rawAuthor, aParenPos))
                        Else
                            authorMain = rawAuthor
                            authorSub = ""
                        End If
                %>
                    <tr class="req-row" data-status="<%= LCase(reqStatus) %>">
                        <td>
                            <strong style="color:#0f172a; font-size:14px;">#<%= rsList("req_id") %></strong><br>
                            <span style="color:#94a3b8; font-size:12px;">Book #<%= rsList("book_id") %></span>
                        </td>
                        <td><strong style="color:#0f172a;"><%= CleanText(rsList("staff_name") & "") %></strong></td>
                        <td><span style="color:#475569; font-weight:600;"><%= CleanText(rsList("staff_number") & "") %></span></td>
                        <td>
                            <strong style="color:#0d9488; font-size:14px; display:block;"><%= titleMain %></strong>
                            <% If titleSub <> "" Then %>
                                <span style="color:#0f172a; font-size:12px; display:block; margin-top:2px;"><%= titleSub %></span>
                            <% End If %>
                        </td>
                        <td>
                            <span style="color:#1e293b; font-size:13px; display:block;"><%= authorMain %></span>
                            <% If authorSub <> "" Then %>
                                <span style="color:#64748b; font-size:12px; display:block; margin-top:2px;"><%= authorSub %></span>
                            <% End If %>
                        </td>
                        <td><strong style="color:#16a34a; font-size:14px;">&#8377;<%= FormatNumber(SafeFloat(rsList("price"), 0), 2) %></strong></td>
                        <td style="white-space:nowrap; color:#475569; font-size:12.5px;">
                            <%= FormatDateTime(rsList("start_date"), 2) %><br>
                            <span style="color:#94a3b8;">to</span> <%= FormatDateTime(rsList("end_date"), 2) %>
                        </td>
                        <td>
                            <% If reqStatus = "Approved" Then %>
                                <span class="status-pill status-pill-approved">Ready for Collection</span>
                                <div style="font-size:11.5px; color:#15803d; font-weight:700; margin-top:3px;">Collect in <%= hoursRemaining %>h</div>
                            <% ElseIf reqStatus = "Collected" Then %>
                                <span class="status-pill status-pill-collected">Collected / Issued</span>
                            <% ElseIf reqStatus = "Expired" Then %>
                                <span class="status-pill status-pill-expired">Expired (48h Passed)</span>
                            <% ElseIf reqStatus = "Rejected" Then %>
                                <span class="status-pill status-pill-rejected">Rejected</span>
                            <% Else %>
                                <span class="status-pill status-pill-pending">Pending Approval</span>
                            <% End If %>
                        </td>
                        <% If IsStaff() Then %>
                            <td class="action-cell">
                                <% If reqStatus = "Pending" Then %>
                                    <div class="action-btn-row">
                                        <form action="requests.asp" method="POST" style="margin:0;">
                                            <input type="hidden" name="action_type" value="approve">
                                            <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                            <button type="submit" class="btn-table-approve">Approve</button>
                                        </form>
                                        <form action="requests.asp" method="POST" style="margin:0;">
                                            <input type="hidden" name="action_type" value="reject">
                                            <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                            <button type="submit" class="btn-table-reject" onclick="return confirm('Reject this request?');">Reject</button>
                                        </form>
                                    </div>
                                <% ElseIf reqStatus = "Approved" Then %>
                                    <form action="requests.asp" method="POST" style="margin:0;">
                                        <input type="hidden" name="action_type" value="collect">
                                        <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                        <button type="submit" class="btn-table-collect">Mark Collected</button>
                                    </form>
                                <% ElseIf reqStatus = "Collected" Then %>
                                    <span class="action-muted-label">Collected</span>
                                <% ElseIf reqStatus = "Rejected" Then %>
                                    <span class="action-muted-label">Rejected</span>
                                <% ElseIf reqStatus = "Expired" Then %>
                                    <span class="action-muted-label">Expired</span>
                                <% Else %>
                                    <span class="action-muted-label"><%= reqStatus %></span>
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
    </div>

    <!-- Table Footer with Pagination matching reference screenshot -->
    <div class="table-footer-bar">
        <div>Showing 1 to <%= rowCount %> of <%= totalReqCount %> requests</div>
        <div class="pagination-controls">
            <button type="button" class="page-btn" title="Previous page">&lt;</button>
            <button type="button" class="page-btn active">1</button>
            <% If totalReqCount > 10 Then %>
                <button type="button" class="page-btn">2</button>
            <% End If %>
            <% If totalReqCount > 20 Then %>
                <button type="button" class="page-btn">3</button>
            <% End If %>
            <button type="button" class="page-btn" title="Next page">&gt;</button>
        </div>
    </div>
</div>

<script>
function filterRequestsTable() {
    var searchVal = document.getElementById("requestsSearchInput").value.toLowerCase();
    var statusVal = document.getElementById("requestsStatusSelect").value.toLowerCase();
    var rows = document.querySelectorAll("#requestsTable tbody tr.req-row");
    var visibleCount = 0;

    rows.forEach(function(row) {
        var rowText = row.innerText.toLowerCase();
        var rowStatus = row.getAttribute("data-status") || "";
        
        var matchesSearch = searchVal === "" || rowText.indexOf(searchVal) > -1;
        var matchesStatus = statusVal === "" || rowStatus.indexOf(statusVal) > -1;

        if (matchesSearch && matchesStatus) {
            row.style.display = "";
            visibleCount++;
        } else {
            row.style.display = "none";
        }
    });

    var noRows = document.getElementById("noRowsFound");
    if (noRows) {
        noRows.style.display = visibleCount === 0 ? "" : "none";
    }
}
</script>

<% RenderFooter %>
