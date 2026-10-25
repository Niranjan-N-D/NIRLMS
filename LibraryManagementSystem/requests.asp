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

' Fetch Requests List with all 7 Metadata Fields
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

RenderHeader IIf(IsStaff(), "Book Requests Approval", "My Book Requests")
%>

<div class="card">
    <div class="card-header">
        <div>
            <h2 style="margin:0; color:#2c3e50;"><%= IIf(IsStaff(), "Staff Book Requests & Collection Approval", "My Book Requests") %></h2>
            <p style="margin:4px 0 0 0; color:#7f8c8d; font-size:13px;">Approved requests must be collected within <strong>48 hours</strong>; otherwise they auto-expire and return to inventory.</p>
        </div>
        <a href="request_book.asp" class="btn btn-success">+ Request a Book</a>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
    <% End If %>

    <% If successMessage <> "" Then %>
        <div class="alert alert-info"><%= CleanText(successMessage) %></div>
    <% End If %>

    <table style="font-size:13px;">
        <thead>
            <tr>
                <th>Req / Book ID</th>
                <th>Staff Name</th>
                <th>Staff Phone</th>
                <th>Internal Ext No.</th>
                <th>Book Title</th>
                <th>Author</th>
                <th>Price (&#8377;)</th>
                <th>Borrow Dates</th>
                <th>Status &amp; 48h Window</th>
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
                    <td colspan="<%= IIf(IsStaff(), 10, 9) %>" style="text-align:center; color:#95a5a6; padding:20px;">No book requests found.</td>
                </tr>
            <%
            Else
                Do While Not rsList.EOF
                    reqStatus = rsList("status")
                    approvedAt = rsList("approved_at")
                    hoursRemaining = 48
                    
                    If reqStatus = "Approved" And Not IsNull(approvedAt) Then
                        hoursPassed = DateDiff("h", approvedAt, Now())
                        hoursRemaining = 48 - hoursPassed
                        If hoursRemaining < 0 Then hoursRemaining = 0
                    End If
            %>
                <tr>
                    <td>
                        <strong style="color:#2c3e50;">Req #<%= rsList("req_id") %></strong><br>
                        <small style="color:#7f8c8d;">Book ID: #<%= rsList("book_id") %></small>
                    </td>
                    <td><strong><%= CleanText(rsList("staff_name") & "") %></strong></td>
                    <td><code><%= CleanText(rsList("staff_number") & "") %></code></td>
                    <td><span class="badge bg-secondary" style="background:#6c757d;"><%= CleanText(rsList("staff_internal_number") & "") %></span></td>
                    <td><strong style="color:#16a085; font-size:14px;"><%= CleanText(rsList("book_name") & "") %></strong></td>
                    <td><%= CleanText(rsList("author_name") & "") %></td>
                    <td><strong style="color:#27ae60;">&#8377;<%= FormatNumber(SafeFloat(rsList("price"), 0), 2) %></strong></td>
                    <td>
                        <small>From: <%= FormatDateTime(rsList("start_date"), 2) %></small><br>
                        <small>To: <%= FormatDateTime(rsList("end_date"), 2) %></small>
                    </td>
                    <td>
                        <% If reqStatus = "Approved" Then %>
                            <span class="badge bg-success">Approved</span><br>
                            <small style="color:#27ae60; font-weight:bold;">Collect in <%= hoursRemaining %>h</small>
                        <% ElseIf reqStatus = "Collected" Then %>
                            <span class="badge bg-primary" style="background:#3498db;">Collected / Issued</span>
                        <% ElseIf reqStatus = "Expired" Then %>
                            <span class="badge bg-danger">Expired (48h Passed)</span>
                        <% ElseIf reqStatus = "Rejected" Then %>
                            <span class="badge bg-danger">Rejected</span>
                        <% Else %>
                            <span class="badge bg-warning">Pending Approval</span>
                        <% End If %>
                    </td>
                    <% If IsStaff() Then %>
                        <td>
                            <% If reqStatus = "Pending" Then %>
                                <form action="requests.asp" method="POST" style="display:inline-block; margin-right:4px;">
                                    <input type="hidden" name="action_type" value="approve">
                                    <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                    <button type="submit" class="btn btn-success" style="padding:4px 8px; font-size:11px;">Approve</button>
                                </form>
                                <form action="requests.asp" method="POST" style="display:inline-block;">
                                    <input type="hidden" name="action_type" value="reject">
                                    <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                    <button type="submit" class="btn btn-danger" style="padding:4px 8px; font-size:11px;" onclick="return confirm('Reject this request?');">Reject</button>
                                </form>
                            <% ElseIf reqStatus = "Approved" Then %>
                                <form action="requests.asp" method="POST" style="display:inline-block;">
                                    <input type="hidden" name="action_type" value="collect">
                                    <input type="hidden" name="request_id" value="<%= rsList("req_id") %>">
                                    <button type="submit" class="btn btn-primary" style="padding:4px 8px; font-size:11px; background:#27ae60;">Mark Collected</button>
                                </form>
                            <% Else %>
                                <span style="color:#95a5a6; font-size:11px;"><%= reqStatus %></span>
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

<% RenderFooter %>
