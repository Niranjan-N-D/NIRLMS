<!--#include file="includes/db_config.asp"-->
<%
' Require Authentication
RequireAuth()

Dim conn, rs
Set conn = GetConnection()

' Metrics Counters
Dim totalBooks, activeBorrowings, totalUsers, pendingRequests, totalFeedbacks, kannadaBooks
Dim myActiveBorrowings, myPendingRequests
totalBooks = 0
activeBorrowings = 0
totalUsers = 0
pendingRequests = 0
totalFeedbacks = 0
kannadaBooks = 0
myActiveBorrowings = 0
myPendingRequests = 0

On Error Resume Next

' Total Catalog Books
Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Books")
If Not rs.EOF Then totalBooks = rs("total")
rs.Close

' Total Feedbacks
Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Feedbacks")
If Not rs.EOF Then totalFeedbacks = rs("total")
rs.Close

' Total Kannada Books
Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Books b LEFT JOIN Categories c ON b.category_id = c.id WHERE c.name LIKE '%Kannada%' OR c.name LIKE '%ಕನ್ನಡ%'")
If Not rs.EOF Then kannadaBooks = rs("total")
rs.Close

If IsStaff() Then
    ' Staff (Admin & Librarian) Counters
    Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Borrowings WHERE status IN ('Borrowed', 'Overdue')")
    If Not rs.EOF Then activeBorrowings = rs("total")
    rs.Close

    Set rs = conn.Execute("SELECT COUNT(*) AS total FROM BookRequests WHERE status = 'Pending'")
    If Not rs.EOF Then pendingRequests = rs("total")
    rs.Close

    Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Users")
    If Not rs.EOF Then totalUsers = rs("total")
    rs.Close

    ' Staff Recent Activity Query
    Dim rsRecent
    Set rsRecent = conn.Execute("SELECT b.id, bk.title, u.username, b.borrow_date, b.due_date, b.status FROM Borrowings b LEFT JOIN Books bk ON b.book_id = bk.id LEFT JOIN Users u ON b.user_id = u.id ORDER BY b.borrow_date DESC LIMIT 5")
Else
    ' Staff User Counters (Filtered to Session UserId)
    Dim currentUserId
    currentUserId = SafeInt(Session("UserId"), 0)

    Set rs = conn.Execute("SELECT COUNT(*) AS total FROM Borrowings WHERE user_id = " & currentUserId & " AND status IN ('Borrowed', 'Overdue')")
    If Not rs.EOF Then myActiveBorrowings = rs("total")
    rs.Close

    Set rs = conn.Execute("SELECT COUNT(*) AS total FROM BookRequests WHERE user_id = " & currentUserId & " AND status = 'Pending'")
    If Not rs.EOF Then myPendingRequests = rs("total")
    rs.Close

    ' Staff User Recent Activity Query
    Set rsRecent = conn.Execute("SELECT b.id, bk.title, u.username, b.borrow_date, b.due_date, b.status FROM Borrowings b LEFT JOIN Books bk ON b.book_id = bk.id LEFT JOIN Users u ON b.user_id = u.id WHERE b.user_id = " & currentUserId & " ORDER BY b.borrow_date DESC LIMIT 5")
End If

On Error GoTo 0

RenderHeader "Dashboard"
%>

<% If IsStaff() Then %>
    <h2 style="color:#2c3e50; margin-top:0;">System Overview (<%= Session("Role") %> Dashboard)</h2>

    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px;">
        <div class="card" style="border-left: 5px solid #3498db; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Total Books</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= totalBooks %></p>
        </div>

        <div class="card" style="border-left: 5px solid #f39c12; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Pending Requests</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= pendingRequests %></p>
        </div>

        <div class="card" style="border-left: 5px solid #2ecc71; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Active Borrowings</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= activeBorrowings %></p>
        </div>

        <div class="card" style="border-left: 5px solid #9b59b6; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Total Users</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= totalUsers %></p>
        </div>

        <div class="card" style="border-left: 5px solid #e67e22; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Reviews & Feedback</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= totalFeedbacks %></p>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h3 style="margin:0; color:#2c3e50;">Recent Borrowing Activity</h3>
            <a href="borrowings.asp" class="btn btn-primary">View All</a>
        </div>

        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Book Title</th>
                    <th>Staff Name</th>
                    <th>Borrow Date</th>
                    <th>Due Date (Day 15)</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsRecent Is Nothing Or rsRecent.State = 0 Or rsRecent.EOF Then
                %>
                    <tr>
                        <td colspan="6" style="text-align:center; color:#95a5a6;">No recent borrowings found.</td>
                    </tr>
                <%
                Else
                    Do While Not rsRecent.EOF
                        Dim dStatus, dDueDate, isOver, rowBg
                        dStatus = rsRecent("status")
                        dDueDate = rsRecent("due_date")
                        isOver = (dStatus <> "Returned" And Date > CDate(dDueDate))
                        
                        If isOver Then
                            rowBg = "style='background-color:#fce4e4; color:#721c24; border-left:4px solid #e74c3c; font-weight:bold;'"
                        Else
                            rowBg = ""
                        End If
                %>
                    <tr <%= rowBg %>>
                        <td>#<%= rsRecent("id") %></td>
                        <td><%= CleanText(rsRecent("title") & "") %></td>
                        <td><%= CleanText(rsRecent("username") & "") %></td>
                        <td><%= FormatDateTime(rsRecent("borrow_date"), 2) %></td>
                        <td><%= FormatDateTime(rsRecent("due_date"), 2) %></td>
                        <td>
                            <% If isOver Then %>
                                <span class="badge bg-danger">🚨 OVERDUE (Day 16+)</span>
                            <% ElseIf dStatus = "Returned" Then %>
                                <span class="badge bg-success">Returned</span>
                            <% Else %>
                                <span class="badge bg-warning">Borrowed</span>
                            <% End If %>
                        </td>
                    </tr>
                <%
                        rsRecent.MoveNext
                    Loop
                    rsRecent.Close
                End If
                %>
            </tbody>
        </table>
    </div>

<% Else %>
    <h2 style="color:#2c3e50; margin-top:0;">Welcome, <%= Session("Username") %>! (Staff Dashboard)</h2>

    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px;">
        <div class="card" style="border-left: 5px solid #3498db; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">Total Books Catalog</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= totalBooks %></p>
        </div>

        <div class="card" style="border-left: 5px solid #2ecc71; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">My Borrowed Books</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= myActiveBorrowings %></p>
        </div>

        <div class="card" style="border-left: 5px solid #f39c12; margin-bottom: 0;">
            <h3 style="margin:0; color:#7f8c8d; font-size: 14px; text-transform: uppercase;">My Pending Requests</h3>
            <p style="font-size: 32px; font-weight: bold; color: #2c3e50; margin: 10px 0 0 0;"><%= myPendingRequests %></p>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <h3 style="margin:0; color:#2c3e50;">My Active & Recent Borrowings</h3>
            <a href="request_book.asp" class="btn btn-success">+ Request New Book</a>
        </div>

        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Book Title</th>
                    <th>Borrow Date</th>
                    <th>Due Date (Day 15)</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsRecent Is Nothing Or rsRecent.State = 0 Or rsRecent.EOF Then
                %>
                    <tr>
                        <td colspan="5" style="text-align:center; color:#95a5a6;">No borrowings found in your account.</td>
                    </tr>
                <%
                Else
                    Do While Not rsRecent.EOF
                        Dim mStatus, mDueDate, mIsOver, mRowBg
                        mStatus = rsRecent("status")
                        mDueDate = rsRecent("due_date")
                        mIsOver = (mStatus <> "Returned" And Date > CDate(mDueDate))
                        
                        If mIsOver Then
                            mRowBg = "style='background-color:#fce4e4; color:#721c24; font-weight:bold;'"
                        Else
                            mRowBg = ""
                        End If
                %>
                    <tr <%= mRowBg %>>
                        <td>#<%= rsRecent("id") %></td>
                        <td><strong><%= CleanText(rsRecent("title") & "") %></strong></td>
                        <td><%= FormatDateTime(rsRecent("borrow_date"), 2) %></td>
                        <td><%= FormatDateTime(rsRecent("due_date"), 2) %></td>
                        <td>
                            <% If mIsOver Then %>
                                <span class="badge bg-danger">🚨 OVERDUE</span>
                            <% ElseIf mStatus = "Returned" Then %>
                                <span class="badge bg-success">Returned</span>
                            <% Else %>
                                <span class="badge bg-warning">Borrowed</span>
                            <% End If %>
                        </td>
                    </tr>
                <%
                        rsRecent.MoveNext
                    Loop
                    rsRecent.Close
                End If
                %>
            </tbody>
        </table>
    </div>
<% End If %>

<%
conn.Close
Set conn = Nothing
RenderFooter
%>