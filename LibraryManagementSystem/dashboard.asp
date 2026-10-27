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

<style>
.dash-metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
    gap: 18px;
    margin-bottom: 26px;
}
.dash-metric-card {
    display: flex;
    align-items: center;
    gap: 16px;
    padding: 18px 20px;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.dash-metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 14px rgba(0,0,0,0.05);
}
.dm-blue { background-color: #f0f7ff; border-color: #d0e7ff; }
.dm-amber { background-color: #fffbeb; border-color: #fef3c7; }
.dm-green { background-color: #f0fdf4; border-color: #dcfce7; }
.dm-purple { background-color: #f3e8ff; border-color: #f3d9fa; }
.dm-teal { background-color: #ecfdf5; border-color: #a7f3d0; }

.dm-icon-box {
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 22px;
    flex-shrink: 0;
}
.dm-icon-blue { background-color: #dbeafe; color: #0284c7; }
.dm-icon-amber { background-color: #fef3c7; color: #d97706; }
.dm-icon-green { background-color: #dcfce7; color: #16a34a; }
.dm-icon-purple { background-color: #f3d9fa; color: #7e22ce; }
.dm-icon-teal { background-color: #ccfbf1; color: #0f766e; }

.dm-val {
    font-size: 28px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1;
    margin-bottom: 4px;
}
.dm-title {
    font-size: 13.5px;
    font-weight: 700;
    color: #334155;
}
.dm-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 2px;
}
</style>

<% If IsStaff() Then %>
    <div style="margin-bottom:20px;">
        <h1 style="color:#0f172a; font-size:24px; font-weight:800; margin:0;">System Overview (<%= Session("Role") %> Dashboard)</h1>
        <p style="color:#64748b; font-size:13.5px; margin:4px 0 0 0;">Real-time metrics, active loans, and borrowing activities at a glance.</p>
    </div>

    <div class="dash-metric-grid">
        <div class="dash-metric-card dm-blue">
            <div class="dm-icon-box dm-icon-blue">📚</div>
            <div>
                <div class="dm-val"><%= totalBooks %></div>
                <div class="dm-title">Total Books</div>
                <div class="dm-sub">In Catalog</div>
            </div>
        </div>

        <div class="dash-metric-card dm-amber">
            <div class="dm-icon-box dm-icon-amber">📋</div>
            <div>
                <div class="dm-val" style="color:#d97706;"><%= pendingRequests %></div>
                <div class="dm-title" style="color:#b45309;">Pending Requests</div>
                <div class="dm-sub">Awaiting action</div>
            </div>
        </div>

        <div class="dash-metric-card dm-green">
            <div class="dm-icon-box dm-icon-green">⏱️</div>
            <div>
                <div class="dm-val" style="color:#16a34a;"><%= activeBorrowings %></div>
                <div class="dm-title" style="color:#15803d;">Active Borrowings</div>
                <div class="dm-sub">On loan</div>
            </div>
        </div>

        <div class="dash-metric-card dm-purple">
            <div class="dm-icon-box dm-icon-purple">👥</div>
            <div>
                <div class="dm-val" style="color:#7e22ce;"><%= totalUsers %></div>
                <div class="dm-title" style="color:#6b21a8;">Total Users</div>
                <div class="dm-sub">Registered staff</div>
            </div>
        </div>

        <div class="dash-metric-card dm-teal">
            <div class="dm-icon-box dm-icon-teal">💬</div>
            <div>
                <div class="dm-val" style="color:#0f766e;"><%= totalFeedbacks %></div>
                <div class="dm-title" style="color:#115e59;">Reviews &amp; Feedback</div>
                <div class="dm-sub">Reader ratings</div>
            </div>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <div>
                <h2 style="margin:0; font-size:18px; font-weight:700; color:#0f172a;">Recent Borrowing Activity</h2>
                <p style="margin:4px 0 0 0; color:#64748b; font-size:13px;">Latest book checkouts and return tracking.</p>
            </div>
            <a href="borrowings.asp" class="btn btn-primary">View All Borrowings ➔</a>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Loan ID</th>
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
                        <td colspan="6" style="text-align:center; color:#94a3b8; padding:24px;">No recent borrowings found.</td>
                    </tr>
                <%
                Else
                    Do While Not rsRecent.EOF
                        Dim dStatus, dDueDate, isOver, rowBg
                        dStatus = rsRecent("status")
                        dDueDate = rsRecent("due_date")
                        isOver = (dStatus <> "Returned" And Date > CDate(dDueDate))
                        
                        If isOver Then
                            rowBg = "style='background-color:#fff5f5; color:#991b1b; font-weight:600;'"
                        Else
                            rowBg = ""
                        End If
                %>
                    <tr <%= rowBg %>>
                        <td><strong style="color:#0f172a;">#<%= rsRecent("id") %></strong></td>
                        <td><strong style="color:#0d9488;"><%= CleanText(rsRecent("title") & "") %></strong></td>
                        <td><%= CleanText(rsRecent("username") & "") %></td>
                        <td style="color:#64748b;"><%= FormatDateTime(rsRecent("borrow_date"), 2) %></td>
                        <td><strong style="color:<%= IIf(isOver, "#dc2626", "#0f172a") %>;"><%= FormatDateTime(rsRecent("due_date"), 2) %></strong></td>
                        <td>
                            <% If isOver Then %>
                                <span class="status-pill status-pill-red">🚨 Overdue</span>
                            <% ElseIf dStatus = "Returned" Then %>
                                <span class="status-pill status-pill-green">Returned</span>
                            <% Else %>
                                <span class="status-pill status-pill-amber">Borrowed</span>
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
    <div style="margin-bottom:20px;">
        <h1 style="color:#0f172a; font-size:24px; font-weight:800; margin:0;">Welcome, <%= Session("Username") %>!</h1>
        <p style="color:#64748b; font-size:13.5px; margin:4px 0 0 0;">Member Portal &bull; Track your loans and request new books.</p>
    </div>

    <div class="dash-metric-grid">
        <div class="dash-metric-card dm-blue">
            <div class="dm-icon-box dm-icon-blue">📚</div>
            <div>
                <div class="dm-val"><%= totalBooks %></div>
                <div class="dm-title">Total Catalog Books</div>
                <div class="dm-sub">Available titles</div>
            </div>
        </div>

        <div class="dash-metric-card dm-green">
            <div class="dm-icon-box dm-icon-green">📖</div>
            <div>
                <div class="dm-val" style="color:#16a34a;"><%= myActiveBorrowings %></div>
                <div class="dm-title" style="color:#15803d;">My Active Loans</div>
                <div class="dm-sub">Currently reading</div>
            </div>
        </div>

        <div class="dash-metric-card dm-amber">
            <div class="dm-icon-box dm-icon-amber">📋</div>
            <div>
                <div class="dm-val" style="color:#d97706;"><%= myPendingRequests %></div>
                <div class="dm-title" style="color:#b45309;">My Pending Requests</div>
                <div class="dm-sub">Awaiting librarian</div>
            </div>
        </div>
    </div>

    <div class="card">
        <div class="card-header">
            <div>
                <h2 style="margin:0; font-size:18px; font-weight:700; color:#0f172a;">My Active &amp; Recent Borrowings</h2>
                <p style="margin:4px 0 0 0; color:#64748b; font-size:13px;">Manage your reading schedule and due dates.</p>
            </div>
            <a href="request_book.asp" class="btn btn-success">+ Request New Book</a>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Loan ID</th>
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
                        <td colspan="5" style="text-align:center; color:#94a3b8; padding:24px;">No borrowings found in your account.</td>
                    </tr>
                <%
                Else
                    Do While Not rsRecent.EOF
                        Dim mStatus, mDueDate, mIsOver, mRowBg
                        mStatus = rsRecent("status")
                        mDueDate = rsRecent("due_date")
                        mIsOver = (mStatus <> "Returned" And Date > CDate(mDueDate))
                        
                        If mIsOver Then
                            mRowBg = "style='background-color:#fff5f5; color:#991b1b; font-weight:600;'"
                        Else
                            mRowBg = ""
                        End If
                %>
                    <tr <%= mRowBg %>>
                        <td><strong style="color:#0f172a;">#<%= rsRecent("id") %></strong></td>
                        <td><strong style="color:#0d9488;"><%= CleanText(rsRecent("title") & "") %></strong></td>
                        <td style="color:#64748b;"><%= FormatDateTime(rsRecent("borrow_date"), 2) %></td>
                        <td><strong style="color:<%= IIf(mIsOver, "#dc2626", "#0f172a") %>;"><%= FormatDateTime(rsRecent("due_date"), 2) %></strong></td>
                        <td>
                            <% If mIsOver Then %>
                                <span class="status-pill status-pill-red">🚨 Overdue</span>
                            <% ElseIf mStatus = "Returned" Then %>
                                <span class="status-pill status-pill-green">Returned</span>
                            <% Else %>
                                <span class="status-pill status-pill-amber">Borrowed</span>
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