<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim conn, errorMessage, bookId, rsBook, defaultStart, defaultEnd
errorMessage = ""
bookId = SafeInt(Request("id"), 0)

Set conn = GetConnection()

' Fetch Current User / Staff Info
Dim currentUserId, currentStaffName, currentStaffNumber, currentStaffInternalNum
currentUserId = SafeInt(Session("UserId"), 0)

Dim rsUser
Set rsUser = conn.Execute("SELECT username, staff_number, staff_internal_number FROM Users WHERE id = " & currentUserId)
If Not (rsUser Is Nothing Or rsUser.EOF) Then
    currentStaffName = rsUser("username") & ""
    currentStaffNumber = rsUser("staff_number") & ""
    currentStaffInternalNum = rsUser("staff_internal_number") & ""
Else
    currentStaffName = Session("Username") & ""
    currentStaffNumber = "9876543210"
    currentStaffInternalNum = "BEL-EXT-101"
End If

' Form Submission (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    bookId = SafeInt(Request.Form("book_id"), 0)
    Dim staffName, staffNumber, staffInternalNum, bookName, authorName, bookPrice
    Dim startDateStr, endDateStr

    staffName = Trim(Request.Form("staff_name"))
    staffNumber = Trim(Request.Form("staff_number"))
    staffInternalNum = Trim(Request.Form("staff_internal_number"))
    startDateStr = Trim(Request.Form("start_date"))
    endDateStr = Trim(Request.Form("end_date"))

    If bookId = 0 Then
        errorMessage = "Please select a valid book to request."
    ElseIf startDateStr = "" Or endDateStr = "" Then
        errorMessage = "Both Start Date and End Date are required."
    ElseIf CDate(endDateStr) < CDate(startDateStr) Then
        errorMessage = "End Date cannot be earlier than Start Date."
    Else
        ' Fetch Book, Author, and Price details from database for safety
        Dim rsSelected
        Set rsSelected = conn.Execute("SELECT b.id, b.title, COALESCE(b.price, 0.00) AS price, COALESCE(a.name, 'Unknown Author') AS author_name " & _
                                      "FROM Books b LEFT JOIN Authors a ON b.author_id = a.id WHERE b.id = " & bookId)
        
        If Not (rsSelected Is Nothing Or rsSelected.EOF) Then
            bookName = rsSelected("title") & ""
            authorName = rsSelected("author_name") & ""
            bookPrice = SafeFloat(rsSelected("price"), 0.00)
        Else
            bookName = "Unknown Book"
            authorName = "Unknown Author"
            bookPrice = 0.00
        End If

        Dim insertSql
        insertSql = "INSERT INTO BookRequests (book_id, user_id, staff_name, staff_number, staff_internal_number, book_name, author_name, price, start_date, end_date, status) VALUES (" & _
                    bookId & ", " & _
                    currentUserId & ", " & _
                    SQLQuote(staffName) & ", " & _
                    SQLQuote(staffNumber) & ", " & _
                    SQLQuote(staffInternalNum) & ", " & _
                    SQLQuote(bookName) & ", " & _
                    SQLQuote(authorName) & ", " & _
                    bookPrice & ", " & _
                    FormatSQLDate(startDateStr) & ", " & _
                    FormatSQLDate(endDateStr) & ", 'Pending')"

        On Error Resume Next
        conn.Execute(insertSql)
        If Err.Number <> 0 Then
            errorMessage = "Failed to submit request: " & Err.Description
        Else
            conn.Close
            Set conn = Nothing
            Response.Redirect("requests.asp?msg=Book+request+submitted+successfully!+Awaiting+Librarian/Admin+approval.")
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' Fetch All Available Books list with Author and Price for dropdown & pre-selection
Dim rsAllBooks
Set rsAllBooks = conn.Execute("SELECT b.id, b.title, b.price, a.name AS author_name FROM Books b LEFT JOIN Authors a ON b.author_id = a.id ORDER BY b.title ASC")

' Default date strings (YYYY-MM-DD) -> Start Today, End Date = 15 Days Later
defaultStart = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
defaultEnd = Year(DateAdd("d", 15, Now)) & "-" & Right("0" & Month(DateAdd("d", 15, Now)), 2) & "-" & Right("0" & Day(DateAdd("d", 15, Now)), 2)

RenderHeader "Request Book"
%>

<div class="card" style="max-width: 650px; margin: 20px auto;">
    <div class="card-header">
        <h2 style="margin:0; color:#2c3e50;">Request a Book (Staff Portal)</h2>
        <a href="books.asp" class="btn btn-secondary">Back to Catalog</a>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
    <% End If %>

    <form action="request_book.asp" method="POST">
        <!-- 1. Staff Metadata Fields -->
        <div style="background:#eef2f7; padding:15px; border-radius:6px; margin-bottom:15px;">
            <h4 style="margin-top:0; margin-bottom:10px; color:#34495e;">Staff Details</h4>
            <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:10px;">
                <div class="form-group" style="margin-bottom:0;">
                    <label for="staff_name" style="font-size:12px;">Staff Name *</label>
                    <input type="text" id="staff_name" name="staff_name" value="<%= CleanText(currentStaffName) %>" required style="padding:6px 10px;">
                </div>
                <div class="form-group" style="margin-bottom:0;">
                    <label for="staff_number" style="font-size:12px;">Staff Number (Phone) *</label>
                    <input type="text" id="staff_number" name="staff_number" value="<%= CleanText(currentStaffNumber) %>" required style="padding:6px 10px;">
                </div>
                <div class="form-group" style="margin-bottom:0;">
                    <label for="staff_internal_number" style="font-size:12px;">Staff Internal Ext No. *</label>
                    <input type="text" id="staff_internal_number" name="staff_internal_number" value="<%= CleanText(currentStaffInternalNum) %>" required style="padding:6px 10px;">
                </div>
            </div>
        </div>

        <!-- 2. Book Selection -->
        <div class="form-group">
            <label for="book_id">Select Book *</label>
            <select id="book_id" name="book_id" required style="font-size:14px; padding:8px 12px;">
                <option value="">-- Choose Book --</option>
                <%
                If Not rsAllBooks Is Nothing And rsAllBooks.State = 1 Then
                    Do While Not rsAllBooks.EOF
                        Dim selectedAttr
                        If CStr(rsAllBooks("id")) = CStr(bookId) Then selectedAttr = "selected" Else selectedAttr = ""
                %>
                    <option value="<%= rsAllBooks("id") %>" <%= selectedAttr %>>
                        [ID: <%= rsAllBooks("id") %>] <%= CleanText(rsAllBooks("title") & "") %> — by <%= CleanText(rsAllBooks("author_name") & "") %> (&#8377;<%= FormatNumber(SafeFloat(rsAllBooks("price"), 0), 2) %>)
                    </option>
                <%
                        rsAllBooks.MoveNext
                    Loop
                End If
                %>
            </select>
        </div>

        <!-- 3. Borrowing Dates -->
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:15px;">
            <div class="form-group">
                <label for="start_date">Start Date (Borrow Date) *</label>
                <input type="date" id="start_date" name="start_date" value="<%= defaultStart %>" required style="padding:8px 12px;">
            </div>

            <div class="form-group">
                <label for="end_date">Return Due Date (15 Days) *</label>
                <input type="date" id="end_date" name="end_date" value="<%= defaultEnd %>" required style="padding:8px 12px;">
            </div>
        </div>

        <div style="background:#fff3cd; color:#856404; padding:12px; border-radius:6px; margin-bottom:20px; font-size:13px; border:1px solid #ffeeba;">
            💡 <strong>Borrowing Rules:</strong><br>
            • Standard duration is <strong>15 days</strong>.<br>
            • On day 16, borrowings automatically trigger <strong><span style="color:#d9534f; font-weight:bold;">OVERDUE</span></strong> status.<br>
            • Approved requests expire after <strong>48 hours</strong> if uncollected.
        </div>

        <button type="submit" class="btn btn-primary" style="width:100%; padding:12px; font-size:16px;">Submit Book Request</button>
    </form>
</div>

<%
conn.Close
Set conn = Nothing
RenderFooter
%>
