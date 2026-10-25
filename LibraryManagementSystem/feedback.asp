<!--#include file="includes/db_config.asp"-->
<%
' Require Authentication & Check Permission
RequireAuth()
If Not HasPermission("feedback") Then
    Response.Redirect("dashboard.asp?msg=Access+Denied:+Feedback+feature+disabled+for+your+account")
    Response.End
End If

Dim conn, errorMessage, successMessage
errorMessage = ""
successMessage = ""
Set conn = GetConnection()

Dim selectedBookId
selectedBookId = SafeInt(Request.QueryString("book_id"), 0)

' Handle Form Submission (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim feedbackType, bookIdVal, rating, comment, currentUserId
    feedbackType = Trim(Request.Form("feedback_type"))
    bookIdVal = SafeInt(Request.Form("book_id"), 0)
    rating = SafeInt(Request.Form("rating"), 5)
    comment = Trim(Request.Form("comment"))
    currentUserId = SafeInt(Session("UserId"), 0)

    If feedbackType <> "book" Then
        feedbackType = "website"
        bookIdVal = 0
    End If

    If rating < 1 Or rating > 5 Then rating = 5

    ' Server-Side 150 Words Limit Validation
    Dim wordList, wordCount
    wordCount = 0
    If comment <> "" Then
        ' Normalize spaces and split
        Dim tempComment
        tempComment = Replace(comment, vbCrLf, " ")
        tempComment = Replace(tempComment, vbTab, " ")
        wordList = Split(tempComment, " ")
        
        Dim i, w
        For i = 0 To UBound(wordList)
            w = Trim(wordList(i))
            If w <> "" Then wordCount = wordCount + 1
        Next
    End If

    If comment = "" Then
        errorMessage = "Feedback comment cannot be empty."
    ElseIf wordCount > 150 Then
        errorMessage = "Validation Error: Your review exceeds the 150-word limit (" & wordCount & " words submitted). Please shorten your feedback."
    Else
        On Error Resume Next
        Dim sqlInsert
        If bookIdVal > 0 And feedbackType = "book" Then
            sqlInsert = "INSERT INTO Feedbacks (user_id, book_id, feedback_type, rating, comment) VALUES (" & _
                        currentUserId & ", " & bookIdVal & ", 'book', " & rating & ", " & SQLQuote(comment) & ")"
        Else
            sqlInsert = "INSERT INTO Feedbacks (user_id, book_id, feedback_type, rating, comment) VALUES (" & _
                        currentUserId & ", NULL, 'website', " & rating & ", " & SQLQuote(comment) & ")"
        End If
        
        conn.Execute(sqlInsert)
        
        If Err.Number <> 0 Then
            errorMessage = "Database error saving feedback: " & Err.Description
        Else
            successMessage = "Thank you! Your feedback/review has been submitted successfully."
            comment = ""
        End If
        On Error GoTo 0
    End If
End If

' Fetch Books for Selection Dropdown
Dim rsBooks
Set rsBooks = conn.Execute("SELECT id, title FROM Books ORDER BY title ASC")

' Fetch Recent Feedbacks & Reviews
Dim rsFeedbacks
Set rsFeedbacks = conn.Execute("SELECT f.id, f.feedback_type, f.rating, f.comment, f.created_at, u.username, b.title AS book_title " & _
                               "FROM Feedbacks f " & _
                               "LEFT JOIN Users u ON f.user_id = u.id " & _
                               "LEFT JOIN Books b ON f.book_id = b.id " & _
                               "ORDER BY f.id DESC LIMIT 20")

RenderHeader "Feedback & Reviews"
%>

<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px;">
    <!-- Left Column: Submit Review / Feedback Form -->
    <div class="card" style="margin-bottom:0;">
        <h2 style="margin-top:0; color:#2c3e50;">✍️ Submit Feedback or Book Review</h2>
        <p style="color:#7f8c8d; font-size:14px;">Share your thoughts about our website or review a book (Strict Limit: <strong>150 words maximum</strong>).</p>

        <% If errorMessage <> "" Then %>
            <div class="alert alert-danger"><%= CleanText(errorMessage) %></div>
        <% End If %>

        <% If successMessage <> "" Then %>
            <div class="alert alert-info"><%= CleanText(successMessage) %></div>
        <% End If %>

        <form action="feedback.asp" method="POST" id="feedbackForm" onsubmit="return validateWordCount();">
            <div class="form-group">
                <label for="feedback_type">Feedback Category</label>
                <select id="feedback_type" name="feedback_type" onchange="toggleBookDropdown();">
                    <option value="website" <%= IIf(selectedBookId = 0, "selected", "") %>>🌐 General Website Feedback</option>
                    <option value="book" <%= IIf(selectedBookId > 0, "selected", "") %>>📚 Book Review</option>
                </select>
            </div>

            <div class="form-group" id="bookSelectGroup" style="<%= IIf(selectedBookId > 0, "", "display:none;") %>">
                <label for="book_id">Select Book</label>
                <select id="book_id" name="book_id">
                    <option value="0">-- Select a Book --</option>
                    <%
                    If Not (rsBooks Is Nothing Or rsBooks.State = 0) Then
                        Do While Not rsBooks.EOF
                    %>
                        <option value="<%= rsBooks("id") %>" <%= IIf(rsBooks("id") = selectedBookId, "selected", "") %>>
                            <%= CleanText(rsBooks("title") & "") %>
                        </option>
                    <%
                            rsBooks.MoveNext
                        Loop
                        rsBooks.MoveFirst
                    End If
                    %>
                </select>
            </div>

            <div class="form-group">
                <label for="rating">Rating (1 to 5 Stars)</label>
                <select id="rating" name="rating">
                    <option value="5" selected>⭐⭐⭐⭐⭐ (5 - Excellent)</option>
                    <option value="4">⭐⭐⭐⭐ (4 - Very Good)</option>
                    <option value="3">⭐⭐⭐ (3 - Good)</option>
                    <option value="2">⭐⭐ (2 - Fair)</option>
                    <option value="1">⭐ (1 - Poor)</option>
                </select>
            </div>

            <div class="form-group">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
                    <label for="comment" style="margin:0;">Comments & Review</label>
                    <span id="wordCounter" style="font-size:12px; font-weight:bold; color:#7f8c8d;">Words: 0 / 150</span>
                </div>
                <textarea id="comment" name="comment" rows="6" placeholder="Write your feedback or book review here (maximum 150 words)..." oninput="updateWordCount();" required></textarea>
                <div id="wordWarning" style="color:#e74c3c; font-size:12px; margin-top:4px; display:none;">⚠️ Warning: You have exceeded the 150-word limit! Please reduce text.</div>
            </div>

            <button type="submit" id="submitBtn" class="btn btn-primary" style="width:100%; padding:10px;">Submit Feedback</button>
        </form>
    </div>

    <!-- Right Column: Recent Feedbacks & Book Reviews Stream -->
    <div class="card" style="margin-bottom:0;">
        <h2 style="margin-top:0; color:#2c3e50;">💬 Community Reviews & Feedback</h2>
        
        <div style="max-height: 520px; overflow-y: auto; padding-right: 5px;">
            <%
            If rsFeedbacks Is Nothing Or rsFeedbacks.State = 0 Or rsFeedbacks.EOF Then
            %>
                <p style="text-align:center; color:#95a5a6; margin-top:40px;">No feedback or reviews submitted yet.</p>
            <%
            Else
                Do While Not rsFeedbacks.EOF
                    Dim stars, rVal
                    rVal = SafeInt(rsFeedbacks("rating"), 5)
                    stars = String(rVal, "⭐")
            %>
                <div style="border-bottom:1px solid #eee; padding-bottom:15px; margin-bottom:15px;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <strong><%= CleanText(rsFeedbacks("username") & "") %></strong>
                        <span style="font-size:12px; color:#f39c12;"><%= stars %></span>
                    </div>
                    <div style="margin:4px 0;">
                        <% If rsFeedbacks("feedback_type") = "book" And Not IsNull(rsFeedbacks("book_title")) Then %>
                            <span class="badge bg-success" style="font-size:11px;">Book Review: <%= CleanText(rsFeedbacks("book_title") & "") %></span>
                        <% Else %>
                            <span class="badge bg-warning" style="font-size:11px;">Website Feedback</span>
                        <% End If %>
                        <span style="font-size:11px; color:#95a5a6; margin-left:8px;"><%= FormatDateTime(rsFeedbacks("created_at"), 2) %></span>
                    </div>
                    <p style="margin:8px 0 0 0; color:#444; font-size:14px; line-height:1.4;">
                        <%= CleanText(rsFeedbacks("comment") & "") %>
                    </p>
                </div>
            <%
                    rsFeedbacks.MoveNext
                Loop
                rsFeedbacks.Close
            End If
            
            conn.Close
            Set conn = Nothing
            %>
        </div>
    </div>
</div>

<script>
function toggleBookDropdown() {
    var typeSelect = document.getElementById('feedback_type');
    var bookGroup = document.getElementById('bookSelectGroup');
    if (typeSelect.value === 'book') {
        bookGroup.style.display = 'block';
    } else {
        bookGroup.style.display = 'none';
    }
}

function getWordCount(text) {
    var trimmed = text.trim();
    if (trimmed === '') return 0;
    return trimmed.split(/\s+/).length;
}

function updateWordCount() {
    var textarea = document.getElementById('comment');
    var counter = document.getElementById('wordCounter');
    var warning = document.getElementById('wordWarning');
    var submitBtn = document.getElementById('submitBtn');
    
    var count = getWordCount(textarea.value);
    counter.innerText = 'Words: ' + count + ' / 150';
    
    if (count > 150) {
        counter.style.color = '#e74c3c';
        warning.style.display = 'block';
        submitBtn.disabled = true;
        submitBtn.style.opacity = '0.6';
    } else {
        counter.style.color = '#7f8c8d';
        warning.style.display = 'none';
        submitBtn.disabled = false;
        submitBtn.style.opacity = '1';
    }
}

function validateWordCount() {
    var textarea = document.getElementById('comment');
    var count = getWordCount(textarea.value);
    if (count > 150) {
        alert('Your review exceeds the maximum limit of 150 words. Please shorten your text before submitting.');
        return false;
    }
    return true;
}
</script>

<% RenderFooter %>
