<!--#include file="includes/db_config.asp"-->
<%
' ====================================================================
' User & Staff Management Module (.asp VBScript Backend)
' Accessible by Admin and Librarian
' ====================================================================
RequireAuth()

If Not IsStaff() Then
    Response.Redirect("dashboard.asp?msg=Access+Denied:+Admin+and+Librarian+only")
    Response.End
End If

Dim conn, errorMessage, successMessage
errorMessage = ""
successMessage = ""
Set conn = GetConnection()

' Handle Form Submissions (Add / Delete User)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim actionType
    actionType = Request.Form("action_type")

    If actionType = "add" Then
        Dim username, staffNumber, internalNumber, sbuVal, dob, email, role
        username = Trim(Request.Form("username"))
        staffNumber = Trim(Request.Form("staff_number"))
        internalNumber = Trim(Request.Form("internal_no"))
        If internalNumber = "" Then internalNumber = Trim(Request.Form("staff_internal_number"))
        If internalNumber = "" Then internalNumber = "BEL-EXT-" & Right(staffNumber, 3)
        sbuVal = Trim(Request.Form("sbu"))
        If sbuVal = "" Then sbuVal = "General"
        dob = Trim(Request.Form("dob"))
        email = Trim(Request.Form("email"))
        role = Trim(Request.Form("role"))

        ' Security: Librarians cannot create Admin accounts
        If Not IsAdmin() And LCase(role) = "admin" Then
            role = "Staff"
        End If
        If role = "" Then role = "Staff"

        ' Validation
        If username = "" Or staffNumber = "" Or dob = "" Or email = "" Then
            errorMessage = "Full Name, Staff Number, Date of Birth (Password), and Email are required."
        Else
            ' Check for existing staff_number or email among active accounts
            Dim rsCheck, checkSql
            checkSql = "SELECT id, staff_number, email FROM Users WHERE (staff_number = " & SQLQuote(staffNumber) & " OR email = " & SQLQuote(email) & ") AND COALESCE(is_deleted, 0) = 0"
            Set rsCheck = Server.CreateObject("ADODB.Recordset")
            rsCheck.Open checkSql, conn, 1, 1

            If Not rsCheck.EOF Then
                If LCase(Trim(rsCheck("staff_number") & "")) = LCase(staffNumber) Then
                    errorMessage = "A user with Staff Number '" & CleanText(staffNumber) & "' already exists."
                Else
                    errorMessage = "A user with Email '" & CleanText(email) & "' already exists."
                End If
                rsCheck.Close
                Set rsCheck = Nothing
            Else
                rsCheck.Close
                Set rsCheck = Nothing

                ' Default Feature Permissions by Role
                Dim pBooks, pRequests, pBorrowings, pAuthors, pCategories, pFeedback
                If LCase(role) = "admin" Or LCase(role) = "librarian" Then
                    pBooks = 1
                    pRequests = 1
                    pBorrowings = 1
                    pAuthors = 1
                    pCategories = 1
                    pFeedback = 1
                Else
                    pBooks = 1
                    pRequests = 1
                    pBorrowings = 1
                    pAuthors = 0
                    pCategories = 0
                    pFeedback = 1
                End If

                On Error Resume Next
                Dim insertSql
                insertSql = "INSERT INTO Users (username, staff_number, staff_internal_number, sbu, dob, email, password_hash, role, perm_books, perm_requests, perm_borrowings, perm_authors, perm_categories, perm_feedback) VALUES (" & _
                    SQLQuote(username) & ", " & _
                    SQLQuote(staffNumber) & ", " & _
                    SQLQuote(internalNumber) & ", " & _
                    SQLQuote(sbuVal) & ", " & _
                    SQLQuote(dob) & ", " & _
                    SQLQuote(email) & ", " & _
                    SQLQuote(dob) & ", " & _
                    SQLQuote(role) & ", " & _
                    pBooks & ", " & _
                    pRequests & ", " & _
                    pBorrowings & ", " & _
                    pAuthors & ", " & _
                    pCategories & ", " & _
                    pFeedback & ")"
                
                conn.Execute(insertSql)

                If Err.Number <> 0 Then
                    errorMessage = "Failed to add user: " & Err.Description
                Else
                    successMessage = "User '" & CleanText(username) & "' (" & CleanText(role) & ") created successfully! Staff No: " & CleanText(staffNumber) & " | Internal No: " & CleanText(internalNumber) & " | SBU: " & CleanText(sbuVal) & " | Login DOB: " & CleanText(dob)
                End If
                On Error GoTo 0
            End If
        End If

    ElseIf actionType = "delete" Then
        Dim deleteId
        deleteId = SafeInt(Request.Form("user_id"), 0)

        If deleteId = 0 Then
            errorMessage = "Invalid user selection."
        ElseIf deleteId = SafeInt(Session("UserId"), 0) Then
            errorMessage = "You cannot delete your own logged-in account!"
        Else
            ' Check target role
            Dim rsTarget
            Set rsTarget = conn.Execute("SELECT role, username FROM Users WHERE id = " & deleteId)
            If rsTarget Is Nothing Or rsTarget.State = 0 Or rsTarget.EOF Then
                errorMessage = "User not found."
            Else
                Dim targetRole, targetName
                targetRole = Trim(rsTarget("role") & "")
                targetName = Trim(rsTarget("username") & "")
                rsTarget.Close

                If Not IsAdmin() And LCase(targetRole) = "admin" Then
                    errorMessage = "Librarians cannot delete Administrator accounts."
                Else
                    ' Check for active borrowings
                    Dim rsActiveLoans
                    Set rsActiveLoans = conn.Execute("SELECT COUNT(*) AS active_count FROM Borrowings WHERE user_id = " & deleteId & " AND status = 'Borrowed'")
                    Dim activeCount
                    activeCount = 0
                    If Not (rsActiveLoans Is Nothing Or rsActiveLoans.State = 0) Then
                        activeCount = SafeInt(rsActiveLoans("active_count"), 0)
                        rsActiveLoans.Close
                    End If

                    If activeCount > 0 Then
                        errorMessage = "Cannot delete user '" & CleanText(targetName) & "' because they have " & activeCount & " active unreturned book loan(s)."
                    Else
                        On Error Resume Next
                        ' Soft delete user: preserve borrowing history, book requests, and audit trail
                        conn.Execute("UPDATE Users SET is_deleted = 1 WHERE id = " & deleteId)

                        If Err.Number <> 0 Then
                            errorMessage = "Failed to delete user: " & Err.Description
                        Else
                            successMessage = "User '" & CleanText(targetName) & "' was successfully removed (audit history preserved)."
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    End If
End If

' Fetch all registered active users (excluding soft-deleted)
Dim rsUsers
Set rsUsers = conn.Execute("SELECT id, username, staff_number, staff_internal_number, sbu, dob, email, role, created_at FROM Users WHERE COALESCE(is_deleted, 0) = 0 ORDER BY FIELD(role, 'Admin', 'Librarian', 'Staff'), username ASC")

RenderHeader "Staff & User Management"
%>

<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
    <div>
        <h2 style="margin:0; color:#2c3e50;">👥 Staff &amp; User Management</h2>
        <p style="margin:4px 0 0 0; color:#7f8c8d; font-size:14px;">Register and manage library staff and employee accounts</p>
    </div>
</div>

<% If errorMessage <> "" Then %>
    <div class="alert alert-danger"><%= Server.HTMLEncode(errorMessage) %></div>
<% End If %>

<% If successMessage <> "" Then %>
    <div class="alert alert-info" style="background:#d4edda; color:#155724; border-color:#c3e6cb;">
        <%= Server.HTMLEncode(successMessage) %>
    </div>
<% End If %>

<div style="display:grid; grid-template-columns: 340px 1fr; gap:24px; align-items:start;">
    <!-- Add User Form -->
    <div class="card" style="box-shadow:0 2px 10px rgba(0,0,0,0.06); padding:22px;">
        <h3 style="margin-top:0; color:#2c3e50; border-bottom:2px solid #3498db; padding-bottom:8px;">
            <span>➕</span> Add New User
        </h3>
        
        <form action="users.asp" method="POST" id="addUserForm">
            <input type="hidden" name="action_type" value="add">

            <div class="form-group">
                <label for="username">Full Name *</label>
                <input type="text" id="username" name="username" required placeholder="e.g. Ramesh Kumar">
            </div>

            <div class="form-group">
                <label for="staff_number">Staff Number (10 Digits) *</label>
                <input type="text" id="staff_number" name="staff_number" required placeholder="e.g. 9876543216" pattern="[0-9]{6,12}" title="Staff Number should be 6 to 12 digits">
                <small style="color:#7f8c8d; font-size:11.5px; display:block; margin-top:3px;">
                    Used by the employee to log in.
                </small>
            </div>

            <div class="form-group">
                <label for="internal_no">Internal No *</label>
                <input type="text" id="internal_no" name="internal_no" required placeholder="e.g. BEL-EXT-101 / 4521">
                <small style="color:#7f8c8d; font-size:11.5px; display:block; margin-top:3px;">
                    Intercom / internal department extension number.
                </small>
            </div>

            <div class="form-group">
                <label for="sbu">Strategic Business Unit (SBU) *</label>
                <select id="sbu" name="sbu" required style="width:100%; padding:8px 12px; border:1px solid #ccc; border-radius:4px; font-size:13.5px; background:#fff;">
                    <option value="" disabled selected>-- Select Strategic Business Unit (SBU) --</option>
                    <option value="Radar SBU">Radar SBU (Weapon Sensors &amp; Radars)</option>
                    <option value="MilCom SBU">MilCom SBU (Military Communications)</option>
                    <option value="Naval Systems">Naval Systems (Sonar &amp; Underwater Defense)</option>
                    <option value="Missile Systems">Missile Systems</option>
                    <option value="Electronic Warfare (EW)">Electronic Warfare (EW)</option>
                    <option value="Opto-Electronics">Opto-Electronics (Mil-Optics &amp; Lasers)</option>
                    <option value="CRL (Research)">Central Research Laboratory (CRL)</option>
                    <option value="Software SBU">Software SBU (Defense Software)</option>
                    <option value="Homeland Security">Homeland Security &amp; Smart Cities</option>
                    <option value="Central HR">Central HR &amp; Administration</option>
                    <option value="Quality Assurance (QA)">Quality Assurance (QA)</option>
                    <option value="Supply Chain / Finance">Supply Chain &amp; Finance</option>
                    <option value="General">General / Administration</option>
                </select>
                <small style="color:#7f8c8d; font-size:11.5px; display:block; margin-top:3px;">
                    Select the employee's assigned Strategic Business Unit.
                </small>
            </div>

            <div class="form-group">
                <label for="dob">Date of Birth (Login Password) *</label>
                <input type="text" id="dob" name="dob" required placeholder="DD-MM-YYYY (e.g. 15-08-1996)">
                <small style="color:#2980b9; font-size:11.5px; display:block; margin-top:3px; font-weight:600;">
                    💡 This DOB automatically serves as their login password!
                </small>
            </div>

            <div class="form-group">
                <label for="email">Email Address *</label>
                <input type="email" id="email" name="email" required placeholder="ramesh.k@bel.com">
            </div>

            <div class="form-group">
                <label for="role">User Role *</label>
                <select id="role" name="role" required style="width:100%; padding:8px 12px; border:1px solid #ccc; border-radius:4px;">
                    <option value="Staff" selected>Staff (Member)</option>
                    <option value="Librarian">Librarian (Manager)</option>
                    <% If IsAdmin() Then %>
                        <option value="Admin">Administrator (Full Access)</option>
                    <% End If %>
                </select>
            </div>

            <button type="submit" class="btn btn-success" style="width:100%; padding:10px; font-size:14.5px; font-weight:600; margin-top:6px;">
                <span>💾</span> Register User
            </button>
        </form>
    </div>

    <!-- Users List Table -->
    <div class="card" style="box-shadow:0 2px 10px rgba(0,0,0,0.06); padding:22px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid #eee; padding-bottom:10px;">
            <h3 style="margin:0; color:#2c3e50;">Registered Staff Directory</h3>
            <input type="text" id="userSearch" placeholder="🔍 Filter users..." onkeyup="filterUsers()" style="padding:6px 12px; border:1px solid #ccc; border-radius:4px; font-size:13px; width:220px;">
        </div>

        <table id="usersTable" style="width:100%; border-collapse:collapse;">
            <thead>
                <tr style="background:#f8f9fa;">
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">Staff Member</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">Staff Number</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">Internal No</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">SBU</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">Role</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">DOB (Password)</th>
                    <th style="padding:10px; text-align:left; border-bottom:2px solid #dee2e6;">Email</th>
                    <th style="padding:10px; text-align:center; border-bottom:2px solid #dee2e6;">Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsUsers Is Nothing Or rsUsers.State = 0 Or rsUsers.EOF Then
                %>
                    <tr>
                        <td colspan="8" style="text-align:center; padding:20px; color:#95a5a6;">No registered users found.</td>
                    </tr>
                <%
                Else
                    Dim uRole, uBadgeBg
                    Do While Not rsUsers.EOF
                        uRole = Trim(rsUsers("role") & "")
                        Select Case LCase(uRole)
                            Case "admin"
                                uBadgeBg = "#e74c3c"
                            Case "librarian"
                                uBadgeBg = "#8e44ad"
                            Case Else
                                uBadgeBg = "#27ae60"
                        End Select
                %>
                    <tr style="border-bottom:1px solid #eee;">
                        <td style="padding:10px;">
                            <strong><%= Server.HTMLEncode(rsUsers("username") & "") %></strong>
                        </td>
                        <td style="padding:10px;">
                            <code style="font-weight:700; color:#2c3e50; font-size:13.5px;"><%= Server.HTMLEncode(rsUsers("staff_number") & "") %></code>
                        </td>
                        <td style="padding:10px;">
                            <span style="font-family:monospace; font-weight:600; color:#0284c7; background:#e0f2fe; padding:2px 8px; border-radius:4px; font-size:12.5px;">
                                <%= Server.HTMLEncode(rsUsers("staff_internal_number") & "") %>
                            </span>
                        </td>
                        <td style="padding:10px;">
                            <span style="font-size:12px; font-weight:600; color:#4338ca; background:#e0e7ff; padding:3px 9px; border-radius:12px; display:inline-block;">
                                <%= Server.HTMLEncode(rsUsers("sbu") & "") %>
                            </span>
                        </td>
                        <td style="padding:10px;">
                            <% If LCase(uRole) = "admin" Then %>
                                <span class="status-pill status-pill-red" style="font-size:12px; padding:4px 14px;">Admin</span>
                            <% ElseIf LCase(uRole) = "librarian" Then %>
                                <span class="status-pill status-pill-purple" style="font-size:12px; padding:4px 14px;">Librarian</span>
                            <% Else %>
                                <span class="status-pill status-pill-blue" style="font-size:12px; padding:4px 14px;">Staff</span>
                            <% End If %>
                        </td>
                        <td style="padding:10px;">
                            <code style="color:#2980b9; font-weight:600; font-size:13px;"><%= Server.HTMLEncode(rsUsers("dob") & "") %></code>
                        </td>
                        <td style="padding:10px; font-size:13px; color:#555;">
                            <%= Server.HTMLEncode(rsUsers("email") & "") %>
                        </td>
                        <td style="padding:10px; text-align:center;">
                            <% If SafeInt(rsUsers("id"), 0) = SafeInt(Session("UserId"), 0) Then %>
                                <span class="action-text-muted" style="font-style:italic;">(You)</span>
                            <% ElseIf (Not IsAdmin()) And LCase(uRole) = "admin" Then %>
                                <span class="action-text-muted">Protected</span>
                            <% Else %>
                                <div style="display:inline-block; width:85px;">
                                    <form action="users.asp" method="POST" style="margin:0;" onsubmit="return confirm('Are you sure you want to delete user \'<%= Replace(rsUsers("username") & "", "'", "\'") %>\'?');">
                                        <input type="hidden" name="action_type" value="delete">
                                        <input type="hidden" name="user_id" value="<%= rsUsers("id") %>">
                                        <button type="submit" class="action-btn action-btn-delete" style="padding:5px 10px; font-size:12px;">Delete</button>
                                    </form>
                                </div>
                            <% End If %>
                        </td>
                    </tr>
                <%
                        rsUsers.MoveNext
                    Loop
                    rsUsers.Close
                End If

                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>
</div>

<script>
function filterUsers() {
    var input = document.getElementById("userSearch");
    var filter = input.value.toLowerCase();
    var table = document.getElementById("usersTable");
    var tr = table.getElementsByTagName("tr");

    for (var i = 1; i < tr.length; i++) {
        var tdText = tr[i].textContent || tr[i].innerText;
        if (tdText.toLowerCase().indexOf(filter) > -1) {
            tr[i].style.display = "";
        } else {
            tr[i].style.display = "none";
        }
    }
}
</script>

<% RenderFooter %>
