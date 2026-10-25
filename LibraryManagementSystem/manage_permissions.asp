<!--#include file="includes/db_config.asp"-->
<%
' Require Authentication & Admin Security Guard
RequireAuth()
If Not IsAdmin() Then
    Response.Redirect("dashboard.asp?msg=Access+Denied:+Admin+only")
    Response.End
End If

Dim conn, errorMessage, successMessage
errorMessage = ""
successMessage = ""
Set conn = GetConnection()

' Handle Permission Updates (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userId, pBooks, pRequests, pBorrowings, pAuthors, pCategories, pFeedback
    
    On Error Resume Next
    Dim rsUsersList
    Set rsUsersList = conn.Execute("SELECT id FROM Users WHERE role <> 'Admin'")
    
    If Not (rsUsersList Is Nothing Or rsUsersList.State = 0) Then
        Do While Not rsUsersList.EOF
            userId = rsUsersList("id")
            
            pBooks = IIf(Request.Form("perm_books_" & userId) = "1", 1, 0)
            pRequests = IIf(Request.Form("perm_requests_" & userId) = "1", 1, 0)
            pBorrowings = IIf(Request.Form("perm_borrowings_" & userId) = "1", 1, 0)
            pAuthors = IIf(Request.Form("perm_authors_" & userId) = "1", 1, 0)
            pCategories = IIf(Request.Form("perm_categories_" & userId) = "1", 1, 0)
            pFeedback = IIf(Request.Form("perm_feedback_" & userId) = "1", 1, 0)
            
            conn.Execute("UPDATE Users SET " & _
                "perm_books = " & pBooks & ", " & _
                "perm_requests = " & pRequests & ", " & _
                "perm_borrowings = " & pBorrowings & ", " & _
                "perm_authors = " & pAuthors & ", " & _
                "perm_categories = " & pCategories & ", " & _
                "perm_feedback = " & pFeedback & " " & _
                "WHERE id = " & userId)
                
            rsUsersList.MoveNext
        Loop
        rsUsersList.Close
    End If
    
    If Err.Number <> 0 Then
        errorMessage = "Error updating permissions: " & Err.Description
    Else
        successMessage = "User feature permissions updated successfully!"
    End If
    On Error GoTo 0
End If

' Fetch all non-admin users
Dim rsUsers
Set rsUsers = conn.Execute("SELECT id, username, email, role, perm_books, perm_requests, perm_borrowings, perm_authors, perm_categories, perm_feedback FROM Users ORDER BY role ASC, username ASC")

RenderHeader "Manage Permissions"
%>

<div class="card">
    <div class="card-header">
        <div>
            <h2 style="margin:0; color:#2c3e50;">⚙️ Role & Feature Permissions Management</h2>
            <p style="margin: 5px 0 0 0; color:#7f8c8d; font-size:14px;">As Admin, use checkboxes to grant or revoke specific system features for Librarians and Members.</p>
        </div>
    </div>

    <% If errorMessage <> "" Then %>
        <div class="alert alert-danger"><%= Server.HTMLEncode(errorMessage) %></div>
    <% End If %>

    <% If successMessage <> "" Then %>
        <div class="alert alert-info"><%= Server.HTMLEncode(successMessage) %></div>
    <% End If %>

    <form action="manage_permissions.asp" method="POST">
        <table>
            <thead>
                <tr>
                    <th>User / Role</th>
                    <th>Email</th>
                    <th style="text-align:center;">📚 Books</th>
                    <th style="text-align:center;">📥 Requests</th>
                    <th style="text-align:center;">📋 Borrowings</th>
                    <th style="text-align:center;">✍️ Authors</th>
                    <th style="text-align:center;">🏷️ Categories</th>
                    <th style="text-align:center;">⭐ Feedback</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsUsers Is Nothing Or rsUsers.State = 0 Or rsUsers.EOF Then
                %>
                    <tr>
                        <td colspan="8" style="text-align:center; color:#95a5a6;">No users found in database.</td>
                    </tr>
                <%
                Else
                    Do While Not rsUsers.EOF
                        Dim uId, uRole, roleColor
                        uId = rsUsers("id")
                        uRole = rsUsers("role")
                        
                        If uRole = "Admin" Then
                            roleColor = "#e74c3c"
                        ElseIf uRole = "Librarian" Then
                            roleColor = "#9b59b6"
                        Else
                            roleColor = "#2ecc71"
                        End If
                %>
                    <tr>
                        <td>
                            <strong><%= Server.HTMLEncode(rsUsers("username") & "") %></strong><br>
                            <span class="badge" style="background:<%= roleColor %>; font-size:11px;"><%= uRole %></span>
                        </td>
                        <td><%= Server.HTMLEncode(rsUsers("email") & "") %></td>
                        <% If uRole = "Admin" Then %>
                            <td colspan="6" style="text-align:center; color:#e74c3c; font-weight:bold;">Full Authority (All Features Granted)</td>
                        <% Else %>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_books_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_books")), "checked", "") %>>
                            </td>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_requests_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_requests")), "checked", "") %>>
                            </td>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_borrowings_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_borrowings")), "checked", "") %>>
                            </td>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_authors_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_authors")), "checked", "") %>>
                            </td>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_categories_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_categories")), "checked", "") %>>
                            </td>
                            <td style="text-align:center;">
                                <input type="checkbox" name="perm_feedback_<%= uId %>" value="1" <%= IIf(CBool(rsUsers("perm_feedback")), "checked", "") %>>
                            </td>
                        <% End If %>
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

        <div style="margin-top:20px; text-align:right;">
            <button type="submit" class="btn btn-success" style="padding:10px 24px; font-size:16px;">💾 Save Permission Changes</button>
        </div>
    </form>
</div>

<% RenderFooter %>
