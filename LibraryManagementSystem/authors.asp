<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

If Not IsStaff() Then
    Response.Redirect("dashboard.asp?msg=Access+denied")
    Response.End
End If

Dim conn, errorMessage, successMessage
errorMessage = ""
successMessage = ""
Set conn = GetConnection()

' Handle Form Submissions (Add / Delete Author)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim actionType
    actionType = Request.Form("action_type")

    If actionType = "add" Then
        Dim name, bio
        name = Trim(Request.Form("name"))
        bio = Trim(Request.Form("bio"))

        If name = "" Then
            errorMessage = "Author Name is required."
        Else
            On Error Resume Next
            conn.Execute("INSERT INTO Authors (name, bio) VALUES (" & SQLQuote(name) & ", " & SQLQuote(bio) & ")")
            If Err.Number <> 0 Then
                errorMessage = "Failed to add author: " & Err.Description
            Else
                successMessage = "Author added successfully."
            End If
            On Error GoTo 0
        End If

    ElseIf actionType = "delete" Then
        Dim authorId
        authorId = SafeInt(Request.Form("author_id"), 0)
        If authorId > 0 Then
            On Error Resume Next
            conn.Execute("DELETE FROM Authors WHERE id = " & authorId)
            If Err.Number <> 0 Then
                errorMessage = "Failed to delete author: " & Err.Description
            Else
                successMessage = "Author deleted successfully."
            End If
            On Error GoTo 0
        End If
    End If
End If

' Fetch Authors List
Dim rsAuthors
Set rsAuthors = conn.Execute("SELECT id, name, bio FROM Authors ORDER BY name ASC")

RenderHeader "Authors Management"
%>

<div style="display:grid; grid-template-columns: 1fr 2fr; gap:24px;">
    <!-- Add Author Form -->
    <div class="card">
        <h3 style="margin-top:0; color:#2c3e50;">Add Author</h3>

        <% If errorMessage <> "" Then %>
            <div class="alert alert-danger"><%= Server.HTMLEncode(errorMessage) %></div>
        <% End If %>

        <% If successMessage <> "" Then %>
            <div class="alert alert-info"><%= Server.HTMLEncode(successMessage) %></div>
        <% End If %>

        <form action="authors.asp" method="POST">
            <input type="hidden" name="action_type" value="add">

            <div class="form-group">
                <label for="name">Author Name *</label>
                <input type="text" id="name" name="name" required placeholder="e.g. Arthur Conan Doyle">
            </div>

            <div class="form-group">
                <label for="bio">Biography</label>
                <textarea id="bio" name="bio" rows="4" placeholder="Brief author background..."></textarea>
            </div>

            <button type="submit" class="btn btn-success" style="width:100%;">Save Author</button>
        </form>
    </div>

    <!-- Authors List Table -->
    <div class="card">
        <h3 style="margin-top:0; color:#2c3e50;">Author Directory</h3>

        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Name</th>
                    <th>Biography</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsAuthors Is Nothing Or rsAuthors.State = 0 Or rsAuthors.EOF Then
                %>
                    <tr>
                        <td colspan="4" style="text-align:center; color:#95a5a6;">No authors registered.</td>
                    </tr>
                <%
                Else
                    Do While Not rsAuthors.EOF
                %>
                    <tr>
                        <td>#<%= rsAuthors("id") %></td>
                        <td><strong><%= Server.HTMLEncode(rsAuthors("name") & "") %></strong></td>
                        <td><%= Server.HTMLEncode(rsAuthors("bio") & "") %></td>
                        <td>
                            <form action="authors.asp" method="POST" style="display:inline;" onsubmit="return confirm('Delete this author?');">
                                <input type="hidden" name="action_type" value="delete">
                                <input type="hidden" name="author_id" value="<%= rsAuthors("id") %>">
                                <button type="submit" class="btn btn-danger" style="padding:4px 8px; font-size:12px;">Delete</button>
                            </form>
                        </td>
                    </tr>
                <%
                        rsAuthors.MoveNext
                    Loop
                    rsAuthors.Close
                End If

                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>
</div>

<% RenderFooter %>
