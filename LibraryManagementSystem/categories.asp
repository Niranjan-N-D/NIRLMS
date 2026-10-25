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

' Handle Form Submissions (Add / Delete Category)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim actionType
    actionType = Request.Form("action_type")

    If actionType = "add" Then
        Dim name, description
        name = Trim(Request.Form("name"))
        description = Trim(Request.Form("description"))

        If name = "" Then
            errorMessage = "Category Name is required."
        Else
            On Error Resume Next
            conn.Execute("INSERT INTO Categories (name, description) VALUES (" & SQLQuote(name) & ", " & SQLQuote(description) & ")")
            If Err.Number <> 0 Then
                errorMessage = "Failed to add category: " & Err.Description
            Else
                successMessage = "Category added successfully."
            End If
            On Error GoTo 0
        End If

    ElseIf actionType = "delete" Then
        Dim categoryId
        categoryId = SafeInt(Request.Form("category_id"), 0)
        If categoryId > 0 Then
            On Error Resume Next
            conn.Execute("DELETE FROM Categories WHERE id = " & categoryId)
            If Err.Number <> 0 Then
                errorMessage = "Failed to delete category: " & Err.Description
            Else
                successMessage = "Category deleted successfully."
            End If
            On Error GoTo 0
        End If
    End If
End If

' Fetch Categories List
Dim rsCategories
Set rsCategories = conn.Execute("SELECT id, name, description FROM Categories ORDER BY name ASC")

RenderHeader "Categories Management"
%>

<div style="display:grid; grid-template-columns: 1fr 2fr; gap:24px;">
    <!-- Add Category Form -->
    <div class="card">
        <h3 style="margin-top:0; color:#2c3e50;">Add Category</h3>

        <% If errorMessage <> "" Then %>
            <div class="alert alert-danger"><%= Server.HTMLEncode(errorMessage) %></div>
        <% End If %>

        <% If successMessage <> "" Then %>
            <div class="alert alert-info"><%= Server.HTMLEncode(successMessage) %></div>
        <% End If %>

        <form action="categories.asp" method="POST">
            <input type="hidden" name="action_type" value="add">

            <div class="form-group">
                <label for="name">Category Name *</label>
                <input type="text" id="name" name="name" required placeholder="e.g. Computer Science">
            </div>

            <div class="form-group">
                <label for="description">Description</label>
                <textarea id="description" name="description" rows="4" placeholder="Brief category description..."></textarea>
            </div>

            <button type="submit" class="btn btn-success" style="width:100%;">Save Category</button>
        </form>
    </div>

    <!-- Categories List Table -->
    <div class="card">
        <h3 style="margin-top:0; color:#2c3e50;">Book Categories</h3>

        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Name</th>
                    <th>Description</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                If rsCategories Is Nothing Or rsCategories.State = 0 Or rsCategories.EOF Then
                %>
                    <tr>
                        <td colspan="4" style="text-align:center; color:#95a5a6;">No categories registered.</td>
                    </tr>
                <%
                Else
                    Do While Not rsCategories.EOF
                %>
                    <tr>
                        <td>#<%= rsCategories("id") %></td>
                        <td><strong><%= Server.HTMLEncode(rsCategories("name") & "") %></strong></td>
                        <td><%= Server.HTMLEncode(rsCategories("description") & "") %></td>
                        <td>
                            <form action="categories.asp" method="POST" style="display:inline;" onsubmit="return confirm('Delete this category?');">
                                <input type="hidden" name="action_type" value="delete">
                                <input type="hidden" name="category_id" value="<%= rsCategories("id") %>">
                                <button type="submit" class="btn btn-danger" style="padding:4px 8px; font-size:12px;">Delete</button>
                            </form>
                        </td>
                    </tr>
                <%
                        rsCategories.MoveNext
                    Loop
                    rsCategories.Close
                End If

                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>
</div>

<% RenderFooter %>
