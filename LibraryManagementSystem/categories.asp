<!--#include file="includes/db_config.asp"-->
<%
RequireAuth()

Dim conn, errorMessage, successMessage
errorMessage = ""
successMessage = ""
Set conn = GetConnection()

' Handle Form Submissions (Add / Delete Category) - Staff Only
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not IsStaff() Then
        Response.Redirect("categories.asp?msg=Access+denied")
        Response.End
    End If

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

' Fetch Categories List with Book Counts & Available Copies
Dim rsCategories, totalCategoriesCount, totalBooksInAll, totalAvailableInAll
totalCategoriesCount = 0
totalBooksInAll = 0
totalAvailableInAll = 0

Dim catSql
catSql = "SELECT c.id, c.name, c.description, " & _
         "COUNT(b.id) AS total_books, " & _
         "COALESCE(SUM(b.copies_available), 0) AS total_available " & _
         "FROM Categories c " & _
         "LEFT JOIN Books b ON c.id = b.category_id " & _
         "GROUP BY c.id, c.name, c.description " & _
         "ORDER BY c.name ASC"

Set rsCategories = conn.Execute(catSql)

' Function to determine a tailored icon for each category domain
Function GetCategoryIcon(catName)
    Dim n
    n = LCase(catName & "")
    If InStr(n, "radar") > 0 Or InStr(n, "defense") > 0 Then
        GetCategoryIcon = "📡"
    ElseIf InStr(n, "telecom") > 0 Or InStr(n, "signal") > 0 Or InStr(n, "communication") > 0 Then
        GetCategoryIcon = "📶"
    ElseIf InStr(n, "software") > 0 Or InStr(n, "embedded") > 0 Or InStr(n, "computer") > 0 Then
        GetCategoryIcon = "💻"
    ElseIf InStr(n, "aero") > 0 Or InStr(n, "space") > 0 Then
        GetCategoryIcon = "✈️"
    ElseIf InStr(n, "kannada") > 0 Or InStr(n, "literature") > 0 Or InStr(n, "ಕನ್ನಡ") > 0 Then
        GetCategoryIcon = "📖"
    ElseIf InStr(n, "leader") > 0 Or InStr(n, "manage") > 0 Or InStr(n, "business") > 0 Then
        GetCategoryIcon = "👔"
    ElseIf InStr(n, "science") > 0 Or InStr(n, "physics") > 0 Then
        GetCategoryIcon = "🔬"
    ElseIf InStr(n, "math") > 0 Then
        GetCategoryIcon = "📐"
    Else
        GetCategoryIcon = "🏷️"
    End If
End Function

RenderHeader IIf(IsStaff(), "Categories Management", "Book Categories")
%>

<style>
/* Category Catalog Modern Styling */
.cat-stat-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}
.cat-stat-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 18px 20px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.cat-stat-icon {
    width: 48px;
    height: 48px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    flex-shrink: 0;
}
.cat-stat-val {
    font-size: 24px;
    font-weight: 800;
    color: #0B1320;
    line-height: 1.1;
}
.cat-stat-lbl {
    font-size: 13px;
    color: #64748b;
    margin-top: 3px;
    font-weight: 500;
}

/* Category Cards Grid */
.category-cards-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
    gap: 20px;
    margin-bottom: 30px;
}
.category-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 22px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    position: relative;
    overflow: hidden;
}
.category-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 10px 24px rgba(11, 19, 32, 0.08);
    border-color: #3b82f6;
}
.category-card-top {
    display: flex;
    align-items: flex-start;
    gap: 14px;
    margin-bottom: 14px;
}
.category-badge-icon {
    width: 46px;
    height: 46px;
    border-radius: 10px;
    background: #f1f5f9;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 24px;
    flex-shrink: 0;
}
.category-card-title {
    font-size: 17px;
    font-weight: 700;
    color: #0B1320;
    margin: 0 0 4px 0;
    line-height: 1.3;
}
.category-card-desc {
    font-size: 13.5px;
    color: #64748b;
    line-height: 1.5;
    margin: 0 0 16px 0;
    min-height: 40px;
}
.category-pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-bottom: 18px;
    padding-top: 10px;
    border-top: 1px dashed #e2e8f0;
}
.cat-pill {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 5px 10px;
    border-radius: 6px;
    font-size: 12.5px;
    font-weight: 600;
}
.cat-pill-books {
    background: #e0f2fe;
    color: #0369a1;
}
.cat-pill-avail {
    background: #dcfce7;
    color: #15803d;
}
.cat-pill-out {
    background: #fee2e2;
    color: #b91c1c;
}
.btn-category-action {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    width: 100%;
    padding: 10px 16px;
    background: linear-gradient(135deg, #0B1320 0%, #1e3a8a 100%);
    color: #ffffff !important;
    text-decoration: none;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    transition: opacity 0.15s ease, transform 0.1s ease;
    box-sizing: border-box;
    box-shadow: 0 2px 4px rgba(11, 19, 32, 0.15);
}
.btn-category-action:hover {
    opacity: 0.94;
    transform: translateY(-1px);
}
.btn-category-empty {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    width: 100%;
    padding: 10px 16px;
    background: #f1f5f9;
    color: #94a3b8 !important;
    text-decoration: none;
    border-radius: 8px;
    font-size: 13.5px;
    font-weight: 600;
    box-sizing: border-box;
    cursor: not-allowed;
}
</style>

<% If errorMessage <> "" Then %>
    <div class="alert alert-danger" style="margin-bottom:20px;"><%= Server.HTMLEncode(errorMessage) %></div>
<% End If %>

<% If successMessage <> "" Then %>
    <div class="alert alert-info" style="margin-bottom:20px;"><%= Server.HTMLEncode(successMessage) %></div>
<% End If %>

<%
' Convert RS to Array or Cache for multiple iterations (Cards & Table)
Dim catArray(), catCount
catCount = 0

If Not (rsCategories Is Nothing Or rsCategories.State = 0 Or rsCategories.EOF) Then
    Dim tmpList
    tmpList = rsCategories.GetRows()
    rsCategories.Close
    
    ' Dimensions of GetRows: (columns, rows)
    ' 0: id, 1: name, 2: description, 3: total_books, 4: total_available
    catCount = UBound(tmpList, 2) + 1
    totalCategoriesCount = catCount
    
    Dim i
    For i = 0 To catCount - 1
        totalBooksInAll = totalBooksInAll + SafeInt(tmpList(3, i), 0)
        totalAvailableInAll = totalAvailableInAll + SafeInt(tmpList(4, i), 0)
    Next
End If
%>

<!-- Summary Stats Header -->
<div class="cat-stat-grid">
    <div class="cat-stat-card">
        <div class="cat-stat-icon" style="background:#e0f2fe; color:#0369a1;">🏷️</div>
        <div>
            <div class="cat-stat-val"><%= totalCategoriesCount %></div>
            <div class="cat-stat-lbl">Active Categories</div>
        </div>
    </div>
    <div class="cat-stat-card">
        <div class="cat-stat-icon" style="background:#fef3c7; color:#d97706;">📚</div>
        <div>
            <div class="cat-stat-val"><%= totalBooksInAll %></div>
            <div class="cat-stat-lbl">Total Books Across Categories</div>
        </div>
    </div>
    <div class="cat-stat-card">
        <div class="cat-stat-icon" style="background:#dcfce7; color:#15803d;">🟢</div>
        <div>
            <div class="cat-stat-val"><%= totalAvailableInAll %></div>
            <div class="cat-stat-lbl">Copies Available to Request</div>
        </div>
    </div>
</div>

<% If IsStaff() Then %>
    <!-- STAFF VIEW: Add Form & Table with Admin Actions -->
    <div style="display:grid; grid-template-columns: 1fr 2fr; gap:24px; align-items:start;">
        <!-- Add Category Form -->
        <div class="card" style="margin:0;">
            <h3 style="margin-top:0; color:#0B1320; font-size:18px;">➕ Add New Category</h3>
            <p style="font-size:13px; color:#64748b; margin-top:-6px; margin-bottom:16px;">Create a new subject classification for library titles.</p>

            <form action="categories.asp" method="POST">
                <input type="hidden" name="action_type" value="add">

                <div class="form-group">
                    <label for="name">Category Name *</label>
                    <input type="text" id="name" name="name" required placeholder="e.g. Artificial Intelligence">
                </div>

                <div class="form-group">
                    <label for="description">Description</label>
                    <textarea id="description" name="description" rows="4" placeholder="Brief subject description, scopes, and domains..."></textarea>
                </div>

                <button type="submit" class="btn btn-success" style="width:100%; font-weight:700; padding:10px;">Save Category</button>
            </form>
        </div>

        <!-- Categories Table for Staff -->
        <div class="card" style="margin:0;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
                <h3 style="margin:0; color:#0B1320; font-size:18px;">Categories Master</h3>
                <span style="font-size:13px; color:#64748b;"><%= catCount %> registered categories</span>
            </div>

            <table>
                <thead>
                    <tr>
                        <th style="width:50px;">ID</th>
                        <th>Category Name</th>
                        <th>Description</th>
                        <th style="text-align:center;">Books</th>
                        <th style="text-align:center;">Available</th>
                        <th style="text-align:center; width:150px;">Action</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                    If catCount = 0 Then
                    %>
                        <tr>
                            <td colspan="6" style="text-align:center; color:#95a5a6; padding:20px;">No categories registered yet.</td>
                        </tr>
                    <%
                    Else
                        For i = 0 To catCount - 1
                            Dim cId, cName, cDesc, cBooks, cAvail
                            cId = tmpList(0, i)
                            cName = tmpList(1, i) & ""
                            cDesc = tmpList(2, i) & ""
                            cBooks = SafeInt(tmpList(3, i), 0)
                            cAvail = SafeInt(tmpList(4, i), 0)
                    %>
                        <tr>
                            <td><code>#<%= cId %></code></td>
                            <td>
                                <a href="books.asp?category_id=<%= cId %>" style="color:#0B1320; font-weight:700; text-decoration:none;" title="Click to view books in <%= CleanText(cName) %>">
                                    <%= GetCategoryIcon(cName) %> <%= CleanText(cName) %>
                                </a>
                            </td>
                            <td><span style="font-size:13px; color:#64748b;"><%= CleanText(cDesc) %></span></td>
                            <td style="text-align:center;">
                                <span class="cat-pill cat-pill-books"><%= cBooks %> titles</span>
                            </td>
                            <td style="text-align:center;">
                                <% If cAvail > 0 Then %>
                                    <span class="cat-pill cat-pill-avail"><%= cAvail %> copies</span>
                                <% Else %>
                                    <span class="cat-pill cat-pill-out">0 copies</span>
                                <% End If %>
                            </td>
                            <td style="text-align:center;">
                                <div style="display:flex; gap:6px; justify-content:center;">
                                    <a href="books.asp?category_id=<%= cId %>" class="action-btn" style="background:#0B1320; color:#fff; padding:5px 8px; font-size:12px; border-radius:5px; text-decoration:none;" title="Browse and view all books">
                                        Browse (<%= cBooks %>)
                                    </a>
                                    <form action="categories.asp" method="POST" style="margin:0; display:inline;" onsubmit="return confirm('Delete this category? Books linked to this category will have category unassigned.');">
                                        <input type="hidden" name="action_type" value="delete">
                                        <input type="hidden" name="category_id" value="<%= cId %>">
                                        <button type="submit" class="action-btn action-btn-delete" style="padding:5px 8px; font-size:12px; border-radius:5px;" title="Delete Category">Delete</button>
                                    </form>
                                </div>
                            </td>
                        </tr>
                    <%
                        Next
                    End If
                    %>
                </tbody>
            </table>
        </div>
    </div>

<% Else %>

    <!-- MEMBER / REGULAR USER VIEW: Rich Category Cards & Direct Book Request Portal -->
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:18px 24px; margin-bottom:24px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
        <div>
            <h3 style="margin:0 0 4px 0; color:#0B1320; font-size:19px;">Explore Books by Category</h3>
            <p style="margin:0; font-size:14px; color:#64748b;">Select a category to view all registered titles, check live shelf availability, and submit a book request.</p>
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
            <input type="text" id="catSearchInput" placeholder="🔍 Filter categories..." onkeyup="filterCategories()" style="padding:8px 14px; font-size:14px; border:1px solid #cbd5e1; border-radius:8px; width:220px;">
        </div>
    </div>

    <!-- Category Cards Grid -->
    <div class="category-cards-grid" id="categoryCardsContainer">
        <%
        If catCount = 0 Then
        %>
            <div style="grid-column: 1 / -1; text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px dashed #cbd5e1; color:#94a3b8;">
                <span style="font-size:36px; display:block; margin-bottom:10px;">🏷️</span>
                <h3>No categories currently available</h3>
                <p>Please check back later or contact library administration.</p>
            </div>
        <%
        Else
            For i = 0 To catCount - 1
                Dim mcId, mcName, mcDesc, mcBooks, mcAvail
                mcId = tmpList(0, i)
                mcName = tmpList(1, i) & ""
                mcDesc = tmpList(2, i) & ""
                mcBooks = SafeInt(tmpList(3, i), 0)
                mcAvail = SafeInt(tmpList(4, i), 0)
        %>
            <div class="category-card" data-catname="<%= LCase(CleanText(mcName)) %>" data-catdesc="<%= LCase(CleanText(mcDesc)) %>">
                <div>
                    <div class="category-card-top">
                        <div class="category-badge-icon">
                            <%= GetCategoryIcon(mcName) %>
                        </div>
                        <div style="flex:1;">
                            <h4 class="category-card-title"><%= CleanText(mcName) %></h4>
                            <span style="font-size:12px; color:#94a3b8; font-weight:600;">Category #<%= mcId %></span>
                        </div>
                    </div>

                    <p class="category-card-desc">
                        <% If mcDesc <> "" Then %>
                            <%= CleanText(mcDesc) %>
                        <% Else %>
                            <em style="color:#cbd5e1;">Curated collection of books in <%= CleanText(mcName) %> discipline.</em>
                        <% End If %>
                    </p>

                    <div class="category-pill-row">
                        <span class="cat-pill cat-pill-books" title="Total registered books in this category">
                            📚 <%= mcBooks %> <%= IIf(mcBooks = 1, "Book", "Books") %>
                        </span>
                        <% If mcAvail > 0 Then %>
                            <span class="cat-pill cat-pill-avail" title="Copies ready for immediate request">
                                🟢 <%= mcAvail %> Available
                            </span>
                        <% ElseIf mcBooks > 0 Then %>
                            <span class="cat-pill cat-pill-out" title="All copies are currently borrowed or reserved">
                                ⏱️ Reserved / Loaned Out
                            </span>
                        <% Else %>
                            <span class="cat-pill" style="background:#f1f5f9; color:#94a3b8;">
                                ⚪ 0 Titles
                            </span>
                        <% End If %>
                    </div>
                </div>

                <div>
                    <% If mcBooks > 0 Then %>
                        <a href="books.asp?category_id=<%= mcId %>" class="btn-category-action" title="View all books in <%= CleanText(mcName) %> and request books">
                            <span>Browse Books &amp; Request</span>
                            <span style="font-size:16px;">&rarr;</span>
                        </a>
                    <% Else %>
                        <div class="btn-category-empty" title="No books added under this category yet">
                            <span>No Books in Category Yet</span>
                        </div>
                    <% End If %>
                </div>
            </div>
        <%
            Next
        End If
        %>
    </div>

    <!-- Quick Table View for Members who prefer a comprehensive list -->
    <div class="card" style="margin-top:20px;">
        <h4 style="margin:0 0 16px 0; color:#0B1320; font-size:16px;">📋 All Categories Overview</h4>
        <table>
            <thead>
                <tr>
                    <th style="width:50px;">#</th>
                    <th>Category Name</th>
                    <th>Description</th>
                    <th style="text-align:center;">Total Titles</th>
                    <th style="text-align:center;">Available Copies</th>
                    <th style="text-align:center; width:180px;">Action</th>
                </tr>
            </thead>
            <tbody>
                <%
                If catCount > 0 Then
                    For i = 0 To catCount - 1
                        Dim tId, tName, tDesc, tBooks, tAvail
                        tId = tmpList(0, i)
                        tName = tmpList(1, i) & ""
                        tDesc = tmpList(2, i) & ""
                        tBooks = SafeInt(tmpList(3, i), 0)
                        tAvail = SafeInt(tmpList(4, i), 0)
                %>
                    <tr>
                        <td><code>#<%= tId %></code></td>
                        <td>
                            <strong style="color:#0B1320;"><%= GetCategoryIcon(tName) %> <%= CleanText(tName) %></strong>
                        </td>
                        <td><span style="font-size:13.5px; color:#64748b;"><%= CleanText(tDesc) %></span></td>
                        <td style="text-align:center;">
                            <span class="cat-pill cat-pill-books"><%= tBooks %> books</span>
                        </td>
                        <td style="text-align:center;">
                            <% If tAvail > 0 Then %>
                                <span class="cat-pill cat-pill-avail"><%= tAvail %> available</span>
                            <% Else %>
                                <span class="cat-pill cat-pill-out">0 available</span>
                            <% End If %>
                        </td>
                        <td style="text-align:center;">
                            <% If tBooks > 0 Then %>
                                <a href="books.asp?category_id=<%= tId %>" class="btn" style="display:inline-block; padding:6px 14px; background:#0B1320; color:#fff; text-decoration:none; border-radius:6px; font-size:12.5px; font-weight:600;">
                                    Browse &amp; Request &rarr;
                                </a>
                            <% Else %>
                                <span style="font-size:12px; color:#94a3b8;">No books yet</span>
                            <% End If %>
                        </td>
                    </tr>
                <%
                    Next
                End If
                %>
            </tbody>
        </table>
    </div>

    <!-- Client-side Fast Live Filtering Script -->
    <script>
    function filterCategories() {
        var input = document.getElementById("catSearchInput");
        var filter = input.value.toLowerCase().trim();
        var cards = document.querySelectorAll(".category-card");
        
        cards.forEach(function(card) {
            var name = card.getAttribute("data-catname") || "";
            var desc = card.getAttribute("data-catdesc") || "";
            if (name.indexOf(filter) > -1 || desc.indexOf(filter) > -1) {
                card.style.display = "flex";
            } else {
                card.style.display = "none";
            }
        });
    }
    </script>

<% End If %>

<%
conn.Close
Set conn = Nothing
RenderFooter
%>
