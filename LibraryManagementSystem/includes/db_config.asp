<%@ Language=VBScript CodePage=65001 %>
<%
' ====================================================================
' Database Configuration & Shared Helpers for Classic ASP (.asp)
' UTF-8 / Kannada Script Support Enabled
' ====================================================================
Option Explicit

' Enable Response Buffering & UTF-8 Session/Output Encoding
Response.Buffer = True
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Session.CodePage = 65001

' Safe UTF-8 HTML Text Escaping (preserves Unicode / Kannada scripts)
Function CleanText(val)
    If IsNull(val) Or val = "" Then
        CleanText = ""
    Else
        Dim s
        s = CStr(val)
        s = Replace(s, "&", "&amp;")
        s = Replace(s, "<", "&lt;")
        s = Replace(s, ">", "&gt;")
        s = Replace(s, """", "&quot;")
        CleanText = s
    End If
End Function

' --------------------------------------------------------------------
' Database Connection Function
' Single verified driver string with ODBC Connection Pooling enabled
' Optimized for MySQL ODBC 5.1 Driver as specified for deployment
' --------------------------------------------------------------------
Function GetConnection()
    Dim conn, connStr, dbDrivers, dbNames, dbServer, dbPort, dbUser, dbPass
    Dim dIdx, nIdx, isConnected
    Set conn = Server.CreateObject("ADODB.Connection")

    conn.ConnectionTimeout = 5
    conn.CommandTimeout = 30

    dbServer = "127.0.0.1"
    dbPort   = "3306"
    dbUser   = "root"
    dbPass   = ""

    ' Drivers in order of priority (32-bit & 64-bit IIS support on Windows Server)
    dbDrivers = Array("{MySQL ODBC 5.1 Driver}", "{MySQL ODBC 8.0 Unicode Driver}", "{MySQL ODBC 8.0 ANSI Driver}", "{MySQL ODBC 26.7 Unicode Driver}", "{MariaDB ODBC 3.1 Driver}", "{MariaDB ODBC 3.0 Driver}")
    dbNames   = Array("librarydb1", "librarydb")

    isConnected = False
    On Error Resume Next
    For dIdx = 0 To UBound(dbDrivers)
        For nIdx = 0 To UBound(dbNames)
            Err.Clear
            If conn.State = 1 Then conn.Close
            connStr = "DRIVER=" & dbDrivers(dIdx) & ";SERVER=" & dbServer & ";PORT=" & dbPort & ";DATABASE=" & dbNames(nIdx) & ";USER=" & dbUser & ";PASSWORD=" & dbPass & ";OPTION=3;STMT=SET NAMES utf8;"
            conn.Open connStr
            If Err.Number = 0 And conn.State = 1 Then
                isConnected = True
                Exit For
            End If
        Next
        If isConnected Then Exit For
    Next

    If Not isConnected Or Err.Number <> 0 Then
        Response.Write("<div style='color:red;font-family:sans-serif;'>")
        Response.Write("<h3>Database Connection Error</h3>")
        Response.Write("<p>" & Err.Description & "</p>")
        Response.Write("<p>Server: <code>" & dbServer & ":" & dbPort & "</code></p>")
        Response.Write("<p>Please verify MySQL service is running and ODBC driver is registered.</p>")
        Response.Write("</div>")
        Response.End
    End If
    On Error GoTo 0

    Set GetConnection = conn
End Function

' --------------------------------------------------------------------
' Role & Permission Helper Functions
' --------------------------------------------------------------------
Function IsAdmin()
    IsAdmin = (LCase(Trim(CStr(Session("Role") & ""))) = "admin")
End Function

Function IsLibrarian()
    IsLibrarian = (LCase(Trim(CStr(Session("Role") & ""))) = "librarian")
End Function

Function IsStaff()
    IsStaff = (IsAdmin() Or IsLibrarian())
End Function

Function IsMember()
    IsMember = (Not IsStaff())
End Function

Function HasPermission(permName)
    If IsAdmin() Then
        HasPermission = True
    ElseIf LCase(permName) = "categories" Then
        ' Categories catalog browsing is accessible to all users (members & staff)
        HasPermission = True
    Else
        Dim pKey
        pKey = "Perm_" & LCase(permName)
        If IsEmpty(Session(pKey)) Or Session(pKey) = "" Then
            ' Default permissions fallback
            If LCase(permName) = "authors" Then
                HasPermission = IsStaff()
            Else
                HasPermission = True
            End If
        Else
            HasPermission = (CBool(Session(pKey)) = True)
        End If
    End If
End Function

Function IIf(expr, trueVal, falseVal)
    If expr Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

Function IsAuthenticated()
    IsAuthenticated = (Not IsEmpty(Session("UserId")) And Session("UserId") <> "")
End Function

' --------------------------------------------------------------------
' Session Authentication Guard
' --------------------------------------------------------------------
Sub RequireAuth()
    If Not IsAuthenticated() Then
        Response.Redirect("login.asp?msg=Please+login+to+continue")
        Response.End
    End If
End Sub

' --------------------------------------------------------------------
' SQL Sanitization Helper (Prevents SQL Injection in VBScript)
' --------------------------------------------------------------------
Function SQLQuote(value)
    If IsNull(value) Or IsEmpty(value) Then
        SQLQuote = "NULL"
    Else
        Dim sanitized
        sanitized = Replace(CStr(value), "'", "''")
        SQLQuote = "'" & sanitized & "'"
    End If
End Function

Function FormatSQLDate(dtVal)
    If IsNull(dtVal) Or IsEmpty(dtVal) Or dtVal = "" Then
        FormatSQLDate = "NULL"
    ElseIf IsDate(dtVal) Then
        Dim d
        d = CDate(dtVal)
        FormatSQLDate = "'" & Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "'"
    Else
        FormatSQLDate = SQLQuote(dtVal)
    End If
End Function

Function SafeInt(value, defaultVal)
    If IsNull(value) Or IsEmpty(value) Then
        SafeInt = defaultVal
    Else
        On Error Resume Next
        Dim sVal
        sVal = CStr(value & "")
        If Err.Number = 0 And IsNumeric(sVal) Then
            SafeInt = CLng(sVal)
        Else
            SafeInt = defaultVal
        End If
        On Error GoTo 0
    End If
End Function

Function SafeFloat(value, defaultVal)
    If IsNull(value) Or IsEmpty(value) Then
        SafeFloat = defaultVal
    Else
        On Error Resume Next
        Dim sVal
        sVal = CStr(value & "")
        If Err.Number = 0 And IsNumeric(sVal) Then
            SafeFloat = CDbl(sVal)
        Else
            SafeFloat = defaultVal
        End If
        On Error GoTo 0
    End If
End Function

' --------------------------------------------------------------------
' Navigation Active Link Helper
' --------------------------------------------------------------------
Function NavActive(targetPage, scriptName)
    If LCase(Trim(scriptName & "")) = LCase(Trim(targetPage & "")) Then
        NavActive = "active"
    Else
        NavActive = ""
    End If
End Function

' --------------------------------------------------------------------
' Navigation Header Component
' --------------------------------------------------------------------
Sub RenderHeader(pageTitle)
    Dim user, role, roleBadgeClass, roleBadgeBg, currentScript, sPath, slashPos
    user = Session("Username")
    role = Session("Role")
    If IsNull(user) Or user = "" Then user = "Guest"
    If IsNull(role) Or role = "" Then role = "Visitor"

    If IsAdmin() Then
        roleBadgeBg = "#e74c3c" ' Red for Admin
    ElseIf IsLibrarian() Then
        roleBadgeBg = "#9b59b6" ' Purple for Librarian
    Else
        roleBadgeBg = "#2ecc71" ' Green for Member
    End If

    sPath = Request.ServerVariables("SCRIPT_NAME") & ""
    slashPos = InStrRev(sPath, "/")
    If slashPos > 0 Then
        currentScript = LCase(Mid(sPath, slashPos + 1))
    Else
        currentScript = LCase(sPath)
    End If
%>
<!DOCTYPE html>
<html lang="kn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><%= pageTitle %> – BEL</title>
    <link rel="icon" type="image/png" sizes="32x32" href="images/bel_favicon_32.png">
    <link rel="icon" type="image/png" sizes="64x64" href="images/bel_favicon.png">
    <link rel="shortcut icon" href="images/bel_favicon.png">
    <link rel="stylesheet" href="styles.css">

</head>
<body>
    <div class="navbar">
        <div class="nav-left">
            <button type="button" class="sidebar-toggle-btn" id="sidebarToggleBtn" onclick="toggleSidebar()" title="Toggle Sidebar (Expand/Collapse)">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                    <line x1="3" y1="12" x2="21" y2="12"/>
                    <line x1="3" y1="6" x2="21" y2="6"/>
                    <line x1="3" y1="18" x2="21" y2="18"/>
                </svg>
            </button>

            <a href="dashboard.asp" class="nav-brand">
                <img src="images/bel_emblem_white.png" alt="BEL Emblem" style="height:22px; width:auto; vertical-align:middle; display:inline-block; object-fit:contain;">
                <span>Central HR Library</span>
            </a>
            <span style="color:#64748b; font-size:13px; font-weight:600;" class="welcome-label">&bull; Information &amp; Resource Centre</span>
        </div>

        <div class="nav-right">
            <% If Not IsEmpty(Session("UserId")) Then %>
                <div class="nav-user-info">
                    <span>Welcome, <strong><%= CleanText(user) %></strong> <span style="font-size:10px; opacity:0.75; vertical-align:middle;">▼</span></span>
                    <span class="badge" style="background:<%= roleBadgeBg %>; border-radius:12px; font-size:12px; padding:3px 10px; font-weight:700;"><%= CleanText(role) %></span>
                </div>
                <a href="logout.asp" class="nav-logout-btn" title="Sign out of your account">
                    <span>🚪</span> Logout
                </a>
            <% Else %>
                <a href="login.asp" class="btn btn-primary" style="padding:6px 16px; font-size:14px; font-weight:600;">Login</a>
            <% End If %>
        </div>
    </div>

    <!-- App Layout: Left Sidebar + Main Content -->
    <div class="app-layout" id="appLayout">
        <% If Not IsEmpty(Session("UserId")) Then %>
            <aside class="app-sidebar" id="appSidebar">
                <!-- 1. MASTER SECTION -->
                <div class="sidebar-group">
                    <div class="sidebar-heading" title="Master Data Entities">
                        <span class="heading-badge">📁</span>
                        <span class="heading-text">MASTER</span>
                    </div>
                    <% If HasPermission("books") Then %>
                        <a href="books.asp" class="side-link <%= NavActive("books.asp", currentScript) %>" title="Books Catalog">
                            <span class="side-icon">📖</span>
                            <span class="side-text"><%= IIf(IsStaff(), "Books Catalog", "Book Catalog") %></span>
                        </a>
                    <% End If %>
                    <% If IsStaff() Then %>
                        <a href="users.asp" class="side-link <%= NavActive("users.asp", currentScript) %>" title="Staff & User Directory">
                            <span class="side-icon">👥</span>
                            <span class="side-text">Staff &amp; Users</span>
                        </a>
                    <% End If %>
                    <% If HasPermission("authors") Then %>
                        <a href="authors.asp" class="side-link <%= NavActive("authors.asp", currentScript) %>" title="Authors Master">
                            <span class="side-icon">✍️</span>
                            <span class="side-text">Authors</span>
                        </a>
                    <% End If %>
                    <% If HasPermission("categories") Then %>
                        <a href="categories.asp" class="side-link <%= NavActive("categories.asp", currentScript) %>" title="Categories Master">
                            <span class="side-icon">🏷️</span>
                            <span class="side-text">Categories</span>
                        </a>
                    <% End If %>
                </div>

                <!-- 2. TRANSACTIONS SECTION -->
                <div class="sidebar-group">
                    <div class="sidebar-heading" title="Operational Transactions">
                        <span class="heading-badge">⚡</span>
                        <span class="heading-text">TRANSACTIONS</span>
                    </div>
                    <% If HasPermission("requests") Then %>
                        <a href="requests.asp" class="side-link <%= NavActive("requests.asp", currentScript) %>" title="Book Requests Approval & Collection">
                            <span class="side-icon">📑</span>
                            <span class="side-text"><%= IIf(IsStaff(), "Book Requests (48h)", "My Requests") %></span>
                        </a>
                    <% End If %>
                    <% If HasPermission("borrowings") Then %>
                        <a href="borrowings.asp" class="side-link <%= NavActive("borrowings.asp", currentScript) %>" title="Borrowings, Loans & Returns">
                            <span class="side-icon">⏱️</span>
                            <span class="side-text"><%= IIf(IsStaff(), "Issue / Return Log", "My Active Loans") %></span>
                        </a>
                    <% End If %>
                    <% If HasPermission("feedback") Then %>
                        <a href="feedback.asp" class="side-link <%= NavActive("feedback.asp", currentScript) %>" title="Book Reviews & Website Ratings">
                            <span class="side-icon">💬</span>
                            <span class="side-text">Reviews &amp; Ratings</span>
                        </a>
                    <% End If %>
                </div>

                <!-- 3. ANALYTICS & REPORTS SECTION -->
                <div class="sidebar-group">
                    <div class="sidebar-heading" title="Analytics, Reporting & Administration">
                        <span class="heading-badge">📊</span>
                        <span class="heading-text">ANALYTICS &amp; SYSTEM</span>
                    </div>
                    <a href="dashboard.asp" class="side-link <%= NavActive("dashboard.asp", currentScript) %>" title="Operational Dashboard">
                        <span class="side-icon">🏠</span>
                        <span class="side-text">Dashboard</span>
                    </a>
                    <% If IsStaff() Then %>
                        <a href="analytics.asp" class="side-link <%= NavActive("analytics.asp", currentScript) %>" title="Circulation Intelligence & Audits">
                            <span class="side-icon">📈</span>
                            <span class="side-text">Circulation Analytics</span>
                        </a>
                    <% End If %>
                    <% If IsAdmin() Then %>
                        <a href="manage_permissions.asp" class="side-link <%= NavActive("manage_permissions.asp", currentScript) %>" title="Role Feature Permissions">
                            <span class="side-icon">⚙️</span>
                            <span class="side-text">Role Permissions</span>
                        </a>
                    <% End If %>
                </div>
            </aside>
        <% End If %>

        <main class="app-main-content">
            <div class="container">
<%
End Sub

Sub RenderFooter()
%>
            </div>
        </main>
    </div>

    <script>
    function toggleSidebar() {
        var layout = document.getElementById("appLayout");
        if (!layout) return;
        layout.classList.toggle("sidebar-collapsed");
        try {
            var isCollapsed = layout.classList.contains("sidebar-collapsed");
            localStorage.setItem("bel_sidebar_collapsed", isCollapsed ? "true" : "false");
        } catch(e) {}
    }
    (function() {
        try {
            if (localStorage.getItem("bel_sidebar_collapsed") === "true") {
                var layout = document.getElementById("appLayout");
                if (layout) layout.classList.add("sidebar-collapsed");
            }
        } catch(e) {}
    })();
    </script>
</body>
</html>
<%
End Sub
%>
