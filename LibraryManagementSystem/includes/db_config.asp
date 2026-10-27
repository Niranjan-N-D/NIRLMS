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
' --------------------------------------------------------------------
Function GetConnection()
    Dim conn, connStr, drivers, passwords, d, p, connected
    Set conn = Server.CreateObject("ADODB.Connection")
    connected = False

    drivers = Array("{MySQL ODBC 5.1 Driver}", "{MySQL ODBC 5.1 Unicode Driver}", "{MySQL ODBC 5.1 ANSI Driver}", "{MySQL ODBC 26.7 Unicode Driver}", "{MySQL ODBC 26.7 ANSI Driver}", "{MySQL ODBC 8.0 Unicode Driver}", "{MySQL ODBC 8.0 ANSI Driver}", "{MySQL ODBC 5.3 Unicode Driver}", "{MySQL ODBC 5.3 ANSI Driver}")
    passwords = Array("", "root", "password")

    On Error Resume Next
    For Each d In drivers
        For Each p In passwords
            connStr = "DRIVER=" & d & ";SERVER=localhost;PORT=3306;DATABASE=librarydb;USER=root;PASSWORD=" & p & ";OPTION=3;STMT=SET NAMES utf8mb4;"
            Err.Clear
            conn.Open connStr
            If Err.Number = 0 Then
                connected = True
                Exit For
            End If
        Next
        If connected Then Exit For
    Next

    If Not connected Then
        Response.Write("<div style='color:red;font-family:sans-serif;'>")
        Response.Write("<h3>Database Connection Error</h3>")
        Response.Write("<p>" & Err.Description & "</p>")
        Response.Write("<p>Please verify connection string in <code>includes/db_config.asp</code>.</p>")
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
    Else
        Dim pKey
        pKey = "Perm_" & LCase(permName)
        If IsEmpty(Session(pKey)) Or Session(pKey) = "" Then
            ' Default permissions fallback
            If LCase(permName) = "authors" Or LCase(permName) = "categories" Then
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

' --------------------------------------------------------------------
' Session Authentication Guard
' --------------------------------------------------------------------
Sub RequireAuth()
    If IsEmpty(Session("UserId")) Or Session("UserId") = "" Then
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
    <title><%= pageTitle %> - BEL Library System</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; font-family: 'Plus Jakarta Sans', 'Segoe UI', Arial, 'Noto Sans Kannada', sans-serif; }
        body { margin: 0; padding: 0; background-color: #f4f6fa; color: #1e293b; min-height: 100vh; }
        
        /* Modern Single-Line Aligned Navigation Bar (Reference Match) */
        .navbar {
            background-color: #0b1320;
            color: #fff;
            padding: 0 28px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            height: 60px;
            position: sticky;
            top: 0;
            z-index: 1000;
            box-shadow: 0 2px 12px rgba(0,0,0,0.25);
            white-space: nowrap;
            border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .nav-left {
            display: flex;
            align-items: center;
            gap: 16px;
            flex-shrink: 0;
        }
        .nav-brand {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            text-decoration: none;
            color: #ffffff;
            font-size: 18px;
            font-weight: 800;
            letter-spacing: 0.3px;
            white-space: nowrap;
            margin-right: 8px;
            flex-shrink: 0;
        }
        .nav-brand:hover {
            color: #38bdf8;
        }
        .nav-links {
            display: flex;
            align-items: center;
            gap: 4px;
            flex-shrink: 0;
        }
        .nav-link {
            color: #cbd5e1;
            text-decoration: none;
            padding: 6px 12px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s ease;
            white-space: nowrap;
            flex-shrink: 0;
        }
        .nav-link span {
            font-size: 15px;
            line-height: 1;
        }
        .nav-link:hover {
            color: #ffffff;
            background: rgba(255, 255, 255, 0.12);
        }
        .nav-link.active {
            color: #ffffff !important;
            background: #0077f6 !important;
            font-weight: 700;
            box-shadow: 0 2px 10px rgba(0, 119, 246, 0.45);
        }
        .nav-link.nav-analytics.active {
            background: #0077f6 !important;
        }
        .nav-link.nav-admin.active {
            background: #0077f6 !important;
        }
        .nav-divider {
            height: 20px;
            width: 1px;
            background: rgba(255, 255, 255, 0.18);
            margin: 0 6px;
            flex-shrink: 0;
        }
        .nav-right {
            display: flex;
            align-items: center;
            gap: 14px;
            flex-shrink: 0;
            margin-left: auto;
        }
        .nav-user-info {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13.5px;
            color: #e2e8f0;
            white-space: nowrap;
        }
        .nav-user-info strong {
            font-weight: 700;
            color: #ffffff;
        }
        .nav-logout-btn {
            background: rgba(239, 68, 68, 0.14);
            color: #f87171;
            border: 1px solid rgba(239, 68, 68, 0.35);
            padding: 6px 14px;
            border-radius: 8px;
            font-size: 13.5px;
            font-weight: 600;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
            white-space: nowrap;
        }
        .nav-logout-btn:hover {
            background: #ef4444;
            color: #ffffff;
            border-color: #ef4444;
            box-shadow: 0 2px 10px rgba(239, 68, 68, 0.4);
        }

        @media (max-width: 1380px) {
            .nav-link { padding: 5px 8px; font-size: 13.5px; gap: 4px; }
            .nav-brand { font-size: 17px; }
            .welcome-label { display: none; }
        }
        @media (max-width: 1100px) {
            .nav-links { overflow-x: auto; scrollbar-width: thin; }
        }

        /* Spacious Enterprise Content Container */
        .container { max-width: 1340px; margin: 24px auto; padding: 0 24px; }
        .card { 
            background: #ffffff; 
            padding: 28px 32px; 
            border-radius: 16px; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.04); 
            border: 1px solid #e2e8f0;
            margin-bottom: 24px; 
        }
        .card-header { 
            display: flex; 
            justify-content: space-between; 
            align-items: flex-start; 
            border-bottom: 1px solid #f1f5f9; 
            padding-bottom: 18px; 
            margin-bottom: 22px; 
        }
        .btn { 
            padding: 9px 18px; 
            border-radius: 8px; 
            border: none; 
            font-size: 13.5px; 
            font-weight: 600;
            cursor: pointer; 
            text-decoration: none; 
            display: inline-flex; 
            align-items: center;
            justify-content: center;
            gap: 6px;
            transition: all 0.15s ease;
        }
        .btn:hover { transform: translateY(-1px); }
        .btn-primary { background: #0077f6; color: white; box-shadow: 0 2px 8px rgba(0,119,246,0.3); }
        .btn-success { background: #16a34a; color: white; box-shadow: 0 2px 8px rgba(22,163,74,0.3); }
        .btn-danger { background: #ef4444; color: white; box-shadow: 0 2px 8px rgba(239,68,68,0.3); }
        .btn-secondary { background: #64748b; color: white; }
        .btn-warning { background: #f59e0b; color: white; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 13px 16px; border-bottom: 1px solid #f1f5f9; text-align: left; vertical-align: middle !important; }
        th { background: #f8fafc; color: #475569; font-weight: 700; font-size: 13px; border-bottom: 1.5px solid #e2e8f0; }
        .badge { padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 700; color: white; }
        .bg-success { background: #16a34a; }
        .bg-danger { background: #ef4444; }
        .bg-warning { background: #f59e0b; }

        /* Unified Consistent Status Pills (Soft Pastel Enterprise System) */
        .status-pill {
            display: inline-block;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12.5px;
            font-weight: 700;
            line-height: 1.2;
            text-align: center;
            white-space: nowrap;
        }
        .status-pill-green, .status-pill-available, .status-pill-returned, .status-pill-approved {
            background-color: #dcfce7;
            color: #15803d;
        }
        .status-pill-solid-green {
            background-color: #16a34a;
            color: #ffffff;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .status-pill-amber, .status-pill-pending, .status-pill-borrowed, .status-pill-loaned {
            background-color: #fef3c7;
            color: #b45309;
        }
        .status-pill-blue, .status-pill-collected, .status-pill-issued, .status-pill-circulation, .status-pill-staff {
            background-color: #e0f2fe;
            color: #0369a1;
        }
        .status-pill-red, .status-pill-rejected, .status-pill-expired, .status-pill-overdue, .status-pill-out, .status-pill-admin {
            background-color: #fee2e2;
            color: #b91c1c;
        }
        .status-pill-purple, .status-pill-librarian {
            background-color: #f3e8ff;
            color: #7e22ce;
        }
            color: #9c36b5;
        }
        .status-pill-gray {
            background-color: #e9ecef;
            color: #495057;
        }
        .status-subtext {
            font-size: 13.5px;
            font-weight: 500;
            color: #576574;
            margin-top: 6px;
            letter-spacing: 0.2px;
        }

        /* Unified Vertical Action Button Stack (Matches Design Spec) */
        .action-stack {
            display: flex;
            flex-direction: column;
            gap: 6px;
            width: 120px;
        }
        .action-btn {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 5px;
            width: 100%;
            padding: 7px 12px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 700;
            color: #ffffff !important;
            text-decoration: none;
            border: none;
            cursor: pointer;
            box-sizing: border-box;
            transition: opacity 0.15s ease, transform 0.1s ease;
            text-align: center;
            line-height: 1.2;
        }
        .action-btn:hover {
            opacity: 0.92;
            transform: translateY(-1px);
        }
        .action-btn-review, .btn-act-review {
            background-color: #f39c12;
        }
        .action-btn-edit, .btn-act-edit, .action-btn-primary {
            background-color: #3498db;
        }
        .action-btn-delete, .btn-act-delete, .action-btn-reject {
            background-color: #e74c3c;
        }
        .action-btn-request, .btn-act-request, .action-btn-approve, .action-btn-return {
            background-color: #2ecc71;
        }
        .action-btn-reserve {
            background-color: #e67e22;
        }
        .action-text-muted {
            color: #868e96;
            font-size: 13.5px;
            font-weight: 600;
        }

        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; font-weight: 600; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        .alert { padding: 12px 16px; border-radius: 4px; margin-bottom: 20px; }
        .alert-info { background: #d9edf7; color: #31708f; border: 1px solid #bce8f1; }
        .alert-danger { background: #f2dede; color: #a94442; border: 1px solid #ebccd1; }
    </style>
</head>
<body>
    <div class="navbar">
        <div class="nav-left">
            <a href="dashboard.asp" class="nav-brand">
                <svg width="24" height="16" viewBox="0 0 52 36" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;">
                    <path d="M0 5h34c4.5 0 8 3.5 8 8s-3.5 8-8 8H6" stroke="#0080ff" stroke-width="6.5" stroke-linecap="round"/>
                    <path d="M0 18h38c5 0 9 4 9 9s-4 9-9 9H10" stroke="#0080ff" stroke-width="6.5" stroke-linecap="round"/>
                    <path d="M0 31h32" stroke="#0080ff" stroke-width="6.5" stroke-linecap="round"/>
                </svg>
                <span>BEL Library</span>
            </a>

            <% If Not IsEmpty(Session("UserId")) Then %>
                <div class="nav-links">
                    <a href="dashboard.asp" class="nav-link <%= NavActive("dashboard.asp", currentScript) %>">
                        <span>🏠</span> Dashboard
                    </a>
                    <% If HasPermission("books") Then %>
                        <a href="books.asp" class="nav-link <%= NavActive("books.asp", currentScript) %>">
                            <span>📖</span> <%= IIf(IsStaff(), "Books", "Book Catalog") %>
                        </a>
                    <% End If %>
                    <% If HasPermission("requests") Then %>
                        <a href="requests.asp" class="nav-link <%= NavActive("requests.asp", currentScript) %>">
                            <span>📑</span> <%= IIf(IsStaff(), "Requests", "My Requests") %>
                        </a>
                    <% End If %>
                    <% If HasPermission("borrowings") Then %>
                        <a href="borrowings.asp" class="nav-link <%= NavActive("borrowings.asp", currentScript) %>">
                            <span>⏱️</span> <%= IIf(IsStaff(), "Borrowings", "My Loans") %>
                        </a>
                    <% End If %>
                    <% If IsStaff() Then %>
                        <a href="analytics.asp" class="nav-link nav-analytics <%= NavActive("analytics.asp", currentScript) %>" style="<%= IIf(NavActive("analytics.asp", currentScript)="active", "", "color:#cbd5e1; font-weight:600;") %>">
                            <span>📊</span> Analytics
                        </a>
                        <a href="users.asp" class="nav-link <%= NavActive("users.asp", currentScript) %>">
                            <span>👥</span> Users
                        </a>
                        <span class="nav-divider"></span>
                    <% End If %>

                    <% If HasPermission("authors") Then %>
                        <a href="authors.asp" class="nav-link <%= NavActive("authors.asp", currentScript) %>">
                            <span>✍️</span> Authors
                        </a>
                    <% End If %>
                    <% If HasPermission("categories") Then %>
                        <a href="categories.asp" class="nav-link <%= NavActive("categories.asp", currentScript) %>">
                            <span>🏷️</span> Categories
                        </a>
                    <% End If %>
                    <% If HasPermission("feedback") Then %>
                        <a href="feedback.asp" class="nav-link <%= NavActive("feedback.asp", currentScript) %>">
                            <span>💬</span> Reviews
                        </a>
                    <% End If %>
                    <% If IsAdmin() Then %>
                        <a href="manage_permissions.asp" class="nav-link nav-admin <%= NavActive("manage_permissions.asp", currentScript) %>" style="<%= IIf(NavActive("manage_permissions.asp", currentScript)="active", "", "color:#f39c12; font-weight:600;") %>">
                            <span>⚙️</span> Permissions
                        </a>
                    <% End If %>
                </div>
            <% End If %>
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
    <div class="container">
<%
End Sub

Sub RenderFooter()
%>
    </div>
</body>
</html>
<%
End Sub
%>
