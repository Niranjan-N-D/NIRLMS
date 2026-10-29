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
    <title><%= pageTitle %> - Central HR Library</title>
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

        /* Option A: Left Sidebar Layout & Navigation Styles */
        .app-layout {
            display: flex;
            min-height: calc(100vh - 60px);
            position: relative;
            background: #f4f6fa;
        }
        .app-sidebar {
            position: sticky;
            top: 60px;
            height: calc(100vh - 60px);
            width: 250px;
            background: #0b1320;
            border-right: 1px solid rgba(255, 255, 255, 0.08);
            color: #cbd5e1;
            padding: 16px 12px 28px 12px;
            display: flex;
            flex-direction: column;
            gap: 16px;
            flex-shrink: 0;
            transition: width 0.22s cubic-bezier(0.4, 0, 0.2, 1), padding 0.22s ease;
            overflow-y: auto;
            overflow-x: hidden;
            user-select: none;
            z-index: 900;
        }
        .app-sidebar::-webkit-scrollbar {
            width: 4px;
        }
        .app-sidebar::-webkit-scrollbar-track {
            background: #0b1320;
        }
        .app-sidebar::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.15);
            border-radius: 4px;
        }
        .sidebar-group {
            display: flex;
            flex-direction: column;
            gap: 3px;
        }
        .sidebar-heading {
            font-size: 11px;
            font-weight: 800;
            color: #64748b;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            padding: 6px 12px 6px 12px;
            display: flex;
            align-items: center;
            gap: 8px;
            white-space: nowrap;
        }
        .heading-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
        }
        .side-link {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 12px;
            border-radius: 8px;
            color: #94a3b8;
            text-decoration: none;
            font-size: 13.5px;
            font-weight: 600;
            transition: all 0.15s ease;
            white-space: nowrap;
        }
        .side-link:hover {
            background: rgba(255, 255, 255, 0.08);
            color: #ffffff;
        }
        .side-link.active {
            background: #0077f6 !important;
            color: #ffffff !important;
            font-weight: 700;
            box-shadow: 0 2px 8px rgba(0, 119, 246, 0.35);
        }
        .side-icon {
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            flex-shrink: 0;
        }
        .side-text {
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        /* Collapsed Sidebar State (68px wide) */
        .sidebar-collapsed .app-sidebar {
            width: 68px;
            padding: 16px 8px 28px 8px;
        }
        .sidebar-collapsed .sidebar-heading .heading-text {
            display: none;
        }
        .sidebar-collapsed .sidebar-heading {
            justify-content: center;
            padding: 6px 0;
        }
        .sidebar-collapsed .side-link {
            justify-content: center;
            padding: 9px 0;
        }
        .sidebar-collapsed .side-text {
            display: none;
        }

        /* Hamburger Toggle Button */
        .sidebar-toggle-btn {
            background: rgba(255, 255, 255, 0.1);
            color: #ffffff;
            border: 1px solid rgba(255, 255, 255, 0.16);
            width: 34px;
            height: 34px;
            border-radius: 7px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: background 0.15s ease;
            padding: 0;
            margin-right: 2px;
            flex-shrink: 0;
        }
        .sidebar-toggle-btn:hover {
            background: rgba(255, 255, 255, 0.2);
        }

        /* Main Content Container */
        .app-main-content {
            flex: 1;
            padding: 24px 30px;
            background: #f4f6fa;
            overflow-x: auto;
            min-width: 0;
        }
        .app-main-content .container {
            max-width: 100%;
            margin: 0;
            padding: 0;
        }

        @media (max-width: 860px) {
            .app-sidebar {
                position: absolute;
                top: 0;
                bottom: 0;
                left: 0;
                box-shadow: 4px 0 20px rgba(0,0,0,0.3);
            }
            .sidebar-collapsed .app-sidebar {
                transform: translateX(-100%);
                width: 250px;
            }
            .app-main-content {
                padding: 16px 14px;
            }
        }
    </style>
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
