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
    <style>
        * { box-sizing: border-box; font-family: 'Segoe UI', Arial, 'Noto Sans Kannada', Tahoma, Geneva, Verdana, sans-serif; }
        body { margin: 0; padding: 0; background-color: #f4f6f9; color: #333; }
        
        /* Modern Single-Line Aligned Navigation Bar */
        .navbar {
            background-color: #16202c;
            color: #fff;
            padding: 0 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            height: 64px;
            position: sticky;
            top: 0;
            z-index: 1000;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            white-space: nowrap;
        }
        .nav-left {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-shrink: 0;
        }
        .nav-brand {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            text-decoration: none;
            color: #ffffff;
            font-size: 19px;
            font-weight: 800;
            letter-spacing: 0.3px;
            white-space: nowrap;
            margin-right: 6px;
            flex-shrink: 0;
        }
        .nav-brand:hover {
            color: #3498db;
        }
        .nav-links {
            display: flex;
            align-items: center;
            gap: 3px;
            flex-shrink: 0;
        }
        .nav-link {
            color: #d1d8e0;
            text-decoration: none;
            padding: 6px 10px;
            border-radius: 6px;
            font-size: 15px;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            transition: all 0.18s ease;
            white-space: nowrap;
            flex-shrink: 0;
        }
        .nav-link span {
            font-size: 16px;
            line-height: 1;
        }
        .nav-link:hover {
            color: #ffffff;
            background: rgba(255, 255, 255, 0.15);
        }
        .nav-link.active {
            color: #ffffff !important;
            background: #3498db !important;
            font-weight: 700;
            box-shadow: 0 2px 8px rgba(52, 152, 219, 0.45);
        }
        .nav-link.nav-analytics.active {
            background: #27ae60 !important;
            box-shadow: 0 2px 8px rgba(39, 174, 96, 0.45);
        }
        .nav-link.nav-admin.active {
            background: #e67e22 !important;
            box-shadow: 0 2px 8px rgba(230, 126, 34, 0.45);
        }
        .nav-divider {
            height: 22px;
            width: 1px;
            background: rgba(255, 255, 255, 0.2);
            margin: 0 4px;
            flex-shrink: 0;
        }
        .nav-right {
            display: flex;
            align-items: center;
            gap: 12px;
            flex-shrink: 0;
            margin-left: auto;
        }
        .nav-user-info {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14.5px;
            color: #ecf0f1;
            white-space: nowrap;
        }
        .nav-user-info strong {
            font-weight: 700;
            color: #ffffff;
        }
        .nav-logout-btn {
            background: rgba(231, 76, 60, 0.18);
            color: #ff7675;
            border: 1px solid rgba(231, 76, 60, 0.4);
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
            white-space: nowrap;
        }
        .nav-logout-btn span {
            font-size: 15px;
            line-height: 1;
        }
        .nav-logout-btn:hover {
            background: #e74c3c;
            color: #ffffff;
            border-color: #e74c3c;
            box-shadow: 0 2px 8px rgba(231, 76, 60, 0.35);
        }

        @media (max-width: 1440px) {
            .nav-link { padding: 5px 8px; font-size: 14px; gap: 4px; }
            .nav-link span { font-size: 15px; }
            .nav-brand { font-size: 18px; margin-right: 4px; }
            .nav-user-info { font-size: 13.5px; }
            .welcome-label { display: none; }
        }
        @media (max-width: 1250px) {
            .nav-links { overflow-x: auto; scrollbar-width: thin; }
        }

        .container { max-width: 1100px; margin: 30px auto; padding: 0 20px; }
        .card { background: white; padding: 24px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 24px; }
        .card-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #f1f1f1; padding-bottom: 12px; margin-bottom: 20px; }
        .btn { padding: 8px 16px; border-radius: 4px; border: none; font-size: 14px; cursor: pointer; text-decoration: none; display: inline-block; }
        .btn-primary { background: #3498db; color: white; }
        .btn-success { background: #2ecc71; color: white; }
        .btn-danger { background: #e74c3c; color: white; }
        .btn-secondary { background: #95a5a6; color: white; }
        .btn-warning { background: #f39c12; color: white; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; border-bottom: 1px solid #eee; text-align: left; }
        th { background: #f8f9fa; color: #2c3e50; font-weight: 600; }
        .badge { padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold; color: white; }
        .bg-success { background: #2ecc71; }
        .bg-danger { background: #e74c3c; }
        .bg-warning { background: #f39c12; }
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
                <span style="font-size:22px;">📖</span>
                <span>BEL Library</span>
            </a>

            <% If Not IsEmpty(Session("UserId")) Then %>
                <div class="nav-links">
                    <a href="dashboard.asp" class="nav-link <%= NavActive("dashboard.asp", currentScript) %>">
                        <span>🏠</span> Dashboard
                    </a>
                    <% If HasPermission("books") Then %>
                        <a href="books.asp" class="nav-link <%= NavActive("books.asp", currentScript) %>">
                            <span>📚</span> <%= IIf(IsStaff(), "Books", "Book Catalog") %>
                        </a>
                    <% End If %>
                    <% If HasPermission("requests") Then %>
                        <a href="requests.asp" class="nav-link <%= NavActive("requests.asp", currentScript) %>">
                            <span>📋</span> <%= IIf(IsStaff(), "Requests", "My Requests") %>
                        </a>
                    <% End If %>
                    <% If HasPermission("borrowings") Then %>
                        <a href="borrowings.asp" class="nav-link <%= NavActive("borrowings.asp", currentScript) %>">
                            <span>⏱️</span> <%= IIf(IsStaff(), "Borrowings", "My Loans") %>
                        </a>
                    <% End If %>
                    <% If IsStaff() Then %>
                        <a href="analytics.asp" class="nav-link nav-analytics <%= NavActive("analytics.asp", currentScript) %>" style="<%= IIf(NavActive("analytics.asp", currentScript)="active", "", "color:#2ecc71; font-weight:600;") %>">
                            <span>📊</span> Analytics
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
                    <span><span class="welcome-label">Welcome, </span><strong><%= CleanText(user) %></strong></span>
                    <span class="badge" style="background:<%= roleBadgeBg %>; font-size:12px; padding:3px 9px; font-weight:700;"><%= CleanText(role) %></span>
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
