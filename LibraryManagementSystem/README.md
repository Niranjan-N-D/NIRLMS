# Bharat Electronics Limited (BEL) – Central HR Library Management System

A high-performance, **100% offline-ready** enterprise Library Management System developed with **Classic ASP (VBScript)**, **ADODB**, and **MySQL / MariaDB**. Engineered specifically for secure corporate intranet deployment within the **Bharat Electronics Limited (BEL)** internal network.

---

## 🏛️ System Overview

The BEL Central HR Library Management System manages a physical and digital catalog of **1,286 books** across **25 categories** and **508 authors**, featuring complete **bilingual Kannada & English** book titles, accession number tracking, automated 48-hour reservation expiration, and real-time circulation intelligence.

### ⚡ Key Enterprise Highlights
- **100% Corporate Intranet & Air-Gap Ready**: Zero external network or internet dependencies. No public CDNs, no Google Fonts, and no external tracking. Built entirely with vanilla HTML5, modern CSS3, and local JavaScript.
- **Bundled Offline Visualizations**: Interactive Doughnut and Bar circulation charts powered by a locally bundled [Chart.js](file:///c:/Projects/library-main%20(1)/library-main1/LibraryManagementSystem/js/chart.min.js) engine.
- **100% Zoom Responsive Layout**: Table catalog with strict column proportioning, compact 2×2 action buttons, and stacked bilingual subtitles to guarantee fit on standard office displays without horizontal clipping.
- **Instantaneous Page Loads (<100ms)**: Sub-second rendering across all modules through optimized ADODB buffering, indexed database schemas, and elimination of VBScript context switches.
- **Multi-Driver ODBC Resilience**: Automatic driver negotiation supporting `{MySQL ODBC 5.1 Driver}`, `{MySQL ODBC 8.0 Unicode Driver}`, `{MySQL ODBC 8.0 ANSI Driver}`, and `{MySQL ODBC 26.7 Unicode Driver}`.

---

## 🌟 Core Modules & Capabilities

### 1. 🔐 Role-Based Access Control (RBAC) & Authentication
- Authentication is governed by **Staff Number / Username** and **Date of Birth (DOB)**:
  - Formats accepted: `DD-MM-YYYY`, `DDMMYYYY`, or standard password hash.
- **Admin**: Full system management, user role assignments, catalog administration, circulation audits, and permission controls.
- **Librarian**: Catalog stock operations, book issuing/returns, book request approvals, and category management.
- **Member (Staff)**: Personal loan status, custom 48-hour book reservations, review submissions, and history tracking.

### 2. 📚 Bilingual Catalog Management (`books.asp`)
- Search by **Title (Kannada/English)**, **Accession Code / ISBN**, **Author**, or **Category**.
- Dynamic dropdown category filter with instant 50-row pagination.
- Default ordering by Book ID ascending (`#1` to `#1286`).

### 3. ⚡ Operational Transactions (`borrowings.asp` & `requests.asp`)
- **Book Issue & Return**: Fast searchable combobox for staff members and available shelf stock.
- **Automated Overdue Auditing**: Throttled background updates flagging overdue loans past due date.
- **48-Hour Book Reservation Pipeline**: Members request books; staff approves for physical pickup. If uncollected after 48 hours, the system automatically expires the request and restores stock inventory.

### 4. 📈 Circulation Analytics & Reports (`analytics.asp`)
- Real-time stock breakdown (Available, Active Loans, Overdue, Returned).
- Category-wise inventory distribution and borrowing velocity.
- Audit history logging.

### 5. 💬 Bilingual Feedback & Reviews (`feedback.asp`)
- Star rating and text review system for catalog books or general library facilities.
- Client and server-side **150-word limit validation**.

---

## 🔐 Default Credentials

| Role | Name | Staff Number / Login | Password / DOB | Permissions |
| :--- | :--- | :--- | :--- | :--- |
| **Admin** | Admin User | `9876543210` or `BEL-ADM-001` | `15-08-1985` | Full System Control, Permissions, Audits |
| **Librarian** | Librarian Sarah | `9876543211` | `20-05-1992` | Book Add/Edit, Issue, Return, Request Approval |
| **Staff Member** | Rajesh Kumar | `9876543212` | `10-10-1995` | Browse Catalog, Request Books, Write Reviews |
| **Staff Member** | Priya Sharma | `9876543213` | `12-12-1996` | Browse Catalog, Request Books, Write Reviews |

*(Master seed script also creates individual user accounts including `Niranjan` as Admin).*

---

## 📁 Repository Directory Structure

```text
library-main1/
├── LibraryManagementSystem/
│   ├── images/                 # BEL official emblems, logos & favicons
│   │   ├── bel_emblem_white.png
│   │   ├── bel_logo.png
│   │   └── bel_favicon_32.png
│   ├── includes/
│   │   └── db_config.asp       # Database connection, pooling, session auth & navigation
│   ├── js/
│   │   └── chart.min.js        # Bundled offline Chart.js v4.4.1 library
│   ├── analytics.asp           # Circulation intelligence, audits & charts
│   ├── authors.asp             # Authors directory master
│   ├── books.asp               # Catalog, bilingual display & pagination
│   ├── books_add.asp           # Add catalog book
│   ├── books_delete.asp        # Soft delete book (preserves audit integrity)
│   ├── books_edit.asp          # Edit book metadata & stock
│   ├── borrowings.asp          # Book issue, return & active loan log
│   ├── categories.asp          # Categories master
│   ├── dashboard.asp           # Role-aware executive dashboard
│   ├── default.asp             # Root redirect to login
│   ├── feedback.asp            # 150-word reviews & facility ratings
│   ├── index.asp               # IIS index redirect
│   ├── login.asp               # Secure Staff Number & DOB sign-in
│   ├── logout.asp              # Session termination
│   ├── manage_permissions.asp  # Dynamic role permission matrix
│   ├── request_book.asp        # Member reservation form
│   ├── requests.asp            # Request collection & approval pipeline
│   ├── schema_mysql.sql        # Complete master database schema (1,286 books)
│   ├── styles.css              # Global responsive stylesheet (system fonts)
│   └── users.asp               # User & staff accounts management
├── Switch_To_Workspace.bat     # Windows batch script to configure IIS virtual site
├── apply_update.cmd            # Automated deployment sync script to IIS
└── README.md                   # System documentation
```

---

## 🚀 Office Server Deployment Guide

### Prerequisites
1. **Windows Server with IIS**:
   - Enable **World Wide Web Services ➔ Application Development Features ➔ ASP** (Classic ASP).
2. **Database**: MySQL 5.7+ or MariaDB 10.3+ on `127.0.0.1:3306`.
3. **ODBC Driver**: Install either `MySQL ODBC 5.1 Driver` or `MySQL ODBC 8.0 Driver`.

### Step 1: Database Initialization
1. Open **SQLyog** (or MySQL CLI / Workbench).
2. Execute the self-contained [schema_mysql.sql](file:///c:/Projects/library-main%20(1)/library-main1/LibraryManagementSystem/schema_mysql.sql).
3. This creates database `librarydb1` (with `utf8_unicode_ci` collation for Kannada text), tables, triggers, indexes, seed accounts, 25 categories, 508 authors, and all **1,286 catalog books**.

### Step 2: IIS Configuration
1. Open **IIS Manager** (`inetmgr`).
2. Add a Website or Virtual Application:
   - **Physical Path**: Point to the `LibraryManagementSystem` folder.
   - **Application Pool**: Ensure 32-bit applications are enabled if using the 32-bit MySQL ODBC 5.1 Driver (`Enable 32-Bit Applications = True`).
3. Under **ASP** features in IIS Manager:
   - Set **Enable Parent Paths** = `True`.
   - Set **Send Errors to Browser** = `False` (for production security).

### Step 3: Access the Portal
Open any intranet browser and navigate to:
```text
http://<server-ip-or-hostname>/login.asp
```
*(e.g., `http://127.0.0.1:8040/login.asp` or `http://localhost:8040/`)*

---

## 🛡️ License & Institutional Attribution
Built exclusively for the **Central Human Resources Department**, **Bharat Electronics Limited (BEL)**, a Government of India Navratna enterprise under the Ministry of Defence.