# BEL Library Management System - Classic ASP (.asp) & MySQL

A complete, feature-rich **Classic ASP (`.asp`)** web application for library management built with **VBScript**, **ADODB**, and **MySQL (SQLyog)**. 

The system features **3-Role Access Control (RBAC)**, custom **Book Requests with Start & End Dates**, dynamic dashboards, and full catalog management.

--- 

## 🌟 Key Features

* **3-Role Access Control (RBAC)**:
  * 🔴 **Admin**: Full control over system overview metrics, books (add/edit/delete), member book request approvals, authors, and categories.
  * 🟣 **Librarian**: Operational management of system overview metrics, books (add/edit), member book request approvals, issue/return transactions, authors, and categories.
  * 🟢 **Member**: Personal dashboard, book catalog search, custom book requesting, and personal borrowing history tracking.
* **Book Request & Approval Pipeline**:
  * Members select books and submit requests specifying custom **Start Date** (Borrow Date) and **End Date** (Due Date).
  * Staff (Admin/Librarian) review pending requests on `requests.asp` to **Approve** or **Reject** them.
  * Approving a request automatically creates an active borrowing transaction and updates available book copies.
* **Dual-View Dashboards**:
  * Staff sees system-wide counters (Total Books, Pending Requests, Active Borrowings, Total Members).
  * Members see personal metrics (Active Borrowings, Pending Requests, Available Books).

---

## 🔐 Default Credentials (Login with Staff No. & Date of Birth)

| Role | Name | Staff Number | Password (DOB) | Access Rights |
| :--- | :--- | :--- | :--- | :--- |
| **Admin** | Admin User | `9876543210` | `15-08-1985` | Full System Control (Metrics, Books Add/Edit/Delete, Requests, Authors, Categories, Permissions) |
| **Librarian** | Librarian Sarah | `9876543211` | `20-05-1992` | Operational Staff (Metrics, Books Add/Edit, Request Approvals/Rejections, Authors, Categories) |
| **Staff** | Rajesh Kumar | `9876543212` | `10-10-1995` | Staff View (Personal Dashboard, Book Requesting, Personal History, Reviews) |
| **Staff** | Priya Sharma | `9876543213` | `12-12-1996` | Staff View (Personal Dashboard, Book Requesting, Personal History, Reviews) |

---

## 🛠️ Project Structure

```
LibraryManagementSystem/
├── includes/
│   └── db_config.asp       # MySQL ODBC connection string, VBScript helpers & Navbar Header/Footer
├── authors.asp             # Authors management (Staff access: Add, List, Delete)
├── books.asp               # Book catalog & live search (Staff: Add/Edit/Delete; Member: Request Book)
├── books_add.asp           # Add new book record (Staff only)
├── books_edit.asp          # Edit book details (Staff only)
├── books_delete.asp        # Delete book record (Admin only)
├── borrowings.asp          # Issue books & process return transactions (Staff view & Member history)
├── categories.asp          # Book categories management (Staff access)
├── dashboard.asp           # Role-aware metrics dashboard (System vs. Member view)
├── login.asp               # User sign-in with password verification & credential references
├── logout.asp              # Session sign-out handler
├── request_book.asp        # Member book request form (Start Date & End Date selection)
├── requests.asp           # Book requests approval/rejection pipeline (Staff) & Request status ledger (Member)
└── schema_mysql.sql        # MySQL database schema & seed data (SQLyog compatible)
```

---

## 🚀 Setup & Installation Guide

### Step 1: Database Setup in SQLyog / MySQL
1. Open **SQLyog** (or MySQL Workbench / Command Line) and connect to your MySQL Server instance (`localhost:3306`).
2. Open `schema_mysql.sql` and execute all statements.
3. This creates the `librarydb` database and seeds users, authors, categories, and books.

### Step 2: Configure IIS Server
1. Press `Win + R`, type `optionalfeatures`, and press **Enter**.
2. Under **Internet Information Services ➔ World Wide Web Services ➔ Application Development Features**, enable **ASP** (Classic ASP).
3. Ensure **MySQL ODBC Driver** (e.g. `MySQL ODBC 5.1 Driver`, `MySQL ODBC 8.0 Driver`, or `MySQL ODBC 26.7 Driver`) is installed.
4. Open **IIS Manager** (`inetmgr`):
   * Create a Virtual Directory / Application pointing to `C:\Projects\library\LibraryManagementSystem`.
   * Open **ASP** settings in IIS Manager and set **Enable Parent Paths** to `True`.

### Step 3: Access the Application
Open browser and navigate to:
👉 **`http://localhost/LibraryManagementSystem/login.asp`**

---

## 🌐 Accessing from Other PCs on Local Wi-Fi / LAN (Zero Setup!)

If the app is running on your PC, other devices on the same Wi-Fi network can access it without installing anything:

1. Open Command Prompt on your PC, type `ipconfig`, and find your **IPv4 Address** (e.g. `192.168.1.15`).
2. On any other PC, laptop, tablet, or mobile phone connected to the same Wi-Fi, open the browser and navigate to:
   ```text
   http://192.168.1.15/LibraryManagementSystem/login.asp
   ```
