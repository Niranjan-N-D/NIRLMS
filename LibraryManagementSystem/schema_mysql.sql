-- ====================================================================
-- BEL Library Management System - MySQL Database Schema (SQLyog Compatible)
-- UTF-8 / Kannada Unicode Supported (utf8)
-- ====================================================================

DROP DATABASE IF EXISTS librarydb1;
CREATE DATABASE librarydb1 DEFAULT CHARACTER SET utf8 COLLATE utf8_unicode_ci;
USE librarydb1;

-- 1. Users Table
CREATE TABLE Users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(50) DEFAULT 'Staff',
    staff_number VARCHAR(50) DEFAULT '9876543210',
    staff_internal_number VARCHAR(50) DEFAULT 'BEL-EXT-101',
    sbu VARCHAR(100) DEFAULT 'General',
    dob VARCHAR(20) DEFAULT '15-08-1985',
    perm_books TINYINT(1) DEFAULT 1,
    perm_requests TINYINT(1) DEFAULT 1,
    perm_borrowings TINYINT(1) DEFAULT 1,
    perm_authors TINYINT(1) DEFAULT 0,
    perm_categories TINYINT(1) DEFAULT 0,
    perm_feedback TINYINT(1) DEFAULT 1,
    is_deleted TINYINT(1) DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 2. Authors Table
CREATE TABLE Authors (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(150) NOT NULL,
    bio TEXT,
    is_deleted TINYINT(1) DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 3. Categories Table
CREATE TABLE Categories (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    description TEXT,
    is_deleted TINYINT(1) DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 4. Books Table
CREATE TABLE Books (
    id INT AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    isbn VARCHAR(50) NOT NULL,
    publisher VARCHAR(150),
    author_id INT,
    category_id INT,
    copies_available INT DEFAULT 1,
    is_available TINYINT(1) DEFAULT 1,
    price DOUBLE DEFAULT 350.00,
    is_deleted TINYINT(1) DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (author_id) REFERENCES Authors(id) ON DELETE SET NULL,
    FOREIGN KEY (category_id) REFERENCES Categories(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 5. Borrowings Table
CREATE TABLE Borrowings (
    id INT AUTO_INCREMENT PRIMARY KEY,
    book_id INT NOT NULL,
    user_id INT NOT NULL,
    borrow_date DATE NOT NULL,
    due_date DATE NOT NULL,
    return_date DATE NULL,
    status VARCHAR(50) DEFAULT 'Borrowed',
    FOREIGN KEY (book_id) REFERENCES Books(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 6. BookRequests Table
CREATE TABLE BookRequests (
    id INT AUTO_INCREMENT PRIMARY KEY,
    book_id INT NOT NULL,
    user_id INT NOT NULL,
    staff_name VARCHAR(100),
    staff_number VARCHAR(50),
    staff_internal_number VARCHAR(50),
    book_name VARCHAR(255),
    author_name VARCHAR(100),
    price DOUBLE DEFAULT 0.00,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    status VARCHAR(50) DEFAULT 'Pending',
    approved_at DATETIME NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (book_id) REFERENCES Books(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 7. Feedbacks & Reviews Table (Max 150 words comment enforced)
CREATE TABLE Feedbacks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    book_id INT NULL,
    feedback_type VARCHAR(50) NOT NULL DEFAULT 'website', -- 'website' or 'book'
    rating INT NOT NULL DEFAULT 5,
    comment TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE CASCADE,
    FOREIGN KEY (book_id) REFERENCES Books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- ====================================================================
-- Performance Indexes (Eliminates full table scans on frequent queries)
-- ====================================================================

-- Borrowings: Most queried table — status filters, user/book lookups, date comparisons
CREATE INDEX idx_borrowings_status ON Borrowings (status);
CREATE INDEX idx_borrowings_book_id ON Borrowings (book_id);
CREATE INDEX idx_borrowings_user_id ON Borrowings (user_id);
CREATE INDEX idx_borrowings_user_status ON Borrowings (user_id, status);
CREATE INDEX idx_borrowings_book_status ON Borrowings (book_id, status);
CREATE INDEX idx_borrowings_due_date ON Borrowings (due_date);
CREATE INDEX idx_borrowings_borrow_date ON Borrowings (borrow_date);

-- BookRequests: Status + expiry lookups on every page load
CREATE INDEX idx_bookrequests_status ON BookRequests (status);
CREATE INDEX idx_bookrequests_user_id ON BookRequests (user_id);
CREATE INDEX idx_bookrequests_user_status ON BookRequests (user_id, status);
CREATE INDEX idx_bookrequests_approved_at ON BookRequests (approved_at);
CREATE INDEX idx_bookrequests_status_approved ON BookRequests (status, approved_at);

-- Books: Soft-delete filter, category lookups, availability checks
CREATE INDEX idx_books_is_deleted ON Books (is_deleted);
CREATE INDEX idx_books_category_id ON Books (category_id);
CREATE INDEX idx_books_category_deleted ON Books (category_id, is_deleted);
CREATE INDEX idx_books_available ON Books (is_available, is_deleted);

-- Feedbacks: Rating/review lookups per book
CREATE INDEX idx_feedbacks_book_id ON Feedbacks (book_id);

-- Users: Soft-delete filter, login lookups
CREATE INDEX idx_users_is_deleted ON Users (is_deleted);
CREATE INDEX idx_users_staff_number ON Users (staff_number);
CREATE INDEX idx_users_staff_internal ON Users (staff_internal_number);

-- Categories: Soft-delete filter
CREATE INDEX idx_categories_is_deleted ON Categories (is_deleted);

-- ====================================================================
-- Triggers for Automatic Borrowing Synchronization
-- ====================================================================

DELIMITER $$

DROP TRIGGER IF EXISTS trg_bookrequests_to_borrowings_insert$$
CREATE TRIGGER trg_bookrequests_to_borrowings_insert
AFTER INSERT ON BookRequests
FOR EACH ROW
BEGIN
    IF NEW.status = 'Approved' THEN
        INSERT INTO Borrowings (book_id, user_id, borrow_date, due_date, status)
        VALUES (NEW.book_id, NEW.user_id, NEW.start_date, NEW.end_date, 'Borrowed');
    END IF;
END$$

DROP TRIGGER IF EXISTS trg_bookrequests_to_borrowings_update$$
CREATE TRIGGER trg_bookrequests_to_borrowings_update
AFTER UPDATE ON BookRequests
FOR EACH ROW
BEGIN
    IF NEW.status = 'Approved' AND (OLD.status <> 'Approved' OR OLD.status IS NULL) THEN
        INSERT INTO Borrowings (book_id, user_id, borrow_date, due_date, status)
        VALUES (NEW.book_id, NEW.user_id, NEW.start_date, NEW.end_date, 'Borrowed');
    END IF;
END$$

DELIMITER ;

-- ====================================================================
-- Initial Seed Data (BEL Defense Electronics & Kannada Literature)
-- ====================================================================

-- 1. Users (Password is Date of Birth - DOB)
INSERT INTO Users (id, username, email, password_hash, dob, role, staff_number, staff_internal_number, sbu, perm_books, perm_requests, perm_borrowings, perm_authors, perm_categories, perm_feedback) VALUES
(1, 'Admin User', 'admin@bel.com', '15-08-1985', '15-08-1985', 'Admin', '9876543210', 'BEL-ADM-001', 'Central HR', 1, 1, 1, 1, 1, 1),
(2, 'Librarian Sarah', 'librarian@bel.com', '20-05-1992', '20-05-1992', 'Librarian', '9876543211', 'BEL-LIB-002', 'Library / HR', 1, 1, 1, 1, 1, 1),
(3, 'Rajesh Kumar', 'rajesh.k@bel.com', '10-10-1995', '10-10-1995', 'Staff', '9876543212', 'BEL-ENG-201', 'Radar SBU', 1, 1, 1, 0, 0, 1),
(4, 'Priya Sharma', 'priya.s@bel.com', '12-12-1996', '12-12-1996', 'Staff', '9876543213', 'BEL-ENG-202', 'MilCom SBU', 1, 1, 1, 0, 0, 1),
(5, 'Ramesh Rao', 'ramesh.r@bel.com', '04-07-1993', '04-07-1993', 'Staff', '9876543214', 'BEL-RAD-301', 'Naval Systems', 1, 1, 1, 0, 0, 1),
(6, 'Anita Desai', 'anita.d@bel.com', '25-01-1998', '25-01-1998', 'Staff', '9876543215', 'BEL-RES-401', 'CRL (Research)', 1, 1, 1, 0, 0, 1)
ON DUPLICATE KEY UPDATE username=VALUES(username), role=VALUES(role), password_hash=VALUES(password_hash), dob=VALUES(dob);

-- Seed Users (Admin, Librarian, Staff accounts)
-- Note: Sample books, authors, and categories have been removed to prepare for finalexcel.xlsx import.
