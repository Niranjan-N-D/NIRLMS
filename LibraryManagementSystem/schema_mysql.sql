-- ====================================================================
-- BEL Library Management System - MySQL Database Schema (SQLyog Compatible)
-- UTF-8 / Kannada Unicode Supported (utf8mb4)
-- ====================================================================

CREATE DATABASE IF NOT EXISTS librarydb DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE librarydb;

-- 1. Users Table
CREATE TABLE IF NOT EXISTS Users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(50) DEFAULT 'Staff',
    staff_number VARCHAR(50) DEFAULT '9876543210',
    staff_internal_number VARCHAR(50) DEFAULT 'BEL-EXT-101',
    dob VARCHAR(20) DEFAULT '15-08-1985',
    perm_books TINYINT(1) DEFAULT 1,
    perm_requests TINYINT(1) DEFAULT 1,
    perm_borrowings TINYINT(1) DEFAULT 1,
    perm_authors TINYINT(1) DEFAULT 0,
    perm_categories TINYINT(1) DEFAULT 0,
    perm_feedback TINYINT(1) DEFAULT 1,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 2. Authors Table
CREATE TABLE IF NOT EXISTS Authors (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(150) NOT NULL,
    bio TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 3. Categories Table
CREATE TABLE IF NOT EXISTS Categories (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    description TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 4. Books Table
CREATE TABLE IF NOT EXISTS Books (
    id INT AUTO_INCREMENT PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    isbn VARCHAR(50) NOT NULL UNIQUE,
    publisher VARCHAR(150),
    author_id INT,
    category_id INT,
    copies_available INT DEFAULT 1,
    is_available TINYINT(1) DEFAULT 1,
    price DOUBLE DEFAULT 350.00,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (author_id) REFERENCES Authors(id) ON DELETE SET NULL,
    FOREIGN KEY (category_id) REFERENCES Categories(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 5. Borrowings Table
CREATE TABLE IF NOT EXISTS Borrowings (
    id INT AUTO_INCREMENT PRIMARY KEY,
    book_id INT NOT NULL,
    user_id INT NOT NULL,
    borrow_date DATE NOT NULL,
    due_date DATE NOT NULL,
    return_date DATE NULL,
    status VARCHAR(50) DEFAULT 'Borrowed',
    FOREIGN KEY (book_id) REFERENCES Books(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 6. BookRequests Table
CREATE TABLE IF NOT EXISTS BookRequests (
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 7. Feedbacks & Reviews Table (Max 150 words comment enforced)
CREATE TABLE IF NOT EXISTS Feedbacks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    book_id INT NULL,
    feedback_type VARCHAR(50) NOT NULL DEFAULT 'website', -- 'website' or 'book'
    rating INT NOT NULL DEFAULT 5,
    comment TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE CASCADE,
    FOREIGN KEY (book_id) REFERENCES Books(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

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
INSERT INTO Users (id, username, email, password_hash, dob, role, staff_number, staff_internal_number, perm_books, perm_requests, perm_borrowings, perm_authors, perm_categories, perm_feedback) VALUES
(1, 'Admin User', 'admin@bel.com', '15-08-1985', '15-08-1985', 'Admin', '9876543210', 'BEL-ADM-001', 1, 1, 1, 1, 1, 1),
(2, 'Librarian Sarah', 'librarian@bel.com', '20-05-1992', '20-05-1992', 'Librarian', '9876543211', 'BEL-LIB-002', 1, 1, 1, 1, 1, 1),
(3, 'Rajesh Kumar', 'rajesh.k@bel.com', '10-10-1995', '10-10-1995', 'Staff', '9876543212', 'BEL-ENG-201', 1, 1, 1, 0, 0, 1),
(4, 'Priya Sharma', 'priya.s@bel.com', '12-12-1996', '12-12-1996', 'Staff', '9876543213', 'BEL-ENG-202', 1, 1, 1, 0, 0, 1),
(5, 'Ramesh Rao', 'ramesh.r@bel.com', '04-07-1993', '04-07-1993', 'Staff', '9876543214', 'BEL-RAD-301', 1, 1, 1, 0, 0, 1),
(6, 'Anita Desai', 'anita.d@bel.com', '25-01-1998', '25-01-1998', 'Staff', '9876543215', 'BEL-RES-401', 1, 1, 1, 0, 0, 1)
ON DUPLICATE KEY UPDATE username=VALUES(username), role=VALUES(role), password_hash=VALUES(password_hash), dob=VALUES(dob);

-- 2. Authors
INSERT INTO Authors (id, name, bio) VALUES
(1, 'Merrill Skolnik', 'Father of modern radar engineering, author of Introduction to Radar Systems and Radar Handbook.'),
(2, 'B. P. Lathi', 'Renowned professor and author in modern digital and analog communication systems.'),
(3, 'Robert C. Martin', 'Renowned software craftsman, advocate of Clean Code and agile software principles.'),
(4, 'Martin Fowler', 'Chief Scientist at ThoughtWorks, author of Refactoring and Enterprise Application Architecture.'),
(5, 'Dr. A. P. J. Abdul Kalam', 'India\'s Missile Man, 11th President of India, renowned aerospace engineer and author.'),
(6, 'ಕುವೆಂಪು (Kuvempu)', 'K. V. Puttappa - Legendary Kannada poet, novelist, and first Jnanapith Award winner for Kannada.'),
(7, 'ಎಸ್. ಎಲ್. ಭೈರಪ್ಪ (S. L. Bhyrappa)', 'Foremost Kannada novelist, philosopher, and Saraswati Samman laureate.'),
(8, 'ಡಾ. ಕೆ. ಶಿವರಾಮ ಕಾರಂತ (Dr. K. Shivarama Karanth)', 'Polymath, novelist, environmentalist, and Jnanapith Award winner.')
ON DUPLICATE KEY UPDATE name=VALUES(name), bio=VALUES(bio);

-- 3. Categories
INSERT INTO Categories (id, name, description) VALUES
(1, 'Radar & Defense Electronics', 'Radar systems, electronic warfare, missile guidance, and RF hardware engineering.'),
(2, 'Digital Signal Processing & Telecom', 'DSP algorithms, wireless communications, antenna design, and satellite links.'),
(3, 'Software & Embedded Systems', 'Embedded C/C++, RTOS, Clean Architecture, microservices, and design patterns.'),
(4, 'Aerospace & Space Technology', 'Avionics, propulsion, aerospace engineering, and defense technology leadership.'),
(5, 'ಕನ್ನಡ ಸಾಹಿತ್ಯ (Kannada Literature)', 'Timeless Kannada classics, epic poems, historical novels, and socio-cultural works.'),
(6, 'Leadership & Management', 'Strategic leadership, organizational culture, innovation, and project delivery.')
ON DUPLICATE KEY UPDATE name=VALUES(name), description=VALUES(description);

-- 4. Books
INSERT INTO Books (id, title, isbn, publisher, author_id, category_id, copies_available, is_available, price) VALUES
(1, 'Introduction to Radar Systems (3rd Ed)', '978-0072881387', 'McGraw-Hill Education', 1, 1, 6, 1, 895.00),
(2, 'Radar Handbook (3rd Edition)', '978-0071485470', 'McGraw-Hill Professional', 1, 1, 3, 1, 1450.00),
(3, 'Modern Digital and Analog Communication Systems', '978-0199370054', 'Oxford University Press', 2, 2, 5, 1, 750.00),
(4, 'Clean Code: A Handbook of Agile Software Craftsmanship', '978-0132350884', 'Prentice Hall', 3, 3, 4, 1, 620.00),
(5, 'Refactoring: Improving the Design of Existing Code', '978-0134757699', 'Addison-Wesley', 4, 3, 4, 1, 680.00),
(6, 'Wings of Fire: An Autobiography', '978-8173711463', 'Universities Press', 5, 4, 8, 1, 350.00),
(7, 'Ignited Minds: Unleashing the Power Within India', '978-0143424123', 'Penguin India', 5, 6, 6, 1, 299.00),
(8, 'ಕಾನೂರು ಹೆಗ್ಗಡಿತಿ (Kanooru Heggadithi)', '978-8128001015', 'Sapna Book House', 6, 5, 5, 1, 450.00),
(9, 'ಮಲೆಗಳಲ್ಲಿ ಮದುಮಗಳು (Malegalalli Madumagalu)', '978-8128001022', 'Vasantha Prakashana', 6, 5, 4, 1, 520.00),
(10, 'ಪರ್ವ (Parva)', '978-8172860100', 'Sahitya Bhandara', 7, 5, 4, 1, 480.00),
(11, 'ಗೃಹಭಂಗ (Gruhabhanga)', '978-8172860124', 'Sahitya Bhandara', 7, 5, 3, 1, 420.00),
(12, 'ಮೂಕಜ್ಜಿಯ ಕನಸುಗಳು (Mookajjiya Kanasugalu)', '978-8172860117', 'SBS Publishers', 8, 5, 3, 1, 395.00)
ON DUPLICATE KEY UPDATE title=VALUES(title), price=VALUES(price), copies_available=VALUES(copies_available);

-- 5. Borrowings
INSERT INTO Borrowings (id, book_id, user_id, borrow_date, due_date, return_date, status) VALUES
(1, 1, 3, '2026-10-18', '2026-11-02', NULL, 'Borrowed'),
(2, 4, 4, '2026-10-20', '2026-11-04', NULL, 'Borrowed'),
(3, 8, 5, '2026-10-22', '2026-11-06', NULL, 'Borrowed'),
(4, 2, 3, '2026-10-01', '2026-10-16', NULL, 'Overdue'),
(5, 10, 6, '2026-09-25', '2026-10-10', NULL, 'Overdue'),
(6, 3, 4, '2026-09-10', '2026-09-25', '2026-09-23', 'Returned'),
(7, 6, 5, '2026-09-15', '2026-09-30', '2026-09-28', 'Returned'),
(8, 12, 6, '2026-09-18', '2026-10-03', '2026-10-01', 'Returned')
ON DUPLICATE KEY UPDATE status=VALUES(status);

-- 6. Book Requests
INSERT INTO BookRequests (id, book_id, user_id, staff_name, staff_number, staff_internal_number, book_name, author_name, price, start_date, end_date, status, approved_at) VALUES
(1, 5, 3, 'Rajesh Kumar', '9876543212', 'BEL-ENG-201', 'Refactoring: Improving the Design of Existing Code', 'Martin Fowler', 680.00, '2026-10-26', '2026-11-10', 'Pending', NULL),
(2, 7, 4, 'Priya Sharma', '9876543213', 'BEL-ENG-202', 'Ignited Minds: Unleashing the Power Within India', 'Dr. A. P. J. Abdul Kalam', 299.00, '2026-10-27', '2026-11-11', 'Pending', NULL),
(3, 9, 5, 'Ramesh Rao', '9876543214', 'BEL-RAD-301', 'ಮಲೆಗಳಲ್ಲಿ ಮದುಮಗಳು (Malegalalli Madumagalu)', 'ಕುವೆಂಪು (Kuvempu)', 520.00, '2026-10-15', '2026-10-30', 'Approved', '2026-10-15 11:30:00'),
(4, 11, 6, 'Anita Desai', '9876543215', 'BEL-RES-401', 'ಗೃಹಭಂಗ (Gruhabhanga)', 'ಎಸ್. ಎಲ್. ಭೈರಪ್ಪ (S. L. Bhyrappa)', 420.00, '2026-10-10', '2026-10-25', 'Rejected', NULL)
ON DUPLICATE KEY UPDATE status=VALUES(status);

-- 7. Feedback & Reviews
INSERT INTO Feedbacks (id, user_id, book_id, feedback_type, rating, comment) VALUES
(1, 3, NULL, 'website', 5, 'The BEL Library portal is fast, modern, and very easy to browse technical radar books.'),
(2, 4, 4, 'book', 5, 'Clean Code by Uncle Bob is mandatory reading for all software engineers at BEL!'),
(3, 5, 8, 'book', 5, 'ಕಾನೂರು ಹೆಗ್ಗಡಿತಿ ಕಾದಂಬರಿ ಕನ್ನಡ ಸಾಹಿತ್ಯದ ಅದ್ಭುತ ಕೃತಿ. ಮಲೆನಾಡಿನ ಜನಜೀವನದ ನೈಜ ಚಿತ್ರಣ ಇಲ್ಲಿದೆ.'),
(4, 6, 6, 'book', 5, 'Wings of Fire is an extraordinary source of motivation for every scientist and defense engineer.')
ON DUPLICATE KEY UPDATE comment=VALUES(comment), rating=VALUES(rating);
