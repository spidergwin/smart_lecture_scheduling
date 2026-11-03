Design a comprehensive, production-grade mobile and tablet application for a university called "Smart Lecture Scheduling and Academic Management System" (Ignatius Ajuru University of Education - IAUE).

Style and Appearance: Let the design tool automatically generate a clean, modern, cohesive, and high-contrast professional color palette suitable for an academic university application with elegant card elevations, clean typography, rounded surfaces, and clear visual hierarchy.

Generate all interconnected screens in one single cohesive prototype flow:

1. Splash & Onboarding Screen: University logo branding ("IAUE Smart Scheduler & Academic Portal"), welcome illustration, tagline "Effortless lecture scheduling, real-time alerts, and seamless textbook management", brief feature highlights, and a primary "Get Started" button leading to login.

2. Authentication & Role Switcher Screen: Top segmented tab switch for [Student | Lecturer | Administrator]. Login card with Institutional Email input ("student.nwabueze@iaue.edu.ng"), Password input with show/hide toggle, "Remember Me" checkbox, "Forgot Password?" link, primary "Sign In" button, and a "Register New Account" link.

3. Student Daily Timetable Dashboard: Top header with student profile ("Wisdom Nwabueze", Matric No: U/2022/22225, Computer Science, 400 Level), notification bell with unread badge count, and horizontal weekday date selector pills. Main feed of chronological lecture cards with prominent status badges:
   - Lecture Card 1 (08:00 AM - 10:00 AM): "CSC 424 - Models of Computation", Lecturer: Dr. Goodluck Nwagor, Venue: Computer Science Lab 1, Status Badge: [ON SCHEDULE].
   - Lecture Card 2 (10:00 AM - 12:00 PM): "CSC 423 - Compiler Construction", Lecturer: Dr. ChiefJames M. Phillip, Venue: Faculty Auditorium B, Status Badge: [RELOCATED FROM HALL 2].
   - Lecture Card 3 (02:00 PM - 04:00 PM): "CSC 422 - Human Computer Interface", Lecturer: Dr. O. M. Thom-Manuel, Venue: ETF Hall 3, Status Badge: [CANCELLED - Emergency Departmental Meeting].
   - Lecture Card 4 (04:00 PM - 06:00 PM): "CSC 421 - Net-Centric Computing", Lecturer: Dr. Aburuotu, Venue: ICT Center Hall A, Status Badge: [ON SCHEDULE].
   Bottom Navigation Bar with tabs: Timetable, Textbooks, Notifications, Profile.

4. Weekly Timetable Grid View: Full Monday to Saturday schedule matrix with semester selector ("First Semester 2026/2027"), search bar for course codes (CSC 421, CSC 422, CSC 423, CSC 424), and interactive time blocks with venue chips.

5. Lecture Detail & Venue Guide Screen: Detailed course view displaying Course Code & Title ("CSC 424 - Models of Computation", 3 Credit Units), Lecturer info card (Dr. Goodluck Nwagor with message/email button), Venue details ("Computer Science Lab 1, Faculty of Natural & Applied Sciences"), campus walking route guidance card, and a button to view required course textbook.

6. Real-Time Notification Center / Alert Inbox: Filter tabs [All | Urgent | Reminders | Textbooks]. List of notification cards:
   - Urgent alert: "CSC 422 Cancelled Today by Dr. O. M. Thom-Manuel (12:15 PM)".
   - Reminder: "CSC 424 Lecture starts in 30 minutes at Computer Science Lab 1".
   - Venue change: "CSC 423 relocated to Faculty Auditorium B".
   - Textbook update: "CSC 424 official textbook now available for pickup via Sales Rep (Matric: U/2022/22225)".
   Each notification has timestamp, read/unread indicator dot, and tap-to-expand details.

7. Lock Screen Push Notification Simulation: Smartphone lock screen displaying time (12:15 PM), date, and high-priority push banner: "SMART SCHEDULER: CSC 422 Lecture Cancelled by Dr. O. M. Thom-Manuel. Reason: Urgent Faculty Board Meeting. Please do not commute to campus."

8. Lecturer Schedule & Course Management Dashboard: Lecturer control center ("Welcome, Dr. Goodluck Nwagor") displaying assigned courses (CSC 424 Models of Computation), upcoming classes, quick action button "Schedule New Lecture", and button "Manage Course Textbooks & Sales Reps". Includes schedule creation modal: Course selector, Day of Week picker, Time start/end, and Venue selector.

9. Hall Conflict Detection & Alternative Venue Modal: Alert pop-up displayed when double-booking occurs: Warning banner "[!] CONFLICT DETECTED - Venue Overlap", stating "Computer Science Lab 1 is already booked for another lecture at this time". Shows automated list of "Recommended Available Alternative Venues" (ICT Center Hall A, Faculty Auditorium B) with a "Select Alternative & Broadcast Push Notice" button.

10. Lecturer Textbook Sales & Student Rep Delegation Console: Lecturer management screen for textbook distribution:
   - Course Textbook details: Title ("Models of Computation Course Handbook"), Unit Price, Total Stock, Total Sold, Total Revenue.
   - "Assign Student Sales Representative" section: Input field to enter Student Matriculation Number (e.g. "U/2022/22225"), textbook quota assigned to this rep, and "Assign Rep" button.
   - List of active sales reps with their individual sales counters and assigned inventory.
   - "Course Payment QR Code" display with a prominent "Export / Share QR Code" button for students to scan and pay.
   - Button to "View Full Student Payment Roster".

11. Lecturer Master Student Textbook Roster & Export Screen: Comprehensive student payment and distribution audit table:
   - Search bar (search by Student Name or Matric No) and Filter pills [All | Paid & Issued | Paid (Pending Pickup) | Unpaid].
   - Paginated data table with columns: Matric Number, Student Name, Course Code (CSC 424), Amount Paid, Payment Status (Paid / Pending), Assigned Sales Rep, Issue Date.
   - Pagination footer: "Showing 1-10 of 120 students", Previous / Next page buttons, and page numbers [1, 2, 3... 12].
   - Top action bar with prominent "Export as Excel (.xlsx)" and "Export as CSV" buttons.

12. Assigned Student Sales Representative Console: Dedicated mobile interface for students assigned by the lecturer to manage textbook sales:
   - Rep Overview: Assigned Course ("CSC 424 - Dr. Goodluck Nwagor"), Books Assigned to Me, Books Sold, Remaining Stock, Cash/Transfer Total Collected.
   - Quick Action: "Scan Student QR Code" button (opens camera viewfinder to scan student's payment/receipt code).
   - Manual verification form: Search/Enter Buyer's Matric No to manually mark as "Paid & Book Issued".
   - Rep's Personal Sales Log: Searchable list of all students this rep has issued books to with timestamps and payment status, plus "Export My Sales Report" button.

13. Student Textbook Store & QR Payment Screen: Student-facing textbook purchasing view:
   - List of required course textbooks:
     * "CSC 424: Models of Computation" by Dr. Goodluck Nwagor
     * "CSC 423: Compiler Construction" by Dr. ChiefJames M. Phillip
     * "CSC 422: Human Computer Interface" by Dr. O. M. Thom-Manuel
     * "CSC 421: Net-Centric Computing" by Dr. Aburuotu
   - Card showing price, lecturer endorsement, and assigned Student Sales Rep pickup details.
   - Payment & Collection section: Scan Lecturer's Official Payment QR Code or make instant bank transfer, upload payment receipt, or generate Student Pickup QR Code to present to the sales rep for instant verification.
   - Status badge: [PAID - READY FOR PICKUP] or [COMPLETED - ISSUED].

14. Student Course Enrollment & App Settings: Course enrollment checklist for 400 Level First Semester (CSC 421, CSC 422, CSC 423, CSC 424) with total credit unit counter, notification reminder preferences (toggles for 60m and 30m before lecture, instant emergency cancellation alert), dark mode toggle, and logout button.
