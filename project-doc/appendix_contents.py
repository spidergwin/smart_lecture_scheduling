# -*- coding: utf-8 -*-
"""
Appendix A (Source Code) and Appendix B (Sample Outputs) content
"""

APP_A_SECTIONS = [
    {
        "type": "heading1",
        "text": "APPENDIX A: CORE SYSTEM SOURCE CODE",
        "bookmark": "_Toc_APPA"
    },
    {
        "type": "heading2",
        "text": "A.1 Database Relational Schema (PostgreSQL DDL: schema.sql)",
        "content": """-- ============================================================
-- Smart Lecture Scheduling and Notification System
-- Database Schema: PostgreSQL 16
-- Author: NWABUEZE, IHIECHUKWU WISDOM (Matric: U/2022/22225)
-- Ignatius Ajuru University of Education (IAUE)
-- ============================================================

CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Table 1: Users Table
CREATE TABLE users (
    id SERIAL PRIMARY KEY,
    full_name VARCHAR(150) NOT NULL,
    email VARCHAR(150) UNIQUE NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(20) NOT NULL CHECK (role IN ('student', 'lecturer', 'admin')),
    matric_or_staff_no VARCHAR(50) UNIQUE NOT NULL,
    department VARCHAR(100) NOT NULL,
    level VARCHAR(10) DEFAULT '100',
    device_token TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Table 2: Venues / Lecture Halls Table
CREATE TABLE venues (
    id SERIAL PRIMARY KEY,
    name VARCHAR(100) UNIQUE NOT NULL,
    building VARCHAR(100) NOT NULL,
    capacity INTEGER NOT NULL CHECK (capacity > 0),
    has_projector BOOLEAN DEFAULT FALSE,
    has_sound_system BOOLEAN DEFAULT FALSE,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Table 3: Courses Table
CREATE TABLE courses (
    id SERIAL PRIMARY KEY,
    course_code VARCHAR(20) UNIQUE NOT NULL,
    course_title VARCHAR(200) NOT NULL,
    credit_units INTEGER NOT NULL CHECK (credit_units BETWEEN 1 AND 6),
    department VARCHAR(100) NOT NULL,
    level VARCHAR(10) NOT NULL,
    semester VARCHAR(10) NOT NULL CHECK (semester IN ('FIRST', 'SECOND')),
    lecturer_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Table 4: Lecture Schedules Table
CREATE TABLE schedules (
    id SERIAL PRIMARY KEY,
    course_id INTEGER NOT NULL REFERENCES courses(id) ON DELETE CASCADE,
    lecturer_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    venue_id INTEGER NOT NULL REFERENCES venues(id) ON DELETE RESTRICT,
    day_of_week VARCHAR(15) NOT NULL CHECK (day_of_week IN ('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')),
    start_time TIME NOT NULL,
    end_time TIME NOT NULL,
    status VARCHAR(20) DEFAULT 'SCHEDULED' CHECK (status IN ('SCHEDULED','RESCHEDULED','CANCELLED','COMPLETED')),
    cancellation_reason TEXT,
    reminder_sent BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT chk_time_order CHECK (end_time > start_time)
);

-- Table 5: Notifications Table
CREATE TABLE notifications (
    id SERIAL PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    schedule_id INTEGER REFERENCES schedules(id) ON DELETE SET NULL,
    title VARCHAR(200) NOT NULL,
    message TEXT NOT NULL,
    type VARCHAR(30) NOT NULL CHECK (type IN ('REMINDER','RESCHEDULE','CANCELLATION','ANNOUNCEMENT')),
    is_read BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Performance Indexes
CREATE INDEX idx_schedules_venue_time ON schedules (venue_id, day_of_week, start_time, end_time);
CREATE INDEX idx_schedules_lecturer ON schedules (lecturer_id, day_of_week);
CREATE INDEX idx_notifications_user ON notifications (user_id, is_read);"""
    },
    {
        "type": "heading2",
        "text": "A.2 Algorithmic Hall-Matching & Conflict Engine (scheduleService.js)",
        "content": """// ============================================================
// Lecture Scheduling & Hall Conflict Detection Engine
// File: src/services/scheduleService.js
// ============================================================

const pool = require('../db/pool');

class ScheduleService {
  /**
   * Verify if a proposed lecture slot conflicts with existing bookings
   */
  static async checkConflicts(venueId, lecturerId, dayOfWeek, startTime, endTime, excludeScheduleId = null) {
    // 1. Check for Venue Overlap
    let venueQuery = `
      SELECT s.id, c.course_code, v.name AS venue_name, s.start_time, s.end_time
      FROM schedules s
      JOIN courses c ON s.course_id = c.id
      JOIN venues v ON s.venue_id = v.id
      WHERE s.venue_id = $1 
        AND s.day_of_week = $2
        AND s.status != 'CANCELLED'
        AND ((s.start_time, s.end_time) OVERLAPS ($3::time, $4::time))
    `;
    const venueParams = [venueId, dayOfWeek, startTime, endTime];
    if (excludeScheduleId) {
      venueQuery += ` AND s.id != $5`;
      venueParams.push(excludeScheduleId);
    }
    const venueRes = await pool.query(venueQuery, venueParams);
    if (venueRes.rows.length > 0) {
      return { hasConflict: true, type: 'VENUE_CLASH', conflictWith: venueRes.rows[0] };
    }

    // 2. Check for Lecturer Overlap
    let lecturerQuery = `
      SELECT s.id, c.course_code, s.start_time, s.end_time
      FROM schedules s
      JOIN courses c ON s.course_id = c.id
      WHERE s.lecturer_id = $1
        AND s.day_of_week = $2
        AND s.status != 'CANCELLED'
        AND ((s.start_time, s.end_time) OVERLAPS ($3::time, $4::time))
    `;
    const lecturerParams = [lecturerId, dayOfWeek, startTime, endTime];
    if (excludeScheduleId) {
      lecturerQuery += ` AND s.id != $5`;
      lecturerParams.push(excludeScheduleId);
    }
    const lecturerRes = await pool.query(lecturerQuery, lecturerParams);
    if (lecturerRes.rows.length > 0) {
      return { hasConflict: true, type: 'LECTURER_CLASH', conflictWith: lecturerRes.rows[0] };
    }

    return { hasConflict: false };
  }

  /**
   * Suggest Alternative Available Halls If Desired Hall is Congested
   */
  static async findAlternativeHalls(dayOfWeek, startTime, endTime, minCapacity = 50) {
    const query = `
      SELECT v.id, v.name, v.building, v.capacity
      FROM venues v
      WHERE v.capacity >= $1
        AND v.is_active = TRUE
        AND v.id NOT IN (
          SELECT s.venue_id FROM schedules s
          WHERE s.day_of_week = $2
            AND s.status != 'CANCELLED'
            AND ((s.start_time, s.end_time) OVERLAPS ($3::time, $4::time))
        )
      ORDER BY v.capacity ASC
      LIMIT 5;
    `;
    const res = await pool.query(query, [minCapacity, dayOfWeek, startTime, endTime]);
    return res.rows;
  }
}

module.exports = ScheduleService;"""
    },
    {
        "type": "heading2",
        "text": "A.3 Firebase Push Notification Dispatcher (notificationService.js)",
        "content": """// ============================================================
// Firebase Cloud Messaging (FCM) Multicast Push Dispatcher
// File: src/services/notificationService.js
// ============================================================

const admin = require('firebase-admin');
const pool = require('../db/pool');

class NotificationService {
  /**
   * Dispatch real-time push notifications to all students enrolled in a course
   */
  static async broadcastLectureUpdate(courseId, scheduleId, title, body, notificationType) {
    // 1. Fetch registered student device tokens for this course
    const tokenQuery = `
      SELECT u.id, u.device_token 
      FROM users u
      JOIN student_courses sc ON u.id = sc.student_id
      WHERE sc.course_id = $1 AND u.device_token IS NOT NULL AND u.device_token != '';
    `;
    const tokenRes = await pool.query(tokenQuery, [courseId]);
    if (tokenRes.rows.length === 0) return { deliveredCount: 0 };

    const tokens = tokenRes.rows.map(r => r.device_token);

    // 2. Persist in-app notification records for audit history
    for (const student of tokenRes.rows) {
      await pool.query(
        `INSERT INTO notifications (user_id, schedule_id, title, message, type)
         VALUES ($1, $2, $3, $4, $5)`,
        [student.id, scheduleId, title, body, notificationType]
      );
    }

    // 3. Construct multicast payload for Firebase Admin SDK
    const messagePayload = {
      notification: { title, body },
      data: {
        scheduleId: String(scheduleId),
        type: notificationType,
        click_action: 'FLUTTER_NOTIFICATION_CLICK'
      },
      tokens: tokens
    };

    const response = await admin.messaging().sendEachForMulticast(messagePayload);
    console.log(`[FCM Broadcast] Delivered: ${response.successCount}, Failed: ${response.failureCount}`);
    return { deliveredCount: response.successCount };
  }
}

module.exports = NotificationService;"""
    },
    {
        "type": "heading2",
        "text": "A.4 Automated Reminder Cron Scheduler (reminderCron.js)",
        "content": """// ============================================================
// Automated Lecture Reminder Background Worker
// File: src/cron/reminderCron.js
// ============================================================

const cron = require('node-cron');
const pool = require('../db/pool');
const NotificationService = require('../services/notificationService');

function initReminderCron() {
  // Execute every 60 seconds (* * * * *)
  cron.schedule('* * * * *', async () => {
    try {
      const now = new Date();
      const currentDay = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][now.getDay()];
      
      // Look for confirmed lectures starting within 30 to 60 minutes
      const query = `
        SELECT s.id, s.course_id, c.course_code, c.course_title, v.name AS venue_name, s.start_time
        FROM schedules s
        JOIN courses c ON s.course_id = c.id
        JOIN venues v ON s.venue_id = v.id
        WHERE s.day_of_week = $1
          AND s.status = 'SCHEDULED'
          AND s.reminder_sent = FALSE
          AND s.start_time BETWEEN CURRENT_TIME AND (CURRENT_TIME + INTERVAL '45 minutes');
      `;
      const res = await pool.query(query, [currentDay]);

      for (const schedule of res.rows) {
        const title = `Lecture Reminder: ${schedule.course_code}`;
        const body = `Your ${schedule.course_code} lecture is starting at ${schedule.start_time} in ${schedule.venue_name}. Please be punctual!`;
        
        await NotificationService.broadcastLectureUpdate(
          schedule.course_id, schedule.id, title, body, 'REMINDER'
        );

        // Mark reminder as sent
        await pool.query('UPDATE schedules SET reminder_sent = TRUE WHERE id = $1', [schedule.id]);
      }
    } catch (err) {
      console.error('[Cron Error] Reminder worker failure:', err.message);
    }
  });
}

module.exports = initReminderCron;"""
    },
    {
        "type": "heading2",
        "text": "A.5 Authentication & Role Authorization Middleware (authMiddleware.js)",
        "content": """// ============================================================
// Authentication & Role-Based Access Control Middleware
// File: src/middlewares/authMiddleware.js
// ============================================================

const jwt = require('jsonwebtoken');

function authMiddleware(req, res, next) {
  const authHeader = req.headers['authorization'];
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ success: false, message: 'Authentication required. No Bearer token provided.' });
  }

  const token = authHeader.split(' ')[1];
  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = decoded; // { id, email, role, matric_or_staff_no }
    next();
  } catch (error) {
    return res.status(401).json({ success: false, message: 'Invalid or expired session token.' });
  }
}

function roleGuard(allowedRoles = []) {
  return (req, res, next) => {
    if (!req.user || !allowedRoles.includes(req.user.role)) {
      return res.status(403).json({ success: false, message: 'Access denied: Insufficient privileges.' });
    }
    next();
  };
}

module.exports = { authMiddleware, roleGuard };"""
    }
]

APP_B_SECTIONS = [
    {
        "type": "heading1",
        "text": "APPENDIX B: SAMPLE SYSTEM OUTPUTS",
        "bookmark": "_Toc_APPB"
    },
    {
        "type": "heading2",
        "text": "B.1 User Authentication and Login Screen",
        "content": """======================================================================
               SMART LECTURE SCHEDULING SYSTEM (IAUE)
                     Mobile User Authentication
======================================================================
 [ Logo: Ignatius Ajuru University of Education, Port Harcourt ]

  Welcome to Smart Lecture Scheduler
  Sign in to view timetables and receive real-time lecture alerts.

  Institutional Email:
  [ student.nwabueze@iaue.edu.ng                          ]

  Password:
  [ **********************                               ]

  [  SIGN IN  ]

  Forgot Password? | New Student? Create Account Here.
======================================================================"""
    },
    {
        "type": "heading2",
        "text": "B.2 Student Personalized Weekly Timetable Dashboard",
        "content": """======================================================================
  TIMETABLE: COMPUTER SCIENCE (400 LEVEL) - FIRST SEMESTER
======================================================================
  Today: Monday, 27 October 2026

  [08:00 AM - 10:00 AM]  CONFIRMED
  --------------------------------------------------------------------
  Course:     CMP 411 - Compiler Construction (3 Units)
  Lecturer:   Dr. Goodluck Nwagor
  Venue:      Computer Science Laboratory 1 (Main Campus)
  Status:     [ ON SCHEDULE ]

  [10:00 AM - 12:00 PM]  VENUE RELOCATED
  --------------------------------------------------------------------
  Course:     CMP 413 - Software Engineering (3 Units)
  Lecturer:   Prof. A. A. Adebayo
  Venue:      Faculty Auditorium B  <-- [RELOCATED FROM HALL 2]
  Status:     [ RELOCATION NOTICE SENT VIA PUSH ]

  [02:00 PM - 04:00 PM]  CANCELLED
  --------------------------------------------------------------------
  Course:     CMP 415 - Artificial Intelligence (3 Units)
  Lecturer:   Dr. M. Abubakar
  Reason:     Emergency Departmental Academic Board Meeting
  Status:     [ CANCELLED - COMMUTE PREVENTED ]
======================================================================"""
    },
    {
        "type": "heading2",
        "text": "B.3 Lecturer Schedule Creation & Hall Conflict Alert Interface",
        "content": """======================================================================
                LECTURER SCHEDULE MANAGEMENT CONSOLE
======================================================================
  Course Selected:  CMP 411 (Compiler Construction)
  Day of Week:      Tuesday
  Time Window:      10:00 AM - 12:00 PM
  Requested Venue:  Computer Lab 1

  [ SYSTEM CONFLICT DETECTION RESULT ]
  --------------------------------------------------------------------
  STATUS: [!] CONFLICT DETECTED - HTTP 409 CONFLICT
  Conflict Reason: Computer Lab 1 is already allocated to CMP 311 
                   from 09:30 AM to 11:30 AM.

  RECOMMENDED AVAILABLE ALTERNATIVE VENUES:
  [1] ICT Centre Hall A     (Capacity: 120, Projector: YES)  [ SELECT ]
  [2] Faculty Lecture Hall 3 (Capacity: 150, Projector: YES)  [ SELECT ]
  [3] ETF Building Room 4   (Capacity: 100, Projector: NO)   [ SELECT ]

  [ Confirm Venue Selection & Broadcast Schedule to Students ]
======================================================================"""
    },
    {
        "type": "heading2",
        "text": "B.4 Real-Time Mobile Push Notification & Notification Inbox",
        "content": """======================================================================
              SMART LECTURE SCHEDULER - NOTIFICATION INBOX
======================================================================
  [PUSH ALERT BANNER - LOCK SCREEN]
  "SMART SCHEDULER: CMP 415 CANCELLED TODAY"
  "Dr. Abubakar has cancelled today's 2:00 PM lecture due to an urgent
   Faculty Board meeting. Please do not commute to Campus."
  Time: 12:15 PM (Received 1.8s after dispatch)

  --------------------------------------------------------------------
  IN-APP NOTIFICATION HISTORY:
  --------------------------------------------------------------------
  1. [URGENT] CMP 415 Cancelled by Dr. M. Abubakar
     12:15 PM | Unread | Reason: Departmental Academic Board Meeting

  2. [REMINDER] CMP 411 Lecture in 30 Minutes
     07:30 AM | Read   | Venue: Computer Science Laboratory 1

  3. [VENUE CHANGE] CMP 413 Relocated to Faculty Auditorium B
     Yesterday | Read  | Venue adjusted due to ongoing hall repairs

  4. [ANNOUNCEMENT] CMP 411 Term Paper Assignment Due Friday
     2 days ago | Read | Upload submission files through student portal
======================================================================"""
    }
]

print("appendix_contents.py successfully loaded!")
