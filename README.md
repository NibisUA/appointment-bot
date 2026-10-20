# appointment-bot
Telegram bot for appointment booking built with Aiogram. Users can choose service and time, book appointments, add comments, receive reminders and contact the master. Admin panel includes calendar, schedule management, client list, broadcast messages and Excel export of bookings.


📅 Appointment Booking Telegram Bot

A production-ready Telegram bot for appointment scheduling and service management built with Aiogram.

Designed for real-world use in service-based businesses (salons, private specialists, tutors) to automate booking workflows and reduce manual scheduling overhead.

⚡ Highlights (Recruiter Summary)
Full appointment booking system inside Telegram
Admin-controlled scheduling with calendar view
Real-time client management (contacts, bookings, notes)
Automated reminders and notifications
Excel export for reporting and analytics
FSM-based structured conversation flow
Extensible architecture for scaling features


👤 Client Flow
Select service type
Choose available time slot
Add optional comment to booking
Receive confirmation and reminders
Contact service provider via Telegram


🧑‍💼 Admin Features
Interactive calendar of all appointments
Manage working days and time slots dynamically
View full client details (Telegram ID, username, phone)
Cancel or modify bookings
Broadcast messages to all users
Export all data to Excel


🧠 Architecture
FSM-based interaction system (Aiogram) for structured user flows
Dynamic scheduling engine based on admin-defined availability
Relational database layer for persistent storage
Admin control module for real-time management
Notification system for reminders and broadcasts


🛠 Tech Stack
Python 3.x
Aiogram (Telegram Bot API framework)
SQLite / PostgreSQL
Pandas / OpenPyXL (data export)
FSM (Finite State Machine)


📊 Data Model
users — Telegram user profiles
appointments — bookings with status, time, comments
schedule — available working days and slots
services — list of offered services
🎯 Business Use Case

This bot replaces manual appointment tracking (messages, spreadsheets, calls) with a fully automated Telegram-based scheduling system.

Suitable for:

Beauty salons
Freelancers / private masters
Tutors
Clinics / service providers
🚀 Roadmap
Payment integration
Google Calendar sync
Multi-admin support
Web dashboard for admin panel
REST API for external integrations
