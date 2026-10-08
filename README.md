# AttendPro — Employee Attendance, HR & Payroll Management System

A full-featured Django system for employee attendance, HR documentation, and
lightweight payroll record-keeping, built for small-to-mid-size UAE
businesses. Orange & White theme throughout.

---

## Features

### Core — Employees & Attendance
- Employee management (add/edit/delete/detail view, calendar view, photo, status, salary)
- Daily attendance marking (Present / Absent / Half Day / Holiday / Leave, In/Out time, OT hours)
- Holiday management (public holidays + auto-generate Sundays for a year)
- Leave management (leave types, requests, approve/reject)
- Bulk import (employees and attendance via Excel templates, with validation + duplicate handling)
- Employee self-service portal (separate login for employees to view their own attendance)

### Documents & Compliance
- Per-employee document tracking (Emirates ID, Passport, Visa, Labour Card, etc.)
- Configurable document types with expiry alert windows
- Dedicated "Expiring Documents" page, plus a live widget on the Dashboard

### Letters (full CRUD)
Two categories, five letter types, each generated as both **Word (.docx)** and
**PDF**, from the same data so both always match:

**Employee Request Letters** (no company letterhead — written by the employee):
- Salary Hike Request Letter
- Leave Request Letter (with optional local mobile/address while on leave, kept for internal records only, never printed)

**HR Issued Letters** (full company letterhead):
- Salary Hike Approval Letter
- Experience Certificate
- No Objection Certificate (NOC)
- Warning / Disciplinary Letter (with Employee Acknowledgement signature line)

Every letter supports Create, Edit, View (live preview), Delete, and a full
searchable/filterable list at **Letters → All Letters**. Each employee's
profile has its own **Letters** tab showing everything issued to/by them.

### Settlements
Manual entry log for periodic OT payments and leave/absence deductions
(no auto-calculation — figures are entered after HR's own calculation).
Full list with filtering, add/edit/delete, and a **Settlements** tab on each
employee's profile showing their full history with the most recent
settlement highlighted.

### Temporary / Outsource Attendance
A separate, lightweight system for ad-hoc hourly workers (no Emirates ID or
full onboarding needed):
- Temporary Worker list (name, designation, mobile, optional hourly rate)
- Per-worker attendance view page with its own totals and date filter
- Bulk "Mark Attendance" by date — enter hours worked per worker; leaving a
  field blank means that worker wasn't present that day (no record created)
- Dedicated **Report** page — select one or more workers, a date range, and
  extra export options (include notes, include auto-calculated summary
  section, sort by worker or chronologically) before exporting to Excel
- Excel export always includes an auto-totalled summary section (days
  worked, total hours, and total amount where an hourly rate is set) —
  individual day entries are manual, totals are automatic

### Reports & Export
- Monthly attendance summary, OT report, Absent frequency report, Late arrivals report — each with its own Excel export
- Full-period attendance export (PDF and Excel, multi-employee, date range, ABSENT in red, holidays highlighted)
- Activity Log — every create/update/delete/export is tracked with user, timestamp, and description

### AI Assistant
Built-in chat assistant (OpenRouter/Groq, with Gemini fallback) that answers
questions about today's attendance, monthly summaries, a specific employee's
record, document expiry, or late arrivals — in plain language — and can
export the answer straight to Excel.

### Global Search
A single search bar (top bar, every page) across Employees, Documents,
Letters (including by reference number), and Settlements at once, grouped
by category on the results page.

### Admin & Security
- Granular per-user permissions (see table below) — the sidebar only shows what a user is allowed to access
- Company Settings page (name, address, logo, contact info, work hours, timezone, date format)
- Activity log of every action across the system

---

## Tech Stack
- **Backend:** Django 4.2 (Python 3.10+)
- **Database:** SQLite (default — swappable via `DATABASES` in settings.py)
- **Documents:** `python-docx` (Word), `reportlab` (PDF), `openpyxl` (Excel)
- **Static files:** WhiteNoise
- **Deployment:** `passenger_wsgi.py` included for cPanel/Passenger hosting; `gunicorn` for standard WSGI hosting

---

## Project Structure (key files)

```
attendance_system/
├── attendance_system/          # Django project settings, root urls
├── attendance_app/
│   ├── models.py                # All data models
│   ├── views.py                 # Core views (employees, attendance, holidays, leaves, reports, documents, import, settings, portal)
│   ├── ai_views.py              # AI chat assistant
│   ├── letters.py               # Letter generation engine (DOCX + PDF builders, all 5 letter types)
│   ├── letter_views.py          # Letter CRUD views
│   ├── settlement_views.py      # Settlements CRUD views
│   ├── search_views.py          # Global search
│   ├── temp_attendance_views.py # Temporary worker + attendance views
│   ├── middleware.py            # Activity logging
│   ├── decorators.py            # Permission decorator
│   ├── context_processors.py    # Injects permissions into every template
│   ├── templatetags/            # Custom template filters (e.g. markdown-bold rendering for letter previews)
│   ├── migrations/
│   └── templates/attendance_app/
├── requirements.txt
├── manage.py
└── README.md
```

---

## Quick Setup

### Requirements
- Python 3.10+
- pip

### Installation

```bash
# 1. Navigate to project folder
cd attendance_system

# 2. (Optional) Create virtual environment
python -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Apply database migrations
python manage.py migrate

# 5. Create an admin account (if you don't already have db.sqlite3 with one)
python manage.py createsuperuser

# 6. Start server
python manage.py runserver
```

### Access
- URL: http://127.0.0.1:8000
- Log in with the superuser account created above (or the existing admin account in your `db.sqlite3` if migrating an existing install)

> ⚠️ Always change default passwords after first login.

---

## Company Customization

Two ways to set your company details:

1. **In-app (recommended):** log in as admin → **Settings** → Company tab. Covers name, address, phone, email, website, TRN, logo, work hours, weekend day, timezone, and date format. This is what Letters and PDF/Excel exports actually pull from (`CompanySettings` model).
2. **settings.py** (fallback values only, used before `CompanySettings` is first saved):
   ```python
   COMPANY_NAME = "Your Company Name"
   COMPANY_ADDRESS = "Your Address, Dubai, UAE"
   COMPANY_PHONE = "+971 XX XXX XXXX"
   COMPANY_EMAIL = "info@yourcompany.com"
   ```

### AI Assistant API Key
Set in `settings.py`:
```python
OPENROUTER_API_KEY = 'your-key-here'   # https://openrouter.ai/keys (free tier available)
```
Without a key, the AI Chat page will show a setup message but the rest of the app works normally.

---

## PDF / Excel Export Layout
- Multi-employee attendance export: configurable employees per page, dates as rows, IN / OUT / OT columns, **ABSENT in red**, holidays highlighted
- Letters: formal letterhead (HR-issued) or clean Ref/Date header (employee-request), bordered details box, signature block — identical content in both DOCX and PDF
- Temporary Attendance: detail rows + auto-calculated per-worker summary + grand total

---

## User Permissions

Each user is assigned individual permissions under **Users → Edit**. The sidebar
only shows modules a user has access to. Superusers bypass all checks.

| Module | Permission Codes | Covers |
|---|---|---|
| Dashboard | `dashboard_view` | Main dashboard |
| Employees | `employee_view`, `employee_detail`, `employee_add`, `employee_edit`, `employee_delete` | Employee CRUD |
| Attendance | `attendance_view`, `attendance_add`, `attendance_edit` | Daily attendance marking/editing |
| Holidays | `holiday_view`, `holiday_add`, `holiday_edit` | Holiday calendar |
| Leaves | `leave_view`, `leave_manage` | Leave requests, approve/reject |
| Export | `export_view` | PDF/Excel exports, reports |
| Activity | `activity_view` | Activity log |
| Users | `user_manage` | User accounts & permissions (admin only) |
| Letters | `letter_view`, `letter_request`, `letter_hr` | View letters; generate employee-request letters; generate HR-issued letters (incl. Warning) |
| Settlements | `settlement_view`, `settlement_manage` | View / add-edit-delete settlement records |
| Temporary Attendance | `temp_attendance_view`, `temp_attendance_manage` | View / manage temporary workers & their attendance |

---

## Notes on Data Model Highlights

- `Employee.current_salary` — kept in sync automatically when an HR Hike
  Approval letter is generated with "update employee record" checked.
- `GeneratedLetter` — one row per letter ever created, storing the structured
  input data (not a static file) so edits regenerate a fresh, always-accurate
  DOCX/PDF on download.
- `TemporaryEmployee` / `TemporaryAttendance` — intentionally separate from
  `Employee` / `Attendance`, since outsource/hourly workers don't need
  Emirates ID, documents, or monthly-salary tracking.
- `Settlement` — plain entry/history record; all figures are entered
  manually by design (no automatic OT or deduction calculation).
