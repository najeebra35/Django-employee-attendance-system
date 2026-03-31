# AttendPro — Employee Attendance Management System

> A full-featured employee attendance management web application built with Django, designed for UAE/GCC businesses. Manage employees, track daily attendance, handle leave requests, store documents, and get AI-powered insights — all from one place.

---

## 📸 Features at a Glance

- ✅ **Daily Attendance Marking** — Mark present / absent / half day / leave / holiday per employee
- 👥 **Employee Management** — Full profiles with photos, Emirates ID, joining date, documents
- 📅 **Leave Management** — Apply, approve, and reject leave requests with leave type tracking
- 🗓️ **Holiday Management** — Add public holidays, auto-generate Sunday holidays for the year
- 📄 **Document Management** — Store passports, visas, Emirates IDs with expiry alerts
- 📊 **Reports & Analytics** — Monthly summary, OT report, absent frequency, late arrivals
- 📤 **Export** — Professional PDF and Excel attendance sheets with IN/OUT/OT columns
- 📥 **Bulk Import** — Import employees and attendance records from Excel templates
- 🤖 **AI Assistant** — Natural language chat to query attendance data (powered by OpenRouter/Groq/Gemini)
- 🧑‍💼 **Employee Self-Service Portal** — Employees view their own attendance calendar
- ⚙️ **Settings** — Company profile, logo, work schedule, timezone, password management
- 📱 **Fully Responsive** — Works on desktop, tablet, and mobile

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.10+, Django 4.2 |
| Database | SQLite (easy switch to PostgreSQL) |
| Frontend | HTML, CSS, Vanilla JS, Chart.js |
| PDF Export | ReportLab |
| Excel Export/Import | openpyxl |
| Static Files | WhiteNoise |
| AI Chat | OpenRouter API (free) / Groq / Gemini |
| Deployment | Gunicorn + cPanel / any Linux server |

---

## 🚀 Quick Start (Local)

### Prerequisites
- Python 3.10 or higher
- pip

### 1. Clone the repository
```bash
git clone https://github.com/najeebra35/Django-employee-attendance-system.git
cd Django-employee-attendance-system
```

### 2. Create virtual environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / Mac
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run migrations
```bash
python manage.py migrate
```

### 5. Create admin user
```bash
python manage.py createsuperuser
```

### 6. Run the server
```bash
python manage.py runserver
```

Open **http://127.0.0.1:8000** in your browser.

---

## ⚙️ Configuration

Open `attendance_system/settings.py` and update:



# AI Chat (optional — get free key from https://openrouter.ai)
OPENROUTER_API_KEY = 'sk-or-v1-xxxxxxxxxxxxxxxx'

# For production
DEBUG       = False
SECRET_KEY  = 'your-strong-random-secret-key'
ALLOWED_HOSTS = ['yourdomain.com']
```

---

## 🤖 AI Chat Setup (Optional)

The AI Assistant lets you query attendance data using natural language:
- *"Who is absent today?"*
- *"Give me Suresh's March attendance"*
- *"Which documents are expiring?"*
- *"Export this month to Excel"*

**Get a free API key:**
1. Go to **https://openrouter.ai** → Sign up
2. Create an API Key
3. Add to `settings.py`:
```python
OPENROUTER_API_KEY = 'sk-or-v1-your-key-here'
```

> Note: Gemini and Groq free tiers are blocked in UAE/Middle East. OpenRouter works worldwide.

---

## 📁 Project Structure

```
attendpro/
├── manage.py
├── requirements.txt
├── setup.bat                  # Windows quick setup
├── setup.sh                   # Linux quick setup
├── attendance_system/
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
└── attendance_app/
    ├── models.py              # Employee, Attendance, Leave, Document, etc.
    ├── views.py               # All main views
    ├── ai_views.py            # AI Chat views
    ├── urls.py
    ├── middleware.py          # Activity logging
    ├── context_processors.py
    ├── decorators.py
    ├── templatetags/
    │   └── attendance_tags.py
    ├── migrations/
    └── templates/
        └── attendance_app/
            ├── base.html
            ├── dashboard.html
            ├── employee_*.html
            ├── attendance_*.html
            ├── document_*.html
            ├── report_*.html
            ├── portal_*.html
            ├── import_*.html
            └── ai_chat.html
```

---

## 📊 Database Models

| Model | Description |
|---|---|
| `Employee` | Employee profiles with all personal details |
| `Attendance` | Daily attendance records (status, in/out time, OT) |
| `Holiday` | Public holidays and Sundays |
| `LeaveType` | Leave categories (annual, sick, etc.) |
| `LeaveRequest` | Employee leave applications |
| `DocumentType` | Document categories (Passport, Visa, Emirates ID…) |
| `EmployeeDocument` | Documents with expiry date tracking |
| `CompanySettings` | Company profile, work schedule, system settings |
| `UserPermission` | Granular permissions per staff user |
| `ActivityLog` | Full audit trail of all actions |

---

## 🔐 User Roles & Permissions

| Role | Access |
|---|---|
| **Superuser / Admin** | Full access to everything |
| **Staff User** | Configurable per-module permissions |
| **Portal User** | Employee self-service only (own attendance) |

Permissions are managed per user from **User Management → Edit User**.

---

## 📤 Export Format

The attendance PDF/Excel export produces a professional sheet with:
- Company header with logo
- Employee columns with IN / OUT / OT sub-columns
- Yellow highlighted rows for holidays and Sundays
- Red **A** for absent, Purple **V** for leave
- **Total OT** row — normal working day overtime only
- **Holiday OT** row — holiday and Sunday overtime only (different pay rate)
- **Absent / Leave** summary per employee

---

## 🌐 Deployment (Godaddy cPanel Hosting)

1. Upload project to cPanel via File Manager
2. Go to **Setup Python App** → Create Application
3. Set startup file to `passenger_wsgi.py`
4. Install packages via Terminal: `pip install -r requirements.txt`
5. Run: `python manage.py migrate && python manage.py collectstatic`
6. Restart app

See full deployment guide in `DEPLOY.md` (if included).

---

## 📋 URL Reference

| URL | Description |
|---|---|
| `/` | Dashboard |
| `/employees/` | Employee list |
| `/attendance/` | View attendance |
| `/attendance/mark/` | Mark attendance |
| `/leaves/` | Leave requests |
| `/holidays/` | Holiday management |
| `/documents/` | Document management |
| `/export/` | Export PDF / Excel |
| `/reports/` | Reports & Analytics |
| `/import/` | Bulk import |
| `/portal/` | Employee self-service portal |
| `/ai/` | AI Assistant chat |
| `/settings/` | System settings (admin only) |
| `/users/` | User management (admin only) |
| `/activity/` | Activity log |

---

## 📦 Requirements

```
Django>=4.2,<5.0
Pillow>=10.0.0
reportlab>=4.0.0
openpyxl>=3.1.0
whitenoise>=6.6.0
gunicorn>=21.2.0
```

---

## 🙏 Credits

Built with:
- [Django](https://www.djangoproject.com/) — Web framework
- [ReportLab](https://www.reportlab.com/) — PDF generation
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel export/import
- [Chart.js](https://www.chartjs.org/) — Dashboard charts
- [Font Awesome](https://fontawesome.com/) — Icons
- [Plus Jakarta Sans](https://fonts.google.com/specimen/Plus+Jakarta+Sans) — Typography
- [OpenRouter](https://openrouter.ai/) — AI API (free tier)

---

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

---

## 👤 Author

**Najeeb Rahman**
- GitHub: [@najeebra35](https://github.com/najeebra35)

---

> Made with ❤️ for UAE businesses