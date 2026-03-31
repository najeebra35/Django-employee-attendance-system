# 👥 Employee Attendance System

A full-featured **Django web application** for managing employee bio data, documents, and attendance records — all in one place.

---

## 📋 Features

- **Employee Bio Data Management** — Add, update, and manage detailed employee profiles
- **Document Management** — Upload and store employee-related documents securely
- **Attendance Tracking** — Record and monitor daily employee attendance
- **Admin Dashboard** — Manage all records through Django's powerful admin interface
- **Search & Filter** — Quickly find employee records and attendance logs

---

## 🛠️ Tech Stack

| Layer      | Technology        |
|------------|-------------------|
| Backend    | Python, Django    |
| Database   | SQLite (default) / PostgreSQL |
| Frontend   | HTML, CSS, Bootstrap |
| Auth       | Django Auth System |

---

## 🚀 Getting Started

### Prerequisites

- Python 3.8+
- pip
- virtualenv (recommended)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/employee-attendance-system.git
   cd employee-attendance-system
   ```

2. **Create and activate a virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate        # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Apply migrations**
   ```bash
   python manage.py makemigrations
   python manage.py migrate
   ```

5. **Create a superuser**
   ```bash
   python manage.py createsuperuser
   ```

6. **Run the development server**
   ```bash
   python manage.py runserver
   ```

7. Open your browser and go to `http://127.0.0.1:8000`

---

## 📁 Project Structure

```
employee-attendance-system/
│
├── attendance/          # Attendance app
├── employees/           # Employee bio data & documents app
├── templates/           # HTML templates
├── static/              # CSS, JS, images
├── media/               # Uploaded documents
├── manage.py
└── requirements.txt
```

---

## ⚙️ Environment Variables

Create a `.env` file in the root directory and configure:

```env
SECRET_KEY=your-secret-key
DEBUG=True
ALLOWED_HOSTS=localhost,127.0.0.1
DATABASE_URL=sqlite:///db.sqlite3
```

---

## 📸 Screenshots

> _Add your project screenshots here_

---

## 🤝 Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Your Name**
- GitHub: [@your-username](https://github.com/your-username)

---

> Built with ❤️ using Django
