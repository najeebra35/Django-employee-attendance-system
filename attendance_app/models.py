from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import datetime


class Employee(models.Model):
    TYPE_CHOICES = [('permanent', 'Permanent'), ('temporary', 'Temporary')]
    STATUS_CHOICES = [('active', 'Active'), ('disabled', 'Disabled')]

    name = models.CharField(max_length=200)
    emirates_id = models.CharField(max_length=50, blank=True, null=True)
    mobile = models.CharField(max_length=20, blank=True, null=True)
    dob = models.DateField(blank=True, null=True)
    joining_date = models.DateField(blank=True, null=True)
    job_title = models.CharField(max_length=100, blank=True, null=True)
    country = models.CharField(max_length=100, blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    description = models.TextField(blank=True, null=True)
    emp_type = models.CharField(max_length=20, choices=TYPE_CHOICES, default='permanent')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='active')
    photo = models.ImageField(upload_to='employee_photos/', blank=True, null=True)
    portal_user = models.OneToOneField('auth.User', on_delete=models.SET_NULL, null=True, blank=True, related_name='employee_profile')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

    class Meta:
        ordering = ['name']


class Holiday(models.Model):
    TYPE_CHOICES = [('public', 'Public Holiday'), ('sunday', 'Sunday')]
    date = models.DateField(unique=True)
    name = models.CharField(max_length=200)
    holiday_type = models.CharField(max_length=20, choices=TYPE_CHOICES, default='public')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} - {self.date}"

    class Meta:
        ordering = ['date']


class Attendance(models.Model):
    STATUS_CHOICES = [
        ('present', 'Present'),
        ('absent', 'Absent'),
        ('half_day', 'Half Day'),
        ('holiday', 'Holiday'),
        ('leave', 'On Leave'),
    ]
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='attendances')
    date = models.DateField()
    in_time = models.TimeField(blank=True, null=True)
    out_time = models.TimeField(blank=True, null=True)
    ot_hours = models.DecimalField(max_digits=4, decimal_places=2, default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='present')
    notes = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    def total_hours(self):
        if self.in_time and self.out_time:
            dt_in = datetime.datetime.combine(self.date, self.in_time)
            dt_out = datetime.datetime.combine(self.date, self.out_time)
            if dt_out < dt_in:
                dt_out += datetime.timedelta(days=1)
            delta = dt_out - dt_in
            return round(delta.seconds / 3600, 2)
        return 0

    def __str__(self):
        return f"{self.employee.name} - {self.date}"

    class Meta:
        unique_together = ['employee', 'date']
        ordering = ['-date']


class LeaveType(models.Model):
    name = models.CharField(max_length=100)
    days_allowed = models.IntegerField(default=0)
    description = models.TextField(blank=True, null=True)

    def __str__(self):
        return self.name


class LeaveRequest(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('approved', 'Approved'),
        ('rejected', 'Rejected'),
    ]
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='leaves')
    leave_type = models.ForeignKey(LeaveType, on_delete=models.CASCADE)
    start_date = models.DateField()
    end_date = models.DateField()
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    applied_on = models.DateTimeField(auto_now_add=True)
    reviewed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    review_note = models.TextField(blank=True, null=True)

    def total_days(self):
        return (self.end_date - self.start_date).days + 1

    def __str__(self):
        return f"{self.employee.name} - {self.leave_type.name} ({self.start_date})"

    class Meta:
        ordering = ['-applied_on']



class UserPermission(models.Model):
    PERMISSION_CHOICES = [
        ('dashboard_view', 'View Dashboard'),
        ('employee_view', 'View Employee List'),
        ('employee_detail', 'View Employee Detail'),
        ('employee_add', 'Add Employee'),
        ('employee_edit', 'Edit Employee'),
        ('employee_delete', 'Delete Employee'),
        ('attendance_view', 'View Attendance'),
        ('attendance_add', 'Mark Attendance'),
        ('attendance_edit', 'Edit Attendance'),
        ('holiday_view', 'View Holidays'),
        ('holiday_add', 'Add Holiday'),
        ('holiday_edit', 'Edit Holiday'),
        ('leave_view', 'View Leave Requests'),
        ('leave_manage', 'Manage Leave Requests'),
        ('export_view', 'Export Attendance'),
        ('activity_view', 'View Activity Log'),
        ('user_manage', 'Manage Users (Admin)'),
    ]
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='custom_permissions')
    permissions = models.JSONField(default=list)

    def has_perm(self, perm):
        if self.user.is_superuser:
            return True
        return perm in self.permissions

    def __str__(self):
        return f"Permissions for {self.user.username}"


class ActivityLog(models.Model):
    ACTION_CHOICES = [
        ('create', 'Created'),
        ('update', 'Updated'),
        ('delete', 'Deleted'),
        ('login', 'Logged In'),
        ('logout', 'Logged Out'),
        ('export', 'Exported'),
        ('view', 'Viewed'),
    ]
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    model_name = models.CharField(max_length=100)
    object_id = models.IntegerField(null=True, blank=True)
    object_repr = models.CharField(max_length=300, blank=True)
    description = models.TextField(blank=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user} {self.action} {self.model_name} at {self.timestamp}"

    class Meta:
        ordering = ['-timestamp']


class CompanySettings(models.Model):
    """Single-row settings table — always use CompanySettings.get_settings()"""
    company_name    = models.CharField(max_length=200, default='Your Company Name')
    company_address = models.TextField(default='Your Address, Dubai, UAE')
    company_phone   = models.CharField(max_length=50, blank=True, default='')
    company_email   = models.EmailField(blank=True, default='')
    company_website = models.URLField(blank=True, default='')
    company_trn     = models.CharField(max_length=50, blank=True, default='', verbose_name='TRN / Tax No.')
    logo            = models.ImageField(upload_to='company/', blank=True, null=True)

    # Work defaults
    default_in_time  = models.TimeField(default=datetime.time(7, 0))
    default_out_time = models.TimeField(default=datetime.time(17, 0))
    work_days        = models.CharField(
        max_length=20, default='Mon-Sat',
        help_text='e.g. Mon-Sat, Mon-Fri'
    )
    weekend_day      = models.CharField(
        max_length=20, default='Sunday',
        choices=[('Sunday','Sunday'),('Friday','Friday'),('Saturday','Saturday'),('Friday-Saturday','Fri & Sat')],
    )

    # System
    timezone_name    = models.CharField(max_length=60, default='Asia/Dubai')
    date_format      = models.CharField(
        max_length=20, default='d-m-Y',
        choices=[('d-m-Y','DD-MM-YYYY'),('m/d/Y','MM/DD/YYYY'),('Y-m-d','YYYY-MM-DD')],
    )
    updated_at       = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Company Settings'

    def __str__(self):
        return self.company_name

    @classmethod
    def get_settings(cls):
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class DocumentType(models.Model):
    """Admin-defined document categories."""
    name        = models.CharField(max_length=100)          # e.g. Passport, Visa, Emirates ID
    description = models.TextField(blank=True, null=True)
    requires_expiry = models.BooleanField(default=True)
    alert_days  = models.IntegerField(default=30,
        help_text='Send alert N days before expiry')
    created_at  = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name

    class Meta:
        ordering = ['name']


class EmployeeDocument(models.Model):
    """A single document file attached to an employee."""
    STATUS_CHOICES = [
        ('valid',    'Valid'),
        ('expiring', 'Expiring Soon'),
        ('expired',  'Expired'),
    ]
    employee    = models.ForeignKey(Employee, on_delete=models.CASCADE,
                                    related_name='documents')
    doc_type    = models.ForeignKey(DocumentType, on_delete=models.PROTECT,
                                    verbose_name='Document Type')
    doc_number  = models.CharField(max_length=100, blank=True, null=True,
                                   verbose_name='Document / Reference Number')
    issue_date  = models.DateField(blank=True, null=True)
    expiry_date = models.DateField(blank=True, null=True)
    file        = models.FileField(upload_to='employee_docs/%Y/', blank=True, null=True)
    notes       = models.TextField(blank=True, null=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL,
                                    null=True, blank=True)
    created_at  = models.DateTimeField(auto_now_add=True)
    updated_at  = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.employee.name} — {self.doc_type.name}"

    @property
    def status(self):
        if not self.expiry_date or not self.doc_type.requires_expiry:
            return 'valid'
        today = datetime.date.today()
        if self.expiry_date < today:
            return 'expired'
        delta = (self.expiry_date - today).days
        if delta <= self.doc_type.alert_days:
            return 'expiring'
        return 'valid'

    @property
    def days_until_expiry(self):
        if not self.expiry_date:
            return None
        return (self.expiry_date - datetime.date.today()).days

    @property
    def file_extension(self):
        if self.file:
            name = self.file.name.lower()
            if name.endswith('.pdf'): return 'pdf'
            if name.endswith(('.jpg','.jpeg','.png','.gif','.webp')): return 'image'
            if name.endswith(('.doc','.docx')): return 'word'
            if name.endswith(('.xls','.xlsx')): return 'excel'
        return 'file'

    class Meta:
        ordering = ['expiry_date']
