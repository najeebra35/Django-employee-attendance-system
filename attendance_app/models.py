#models.py
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
    current_salary = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True,
                                          help_text='Current monthly salary (AED) — used on salary/letter documents')

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
        ('letter_view', 'View Letters'),
        ('letter_request', 'Generate Employee Request Letters'),
        ('letter_hr', 'Generate HR Letters'),
        ('settlement_view', 'View Settlements'),
        ('settlement_manage', 'Add/Edit Settlements'),
        ('temp_attendance_view', 'View Temporary Attendance'),
        ('temp_attendance_manage', 'Manage Temporary Attendance'),
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


class GeneratedLetter(models.Model):
    """Log of every HR / employee request letter generated, for history & audit."""
    CATEGORY_CHOICES = [
        ('request', 'Employee Request Letter'),
        ('hr', 'HR Issued Letter'),
    ]
    LETTER_TYPE_CHOICES = [
        ('salary_hike_request', 'Salary Hike Request Letter'),
        ('leave_request', 'Leave Request Letter'),
        ('hike_approval', 'Salary Hike Approval Letter'),
        ('experience', 'Experience Letter'),
        ('noc', 'No Objection Certificate (NOC)'),
        ('warning', 'Warning / Disciplinary Letter'),
    ]
    employee     = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='letters')
    category     = models.CharField(max_length=20, choices=CATEGORY_CHOICES)
    letter_type  = models.CharField(max_length=30, choices=LETTER_TYPE_CHOICES)
    reference_no = models.CharField(max_length=50, blank=True, default='')
    details      = models.JSONField(default=dict, blank=True)
    generated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='letters_created')
    updated_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='letters_updated')
    created_at   = models.DateTimeField(auto_now_add=True)
    updated_at   = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.get_letter_type_display()} — {self.employee.name}"

    class Meta:
        ordering = ['-created_at']


class Settlement(models.Model):
    """
    Manual OT / leave-deduction settlement record. Entries are made periodically
    (e.g. every 2-4 months) to settle OT payments and any leave/absence deductions
    for a given period. No automatic calculation — figures are entered manually
    by HR/admin based on their own calculation, this model is purely for entry
    and history/audit purposes.
    """
    employee               = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='settlements')
    from_date               = models.DateField(verbose_name='Period From')
    to_date                 = models.DateField(verbose_name='Period To')
    ot_amount                = models.DecimalField(max_digits=10, decimal_places=2, default=0,
                                                    verbose_name='OT Amount (AED)')
    leave_days               = models.DecimalField(max_digits=5, decimal_places=1, default=0, blank=True,
                                                    verbose_name='Total Leave/Absent Days')
    leave_deduction_amount   = models.DecimalField(max_digits=10, decimal_places=2, default=0, blank=True,
                                                    verbose_name='Leave Deduction Amount (AED)')
    net_amount               = models.DecimalField(max_digits=10, decimal_places=2,
                                                    verbose_name='Net Settlement Amount (AED)')
    settlement_date          = models.DateField(default=timezone.now, verbose_name='Settlement Date')
    notes                     = models.TextField(blank=True, null=True)
    created_by                = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                                   related_name='settlements_created')
    updated_by                = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                                   related_name='settlements_updated')
    created_at                = models.DateTimeField(auto_now_add=True)
    updated_at                = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Settlement — {self.employee.name} ({self.from_date} to {self.to_date})"

    class Meta:
        ordering = ['-settlement_date', '-created_at']


class TemporaryEmployee(models.Model):
    """
    Outsource / ad-hoc hourly worker — lightweight record, separate from the
    main Employee model. Used for casual/urgent labour brought in for short
    jobs, tracked by hours worked rather than monthly salary + in/out time.
    """
    STATUS_CHOICES = [('active', 'Active'), ('inactive', 'Inactive')]

    name         = models.CharField(max_length=200)
    designation  = models.CharField(max_length=100, blank=True, null=True)
    mobile       = models.CharField(max_length=20, blank=True, null=True)
    hourly_rate  = models.DecimalField(max_digits=8, decimal_places=2, blank=True, null=True,
                                        help_text='Optional — used to calculate payable amount on export')
    status       = models.CharField(max_length=20, choices=STATUS_CHOICES, default='active')
    notes        = models.TextField(blank=True, null=True)
    created_by   = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    created_at   = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name

    class Meta:
        ordering = ['name']


class TemporaryAttendance(models.Model):
    """
    A single day's worked hours for a TemporaryEmployee. Only created for days
    the person actually worked (present) — hours are entered manually since
    breaks etc. mean simple in/out time subtraction wouldn't be accurate.
    Total hours across a period are summed automatically on export.
    """
    temp_employee = models.ForeignKey(TemporaryEmployee, on_delete=models.CASCADE, related_name='attendances')
    date          = models.DateField()
    hours_worked  = models.DecimalField(max_digits=5, decimal_places=2, verbose_name='Total Working Hours')
    notes         = models.TextField(blank=True, null=True)
    created_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                       related_name='temp_attendance_created')
    updated_by    = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True,
                                       related_name='temp_attendance_updated')
    created_at    = models.DateTimeField(auto_now_add=True)
    updated_at    = models.DateTimeField(auto_now=True)

    def amount(self):
        if self.temp_employee.hourly_rate is not None:
            return round(float(self.hours_worked) * float(self.temp_employee.hourly_rate), 2)
        return None

    def __str__(self):
        return f"{self.temp_employee.name} - {self.date}"

    class Meta:
        unique_together = ['temp_employee', 'date']
        ordering = ['-date']
