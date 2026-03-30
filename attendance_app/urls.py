from django.urls import path
from . import views, ai_views

urlpatterns = [
    # Auth
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),

    # Dashboard
    path('', views.dashboard, name='dashboard'),

    # Employees
    path('employees/', views.employee_list, name='employee_list'),
    path('employees/add/', views.employee_add, name='employee_add'),
    path('employees/<int:pk>/', views.employee_detail, name='employee_detail'),
    path('employees/<int:pk>/edit/', views.employee_edit, name='employee_edit'),
    path('employees/<int:pk>/delete/', views.employee_delete, name='employee_delete'),
    path('employees/<int:pk>/calendar-data/', views.employee_calendar_data, name='employee_calendar_data'),

    # Attendance
    path('attendance/', views.attendance_list, name='attendance_list'),
    path('attendance/mark/', views.attendance_mark, name='attendance_mark'),
    path('attendance/save/', views.attendance_save, name='attendance_save'),
    path('attendance/edit/<int:pk>/', views.attendance_edit, name='attendance_edit'),

    # Holidays
    path('holidays/', views.holiday_list, name='holiday_list'),
    path('holidays/add/', views.holiday_add, name='holiday_add'),
    path('holidays/delete/<int:pk>/', views.holiday_delete, name='holiday_delete'),
    path('holidays/generate-sundays/', views.generate_sundays, name='generate_sundays'),

    # Leaves
    path('leaves/', views.leave_list, name='leave_list'),
    path('leaves/add/', views.leave_add, name='leave_add'),
    path('leaves/<int:pk>/approve/', views.leave_approve, name='leave_approve'),
    path('leaves/<int:pk>/reject/', views.leave_reject, name='leave_reject'),
    path('leave-types/', views.leave_type_list, name='leave_type_list'),
    path('leave-types/add/', views.leave_type_add, name='leave_type_add'),

    # Export
    path('export/', views.export_attendance, name='export_attendance'),
    path('export/employees/', views.export_get_employees, name='export_get_employees'),
    path('export/pdf/', views.export_pdf, name='export_pdf'),
    path('export/excel/', views.export_excel, name='export_excel'),

    # Users
    path('users/', views.user_list, name='user_list'),
    path('users/add/', views.user_add, name='user_add'),
    path('users/<int:pk>/edit/', views.user_edit, name='user_edit'),
    path('users/<int:pk>/delete/', views.user_delete, name='user_delete'),

    # Employee Self-Service Portal
    path('portal/', views.portal_dashboard, name='portal_dashboard'),
    path('portal/attendance/', views.portal_attendance, name='portal_attendance'),
    path('portal/attendance/ajax/', views.portal_attendance_ajax, name='portal_attendance_ajax'),

    # Reports
    path('reports/', views.reports_view, name='reports'),
    path('reports/monthly/', views.report_monthly, name='report_monthly'),
    path('reports/monthly/export/', views.report_monthly_export, name='report_monthly_export'),
    path('reports/ot/', views.report_ot, name='report_ot'),
    path('reports/ot/export/', views.report_ot_export, name='report_ot_export'),
    path('reports/absent/', views.report_absent, name='report_absent'),
    path('reports/absent/export/', views.report_absent_export, name='report_absent_export'),
    path('reports/late/', views.report_late, name='report_late'),
    path('reports/late/export/', views.report_late_export, name='report_late_export'),

    # Documents
    path('documents/', views.document_list, name='document_list'),
    path('documents/expiring/', views.document_expiring, name='document_expiring'),
    path('documents/add/', views.document_add, name='document_add'),
    path('documents/<int:pk>/edit/', views.document_edit, name='document_edit'),
    path('documents/<int:pk>/delete/', views.document_delete, name='document_delete'),
    path('documents/<int:pk>/download/', views.document_download, name='document_download'),
    path('document-types/', views.doctype_list, name='doctype_list'),
    path('document-types/add/', views.doctype_add, name='doctype_add'),
    path('document-types/<int:pk>/edit/', views.doctype_edit, name='doctype_edit'),
    path('document-types/<int:pk>/delete/', views.doctype_delete, name='doctype_delete'),

    # Bulk Import
    path('import/', views.import_hub, name='import_hub'),
    path('import/employees/', views.import_employees, name='import_employees'),
    path('import/employees/template/', views.import_employees_template, name='import_employees_template'),
    path('import/attendance/', views.import_attendance, name='import_attendance'),
    path('import/attendance/template/', views.import_attendance_template, name='import_attendance_template'),

    # Settings
    path('settings/', views.settings_view, name='settings'),

    # Activity Log
    path('activity/', views.activity_log, name='activity_log'),

    path('ai/', ai_views.ai_chat_page, name='ai_chat'),
    path('ai/chat/', ai_views.ai_chat_message, name='ai_chat_message'),
    path('ai/export/', ai_views.ai_chat_export, name='ai_chat_export'),
]
