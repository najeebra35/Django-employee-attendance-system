#letter_views.py
import datetime
from collections import defaultdict

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse, HttpResponse
from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone

from .models import Employee, LeaveType, CompanySettings, GeneratedLetter
from .middleware import log_activity
from .views import has_perm
from . import letters as L


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _next_ref_no(prefix):
    today = timezone.localdate()
    seq = GeneratedLetter.objects.count() + 1
    return f"{prefix}/{today.year}/{seq:04d}"


def _perm_for_category(category):
    return 'letter_hr' if category == 'hr' else 'letter_request'


def _can_access(request, category):
    return has_perm(request.user, _perm_for_category(category))


def _log(request, employee, category, letter_type, action, ref_no=''):
    log_activity(
        request.user, action, 'GeneratedLetter',
        description=f'{action.title()}d {letter_type.replace("_"," ").title()} for {employee.name}',
        request=request,
    )


def _respond_letter(request, ctx, action, filename_base):
    if action == 'pdf':
        data = L.build_letter_pdf(ctx)
        resp = HttpResponse(data, content_type='application/pdf')
        resp['Content-Disposition'] = f'attachment; filename="{filename_base}.pdf"'
        return resp
    else:
        data = L.build_letter_docx(ctx)
        resp = HttpResponse(
            data,
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        resp['Content-Disposition'] = f'attachment; filename="{filename_base}.docx"'
        return resp


def _safe_name(name):
    return "_".join(name.strip().split()).lower()


def _f(request, key, default=None):
    v = request.POST.get(key, '')
    if v == '':
        return default
    try:
        return float(v)
    except ValueError:
        return default


def _d(request, key):
    v = request.POST.get(key, '').strip()
    try:
        return datetime.date.fromisoformat(v)
    except ValueError:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Hub
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def letter_hub(request):
    if not has_perm(request.user, 'letter_view'):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    recent = GeneratedLetter.objects.select_related('employee', 'generated_by')[:8]
    return render(request, 'attendance_app/letter_hub.html', {
        'recent': recent,
        'can_request': has_perm(request.user, 'letter_request'),
        'can_hr': has_perm(request.user, 'letter_hr'),
        'preselect_emp': request.GET.get('employee', ''),
    })


@login_required
def letter_list(request):
    if not has_perm(request.user, 'letter_view'):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    logs = GeneratedLetter.objects.select_related('employee', 'generated_by', 'updated_by').all()
    category = request.GET.get('category', '')
    letter_type = request.GET.get('type', '')
    emp_id = request.GET.get('employee', '')
    q = request.GET.get('q', '').strip()

    if category:
        logs = logs.filter(category=category)
    if letter_type:
        logs = logs.filter(letter_type=letter_type)
    if emp_id:
        logs = logs.filter(employee_id=emp_id)
    if q:
        logs = logs.filter(employee__name__icontains=q)

    return render(request, 'attendance_app/letter_list.html', {
        'logs': logs[:300],
        'category': category, 'letter_type': letter_type, 'emp_id': emp_id, 'q': q,
        'employees': Employee.objects.all().order_by('name'),
        'letter_types': GeneratedLetter.LETTER_TYPE_CHOICES,
    })


# Alias so any old links to letter_history keep working
letter_history = letter_list


EDIT_URL_NAMES = {
    'salary_hike_request': 'letter_request_salary_edit',
    'leave_request': 'letter_request_leave_edit',
    'hike_approval': 'letter_hr_hike_approval_edit',
    'experience': 'letter_hr_experience_edit',
    'noc': 'letter_hr_noc_edit',
    'warning': 'letter_hr_warning_edit',
}


@login_required
def letter_detail(request, pk):
    rec = get_object_or_404(GeneratedLetter, pk=pk)
    if not has_perm(request.user, 'letter_view'):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    cs = CompanySettings.get_settings()
    try:
        ctx = L.build_ctx_for_record(rec.letter_type, rec.employee, cs, rec.details)
    except Exception:
        ctx = None

    return render(request, 'attendance_app/letter_detail.html', {
        'rec': rec, 'ctx': ctx,
        'can_edit': _can_access(request, rec.category),
        'edit_url_name': EDIT_URL_NAMES.get(rec.letter_type),
    })


@login_required
def letter_download(request, pk, filetype):
    rec = get_object_or_404(GeneratedLetter, pk=pk)
    if not has_perm(request.user, 'letter_view'):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    cs = CompanySettings.get_settings()
    ctx = L.build_ctx_for_record(rec.letter_type, rec.employee, cs, rec.details)
    fname = f"{rec.letter_type}_{_safe_name(rec.employee.name)}_{rec.created_at.date()}"
    return _respond_letter(request, ctx, filetype, fname)


@login_required
def letter_delete(request, pk):
    rec = get_object_or_404(GeneratedLetter, pk=pk)
    if not _can_access(request, rec.category):
        messages.error(request, 'Access denied.')
        return redirect('letter_list')

    if request.method == 'POST':
        emp_id = rec.employee_id
        name = rec.employee.name
        letter_type = rec.letter_type
        return_to = request.POST.get('return_to', '')
        rec.delete()
        log_activity(request.user, 'delete', 'GeneratedLetter',
                     description=f'Deleted {letter_type.replace("_"," ").title()} for {name}', request=request)
        messages.success(request, 'Letter deleted.')
        if return_to == 'employee':
            return redirect(f'/employees/{emp_id}/?tab=letters')
    return redirect('letter_list')


@login_required
def letter_employee_info(request):
    """AJAX — return employee details for autofill."""
    emp_id = request.GET.get('employee_id')
    try:
        emp = Employee.objects.get(pk=emp_id)
    except (Employee.DoesNotExist, ValueError, TypeError):
        return JsonResponse({'error': 'Not found'}, status=404)

    return JsonResponse({
        'name': emp.name,
        'job_title': emp.job_title or '',
        'emirates_id': emp.emirates_id or '',
        'joining_date': emp.joining_date.isoformat() if emp.joining_date else '',
        'current_salary': str(emp.current_salary) if emp.current_salary is not None else '',
        'mobile': emp.mobile or '',
    })


# ─────────────────────────────────────────────────────────────────────────────
# EMPLOYEE REQUEST LETTERS
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def letter_request_salary(request, pk=None):
    if not has_perm(request.user, 'letter_request'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.filter(status='active').order_by('name')
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='salary_hike_request') if pk else None

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        current_salary = _f(request, 'current_salary')
        requested_salary = _f(request, 'requested_salary')
        extra_note = request.POST.get('extra_note', '').strip()
        letter_date = _d(request, 'letter_date') or timezone.localdate()

        if not current_salary:
            messages.error(request, 'Current salary is required.')
            return render(request, 'attendance_app/letter_request_salary.html', {
                'employees': employees, 'form_data': request.POST, 'rec': rec, 'initial': defaultdict(str), 'preselect_emp': '',
            })

        if request.POST.get('save_salary') == '1':
            emp.current_salary = current_salary
            emp.save(update_fields=['current_salary'])

        ref_no = rec.reference_no if rec else _next_ref_no('SHR')
        details = {
            'current_salary': current_salary, 'requested_salary': requested_salary,
            'extra_note': extra_note, 'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'request', 'salary_hike_request', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='request', letter_type='salary_hike_request',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'request', 'salary_hike_request', 'create', ref_no)

        ctx = L.letter_ctx_salary_hike_request(
            emp, cs, current_salary=current_salary, requested_salary=requested_salary,
            extra_note=extra_note, ref_no=ref_no, today=letter_date,
        )

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"salary_hike_request_{_safe_name(emp.name)}_{letter_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_request_salary.html', {
        'employees': employees, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })


@login_required
def letter_request_leave(request, pk=None):
    if not has_perm(request.user, 'letter_request'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.filter(status='active').order_by('name')
    leave_types = LeaveType.objects.all()
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='leave_request') if pk else None

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        leave_type_id = request.POST.get('leave_type')
        try:
            leave_type_name = LeaveType.objects.get(pk=leave_type_id).name
        except (LeaveType.DoesNotExist, ValueError, TypeError):
            leave_type_name = request.POST.get('leave_type_custom', 'Leave')

        start_date = _d(request, 'start_date')
        end_date = _d(request, 'end_date')
        reason = request.POST.get('reason', '').strip()
        letter_date = _d(request, 'letter_date') or timezone.localdate()

        if not start_date or not end_date or end_date < start_date:
            messages.error(request, 'Please provide a valid date range.')
            return render(request, 'attendance_app/letter_request_leave.html', {
                'employees': employees, 'leave_types': leave_types, 'form_data': request.POST, 'rec': rec, 'initial': defaultdict(str), 'preselect_emp': '',
            })

        total_days = (end_date - start_date).days + 1
        ref_no = rec.reference_no if rec else _next_ref_no('LEV')

        details = {
            'leave_type': leave_type_name, 'start_date': str(start_date), 'end_date': str(end_date),
            'total_days': total_days, 'reason': reason, 'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'request', 'leave_request', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='request', letter_type='leave_request',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'request', 'leave_request', 'create', ref_no)

        ctx = L.letter_ctx_leave_request(
            emp, cs, leave_type_name=leave_type_name, start_date=start_date, end_date=end_date,
            total_days=total_days, reason=reason, ref_no=ref_no, today=letter_date,
        )

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"leave_request_{_safe_name(emp.name)}_{start_date}_{end_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_request_leave.html', {
        'employees': employees, 'leave_types': leave_types, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })


# ─────────────────────────────────────────────────────────────────────────────
# HR ISSUED LETTERS
# ─────────────────────────────────────────────────────────────────────────────

@login_required
def letter_hr_hike_approval(request, pk=None):
    if not has_perm(request.user, 'letter_hr'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.filter(status='active').order_by('name')
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='hike_approval') if pk else None

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        last_salary = _f(request, 'last_salary', 0) or 0
        new_salary = _f(request, 'new_salary', 0) or 0
        effective_date = _d(request, 'effective_date') or timezone.localdate()
        letter_date = _d(request, 'letter_date') or timezone.localdate()

        if last_salary <= 0 or new_salary <= 0:
            messages.error(request, 'Both last salary and new salary are required.')
            return render(request, 'attendance_app/letter_hr_hike_approval.html', {
                'employees': employees, 'form_data': request.POST, 'rec': rec, 'initial': defaultdict(str), 'preselect_emp': '',
            })

        hike_percent = ((new_salary - last_salary) / last_salary) * 100 if last_salary else 0
        ref_no = rec.reference_no if rec else _next_ref_no('INC')

        details = {
            'last_salary': last_salary, 'new_salary': new_salary,
            'hike_percent': round(hike_percent, 2), 'effective_date': str(effective_date),
            'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if request.POST.get('update_record') == '1':
            emp.current_salary = new_salary
            emp.save(update_fields=['current_salary'])

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'hr', 'hike_approval', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='hr', letter_type='hike_approval',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'hr', 'hike_approval', 'create', ref_no)

        ctx = L.letter_ctx_hike_approval(
            emp, cs, last_salary=last_salary, new_salary=new_salary, hike_percent=hike_percent,
            effective_date=effective_date, ref_no=ref_no, today=letter_date,
        )

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"hike_approval_{_safe_name(emp.name)}_{letter_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_hr_hike_approval.html', {
        'employees': employees, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })


@login_required
def letter_hr_experience(request, pk=None):
    if not has_perm(request.user, 'letter_hr'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.all().order_by('name')
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='experience') if pk else None

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        still_employed = request.POST.get('still_employed') == '1'
        last_working_date = None
        if not still_employed:
            last_working_date = _d(request, 'last_working_date')
            if not last_working_date:
                messages.error(request, 'Please provide the last working date.')
                return render(request, 'attendance_app/letter_hr_experience.html', {
                    'employees': employees, 'form_data': request.POST, 'rec': rec, 'initial': defaultdict(str), 'preselect_emp': '',
                })

        job_title_override = request.POST.get('job_title_override', '').strip()
        letter_date = _d(request, 'letter_date') or timezone.localdate()
        ref_no = rec.reference_no if rec else _next_ref_no('EXP')

        details = {
            'still_employed': still_employed,
            'last_working_date': str(last_working_date) if last_working_date else '',
            'job_title_override': job_title_override, 'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'hr', 'experience', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='hr', letter_type='experience',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'hr', 'experience', 'create', ref_no)

        ctx = L.letter_ctx_experience(
            emp, cs, last_working_date=last_working_date, still_employed=still_employed,
            job_title_override=job_title_override, ref_no=ref_no, today=letter_date,
        )

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"experience_letter_{_safe_name(emp.name)}_{letter_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_hr_experience.html', {
        'employees': employees, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })


@login_required
def letter_hr_noc(request, pk=None):
    if not has_perm(request.user, 'letter_hr'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.filter(status='active').order_by('name')
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='noc') if pk else None

    NOC_PURPOSES = [
        'Visa Processing', 'Bank Loan Application', 'Driving License Application',
        'Family Visa Sponsorship', 'Other Employment / Job Application', 'Other',
    ]

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        purpose = request.POST.get('purpose', '').strip()
        if purpose == 'Other':
            purpose = request.POST.get('purpose_custom', '').strip() or 'the stated purpose'
        addressed_to = request.POST.get('addressed_to', '').strip()
        letter_date = _d(request, 'letter_date') or timezone.localdate()
        ref_no = rec.reference_no if rec else _next_ref_no('NOC')

        details = {
            'purpose': purpose, 'addressed_to': addressed_to,
            'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'hr', 'noc', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='hr', letter_type='noc',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'hr', 'noc', 'create', ref_no)

        ctx = L.letter_ctx_noc(emp, cs, purpose=purpose, addressed_to=addressed_to, ref_no=ref_no, today=letter_date)

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"noc_letter_{_safe_name(emp.name)}_{letter_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_hr_noc.html', {
        'employees': employees, 'purposes': NOC_PURPOSES, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })


@login_required
def letter_hr_warning(request, pk=None):
    if not has_perm(request.user, 'letter_hr'):
        messages.error(request, 'Access denied.')
        return redirect('letter_hub')

    employees = Employee.objects.filter(status='active').order_by('name')
    cs = CompanySettings.get_settings()
    rec = get_object_or_404(GeneratedLetter, pk=pk, letter_type='warning') if pk else None

    VIOLATION_TYPES = [
        'Late Attendance', 'Absenteeism', 'Misconduct', 'Safety Violation',
        'Policy Violation', 'Insubordination', 'Poor Performance', 'Other',
    ]
    WARNING_LEVELS = ['First Warning', 'Second Warning', 'Final Warning']

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        action = request.POST.get('action', 'docx')

        violation_type = request.POST.get('violation_type', '').strip()
        if violation_type == 'Other':
            violation_type = request.POST.get('violation_type_custom', '').strip() or 'Policy Violation'
        warning_level = request.POST.get('warning_level', 'First Warning')
        incident_date = _d(request, 'incident_date')
        description = request.POST.get('description', '').strip()
        corrective_action = request.POST.get('corrective_action', '').strip()
        letter_date = _d(request, 'letter_date') or timezone.localdate()

        if not violation_type or not incident_date:
            messages.error(request, 'Violation type and incident date are required.')
            return render(request, 'attendance_app/letter_hr_warning.html', {
                'employees': employees, 'form_data': request.POST, 'rec': rec, 'initial': defaultdict(str),
                'violation_types': VIOLATION_TYPES, 'warning_levels': WARNING_LEVELS, 'preselect_emp': '',
            })

        ref_no = rec.reference_no if rec else _next_ref_no('WRN')
        details = {
            'violation_type': violation_type, 'warning_level': warning_level,
            'incident_date': str(incident_date), 'description': description,
            'corrective_action': corrective_action, 'letter_date': str(letter_date), 'ref_no': ref_no,
        }

        if rec:
            rec.employee = emp
            rec.details = details
            rec.updated_by = request.user
            rec.save()
            _log(request, emp, 'hr', 'warning', 'update', ref_no)
        else:
            rec = GeneratedLetter.objects.create(
                employee=emp, category='hr', letter_type='warning',
                reference_no=ref_no, details=details, generated_by=request.user,
            )
            _log(request, emp, 'hr', 'warning', 'create', ref_no)

        ctx = L.letter_ctx_warning(
            emp, cs, violation_type=violation_type, incident_date=incident_date,
            warning_level=warning_level, description=description,
            corrective_action=corrective_action, ref_no=ref_no, today=letter_date,
        )

        if action == 'save':
            messages.success(request, 'Letter saved.')
            return redirect('letter_detail', pk=rec.pk)

        fname = f"warning_letter_{_safe_name(emp.name)}_{letter_date}"
        return _respond_letter(request, ctx, action, fname)

    initial = defaultdict(str, rec.details) if rec else defaultdict(str)
    if rec:
        initial['employee'] = rec.employee_id
        initial.setdefault('letter_date', str(timezone.localdate()))
    return render(request, 'attendance_app/letter_hr_warning.html', {
        'employees': employees, 'rec': rec, 'initial': initial, 'form_data': defaultdict(str),
        'violation_types': VIOLATION_TYPES, 'warning_levels': WARNING_LEVELS,
        'preselect_emp': request.GET.get('employee', '') or (str(rec.employee_id) if rec else ''),
    })
