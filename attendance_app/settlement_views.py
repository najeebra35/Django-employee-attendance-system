#settlement_views.py
"""
Manual OT / leave-deduction settlement entries. No automatic calculation —
HR enters figures they've already worked out; this module is purely for
entry, listing, and per-employee history.
"""
import datetime
from collections import defaultdict

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone

from .models import Employee, Settlement
from .middleware import log_activity
from .views import has_perm


def _f(request, key, default=0):
    v = request.POST.get(key, '').strip()
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


@login_required
def settlement_list(request):
    if not has_perm(request.user, 'settlement_view'):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    settlements = Settlement.objects.select_related('employee', 'created_by').all()
    emp_id = request.GET.get('employee', '')
    if emp_id:
        settlements = settlements.filter(employee_id=emp_id)

    return render(request, 'attendance_app/settlement_list.html', {
        'settlements': settlements[:300],
        'employees': Employee.objects.all().order_by('name'),
        'emp_id': emp_id,
        'can_manage': has_perm(request.user, 'settlement_manage'),
    })


@login_required
def settlement_add(request):
    if not has_perm(request.user, 'settlement_manage'):
        messages.error(request, 'Access denied.')
        return redirect('settlement_list')

    employees = Employee.objects.filter(status='active').order_by('name')

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        from_date = _d(request, 'from_date')
        to_date = _d(request, 'to_date')

        if not from_date or not to_date or to_date < from_date:
            messages.error(request, 'Please provide a valid period (From Date / To Date).')
            return render(request, 'attendance_app/settlement_form.html', {
                'employees': employees, 'form_data': request.POST, 'action': 'Add', 'preselect_emp': '',
            })

        s = Settlement.objects.create(
            employee=emp,
            from_date=from_date,
            to_date=to_date,
            ot_amount=_f(request, 'ot_amount', 0),
            leave_days=_f(request, 'leave_days', 0),
            leave_deduction_amount=_f(request, 'leave_deduction_amount', 0),
            net_amount=_f(request, 'net_amount', 0),
            settlement_date=_d(request, 'settlement_date') or timezone.localdate(),
            notes=request.POST.get('notes', '').strip(),
            created_by=request.user,
        )
        log_activity(request.user, 'create', 'Settlement', s,
                     f'Added settlement for {emp.name} ({from_date} to {to_date})', request)
        messages.success(request, f'Settlement recorded for {emp.name}.')
        return redirect('settlement_list')

    return render(request, 'attendance_app/settlement_form.html', {
        'employees': employees, 'action': 'Add', 'form_data': defaultdict(str),
        'preselect_emp': request.GET.get('employee', ''),
    })


@login_required
def settlement_edit(request, pk):
    if not has_perm(request.user, 'settlement_manage'):
        messages.error(request, 'Access denied.')
        return redirect('settlement_list')

    s = get_object_or_404(Settlement, pk=pk)
    employees = Employee.objects.all().order_by('name')

    if request.method == 'POST':
        emp = get_object_or_404(Employee, pk=request.POST.get('employee'))
        from_date = _d(request, 'from_date')
        to_date = _d(request, 'to_date')

        if not from_date or not to_date or to_date < from_date:
            messages.error(request, 'Please provide a valid period (From Date / To Date).')
            return render(request, 'attendance_app/settlement_form.html', {
                'employees': employees, 'form_data': request.POST, 'action': 'Edit', 's': s, 'preselect_emp': '',
            })

        s.employee = emp
        s.from_date = from_date
        s.to_date = to_date
        s.ot_amount = _f(request, 'ot_amount', 0)
        s.leave_days = _f(request, 'leave_days', 0)
        s.leave_deduction_amount = _f(request, 'leave_deduction_amount', 0)
        s.net_amount = _f(request, 'net_amount', 0)
        s.settlement_date = _d(request, 'settlement_date') or s.settlement_date
        s.notes = request.POST.get('notes', '').strip()
        s.updated_by = request.user
        s.save()

        log_activity(request.user, 'update', 'Settlement', s,
                     f'Updated settlement for {emp.name} ({from_date} to {to_date})', request)
        messages.success(request, 'Settlement updated.')
        return redirect('settlement_list')

    return render(request, 'attendance_app/settlement_form.html', {
        'employees': employees, 'action': 'Edit', 's': s, 'form_data': defaultdict(str), 'preselect_emp': '',
    })


@login_required
def settlement_delete(request, pk):
    if not has_perm(request.user, 'settlement_manage'):
        messages.error(request, 'Access denied.')
        return redirect('settlement_list')

    s = get_object_or_404(Settlement, pk=pk)
    if request.method == 'POST':
        emp_id = s.employee_id
        name = s.employee.name
        s.delete()
        log_activity(request.user, 'delete', 'Settlement',
                     description=f'Deleted settlement for {name}', request=request)
        messages.success(request, 'Settlement deleted.')
        return_to = request.POST.get('return_to', '')
        if return_to == 'employee':
            return redirect(f'/employees/{emp_id}/?tab=settlements')
    return redirect('settlement_list')
