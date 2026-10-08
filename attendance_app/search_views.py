#search_views.py
"""Global search across Employees, Documents, Letters, and Settlements."""
from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import render

from .models import Employee, EmployeeDocument, GeneratedLetter, Settlement


@login_required
def global_search(request):
    q = request.GET.get('q', '').strip()

    employees = []
    documents = []
    letters = []
    settlements = []

    if q:
        employees = Employee.objects.filter(
            Q(name__icontains=q) | Q(emirates_id__icontains=q) |
            Q(mobile__icontains=q) | Q(job_title__icontains=q)
        ).order_by('name')[:25]

        documents = EmployeeDocument.objects.select_related('employee', 'doc_type').filter(
            Q(employee__name__icontains=q) | Q(doc_number__icontains=q) | Q(doc_type__name__icontains=q)
        ).order_by('expiry_date')[:25]

        letters = GeneratedLetter.objects.select_related('employee').filter(
            Q(employee__name__icontains=q) | Q(reference_no__icontains=q)
        ).order_by('-created_at')[:25]

        settlements = Settlement.objects.select_related('employee').filter(
            Q(employee__name__icontains=q) | Q(notes__icontains=q)
        ).order_by('-settlement_date')[:25]

    total = len(employees) + len(documents) + len(letters) + len(settlements)

    return render(request, 'attendance_app/search_results.html', {
        'q': q,
        'employees': employees,
        'documents': documents,
        'letters': letters,
        'settlements': settlements,
        'total': total,
    })
