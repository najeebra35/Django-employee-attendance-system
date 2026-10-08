from django.conf import settings
from django.http import JsonResponse
from django.views.decorators.http import require_GET

from .models import Employee


def _check_api_key(request):
    provided = request.headers.get('X-Api-Key') or request.GET.get('api_key', '')
    expected = getattr(settings, 'EXTERNAL_API_KEY', '')
    return bool(expected) and provided == expected


def _cors(response):
    response['Access-Control-Allow-Origin'] = '*'
    response['Access-Control-Allow-Methods'] = 'GET'
    response['Access-Control-Allow-Headers'] = 'X-Api-Key, Content-Type'
    return response


@require_GET
def api_active_employees(request):
    if not _check_api_key(request):
        return _cors(JsonResponse({'error': 'Unauthorized — missing or invalid API key'}, status=401))

    employees = Employee.objects.filter(status='active').order_by('name')

    q = request.GET.get('q', '').strip()
    if q:
        employees = employees.filter(name__icontains=q)

    emp_type = request.GET.get('emp_type', '').strip()
    if emp_type in ('permanent', 'temporary'):
        employees = employees.filter(emp_type=emp_type)

    data = [
        {
            'id': e.id,
            'name': e.name,
            'job_title': e.job_title or '',
            'mobile': e.mobile or '',
            'emirates_id': e.emirates_id or '',
            'emp_type': e.emp_type,
        }
        for e in employees
    ]

    return _cors(JsonResponse({'count': len(data), 'results': data}))