from .models import ActivityLog


class ActivityLogMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        return response


def log_activity(user, action, model_name, obj=None, description='', request=None):
    """Helper to create activity logs"""
    ip = None
    if request:
        x_forwarded = request.META.get('HTTP_X_FORWARDED_FOR')
        ip = x_forwarded.split(',')[0] if x_forwarded else request.META.get('REMOTE_ADDR')

    ActivityLog.objects.create(
        user=user,
        action=action,
        model_name=model_name,
        object_id=obj.pk if obj else None,
        object_repr=str(obj) if obj else '',
        description=description,
        ip_address=ip,
    )
