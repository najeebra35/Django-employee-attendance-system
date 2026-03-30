from functools import wraps
from django.shortcuts import redirect
from django.contrib import messages


def permission_required(perm):
    def decorator(view_func):
        @wraps(view_func)
        def wrapper(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('login')
            if request.user.is_superuser:
                return view_func(request, *args, **kwargs)
            try:
                from .models import UserPermission
                up = UserPermission.objects.get(user=request.user)
                if perm in up.permissions:
                    return view_func(request, *args, **kwargs)
            except Exception:
                pass
            messages.error(request, 'You do not have permission to access this page.')
            return redirect('dashboard')
        return wrapper
    return decorator
