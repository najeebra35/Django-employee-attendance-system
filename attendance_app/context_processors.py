from .models import UserPermission, ActivityLog

from .models import CompanySettings

def company_settings(request):
    return {
        'company': CompanySettings.get_settings()
    }

def user_permissions(request):
    """Add user permissions and portal flag to all template contexts."""
    if request.user.is_authenticated:
        if request.user.is_superuser:
            perms = [p[0] for p in UserPermission.PERMISSION_CHOICES]
        else:
            try:
                up = UserPermission.objects.get(user=request.user)
                perms = up.permissions
            except UserPermission.DoesNotExist:
                perms = []

        # Check if user is an employee portal user
        is_portal_user = False
        try:
            if request.user.employee_profile:
                is_portal_user = True
        except Exception:
            pass

        return {
            'user_perms': perms,
            'is_admin': request.user.is_superuser,
            'is_portal_user': is_portal_user,
        }
    return {'user_perms': [], 'is_admin': False, 'is_portal_user': False}
