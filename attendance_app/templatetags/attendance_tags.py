from django import template

register = template.Library()

@register.filter
def split(value, delimiter=','):
    """Split a string by delimiter"""
    return value.split(delimiter)

@register.filter
def make_list(value):
    """Convert string to list of characters"""
    return list(str(value))

@register.filter  
def get_item(dictionary, key):
    """Get item from dict by key"""
    return dictionary.get(key)

@register.filter
def index(lst, i):
    """Get item from list by index"""
    try:
        return lst[int(i)]
    except (IndexError, ValueError, TypeError):
        return ''

@register.filter
def has_perm_tag(user_perms, perm):
    """Check if perm is in user_perms list"""
    return perm in user_perms

@register.simple_tag
def attendance_status_badge(status):
    """Return badge class for attendance status"""
    mapping = {
        'present': 'badge-green',
        'absent': 'badge-red',
        'half_day': 'badge-yellow',
        'leave': 'badge-purple',
        'holiday': 'badge-blue',
    }
    return mapping.get(status, 'badge-gray')


@register.filter
def att_status(att_dict, date_str):
    """Get attendance status for a given date string from dict."""
    att = att_dict.get(date_str)
    if att:
        return att.status if hasattr(att, 'status') else ''
    return ''


@register.filter
def att_obj(att_dict, date_key):
    """Get attendance object for a given date key from dict."""
    return att_dict.get(date_key)
