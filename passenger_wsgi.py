'''import os
import sys


def application(environ, start_response):
    start_response('200 OK', [('Content-Type', 'text/plain')])
    message = 'It works!\n'
    version = 'Python %s\n' % sys.version.split()[0]
    response = '\n'.join([message, version])
    return [response.encode()]'''

import os
import sys

# Add project to path
sys.path.insert(0, os.path.dirname(__file__))

os.environ['DJANGO_SETTINGS_MODULE'] = 'attendance_system.settings'

from django.core.wsgi import get_wsgi_application
application = get_wsgi_application()