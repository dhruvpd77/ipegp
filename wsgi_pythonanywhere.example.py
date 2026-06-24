# Copy this into PythonAnywhere → Web → WSGI configuration file.
# Replace YOUR_USERNAME with your PythonAnywhere username.

import os
import sys

PROJECT_HOME = '/home/YOUR_USERNAME/ipegp'
if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

os.environ['DJANGO_SETTINGS_MODULE'] = 'ipe_gp.settings'
os.environ['DJANGO_DEBUG'] = 'False'
os.environ['DJANGO_ALLOWED_HOSTS'] = 'YOUR_USERNAME.pythonanywhere.com'
os.environ['DJANGO_CSRF_TRUSTED_ORIGINS'] = 'https://YOUR_USERNAME.pythonanywhere.com'
os.environ['DJANGO_SECRET_KEY'] = 'replace-with-a-long-random-secret-key'

from django.core.wsgi import get_wsgi_application

application = get_wsgi_application()
