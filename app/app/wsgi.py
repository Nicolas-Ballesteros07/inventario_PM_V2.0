import os
import sys
from django.core.wsgi import get_wsgi_application

# Agregar el directorio raíz al path
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

app_dir = os.path.join(project_root, 'app')
if app_dir not in sys.path:
    sys.path.insert(0, app_dir)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.app.settings')

application = get_wsgi_application()
app = application