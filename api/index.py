import os
import sys

# Obtenemos la ruta de la carpeta 'app' (donde están manage.py y tus apps)
# __file__ es /var/task/api/index.py, así que subimos un nivel y entramos a 'app'
path_to_apps = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'app')
sys.path.append(path_to_apps)

# Establecemos el módulo de configuración
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "app.settings")

from django.core.wsgi import get_wsgi_application
app = get_wsgi_application()