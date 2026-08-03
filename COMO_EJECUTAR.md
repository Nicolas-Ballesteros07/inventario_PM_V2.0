# Cómo ejecutar el proyecto

## 1. Introducción breve

Esta aplicación es un sistema web de Django para gestionar inventario, cargar reportes de ventas y stock, analizar rotación, categorizar productos, registrar compras, controlar vencimientos y exportar reportes en Excel. El proyecto está organizado como un proyecto Django estándar con una única app principal llamada [app/informe](app/informe), que concentra la lógica de negocio, los modelos, las vistas y los utilitarios.

El framework usado es Django 6.0.7. La parte web está montada sobre la estructura típica de Django: [app/app](app/app) contiene la configuración del proyecto y [app/informe](app/informe) contiene los módulos funcionales del negocio.

## 2. Funcionamiento

El flujo real del proyecto, según el código actual, es el siguiente:

1. Una petición HTTP entra al proyecto Django a través de [app/app/urls.py](app/app/urls.py). Allí se enruta todo lo que no sea el panel de administración a [app/informe/urls.py](app/informe/urls.py).
2. Cada ruta de [app/informe/urls.py](app/informe/urls.py) está asociada a una vista en [app/informe/views.py](app/informe/views.py). Por ejemplo, la ruta raíz carga la vista de informe, mientras que las rutas de productos, compras, categorización y vencimientos delegan en vistas específicas.
3. La vista principal de informe consulta la última carga de reporte almacenada en la base de datos, obtiene los detalles asociados y aplica filtros para excluir productos. Si los datos no están ya en caché, los construye a partir de los modelos [app/informe/models.py](app/informe/models.py) y los guarda en una tabla de caché configurada en [app/app/settings.py](app/app/settings.py).
4. Para cargar un nuevo informe, la ruta [app/informe/urls.py](app/informe/urls.py) llama a la vista de subida. Esa vista valida los archivos recibidos con [app/informe/forms.py](app/informe/forms.py), los lee en memoria y delega el procesamiento a [app/informe/utils.py](app/informe/utils.py) mediante la función `procesar_archivos`.
5. La carga de Excel no escribe los archivos en disco; los procesa desde el objeto subido en memoria. `procesar_archivos` lee el listado básico y el archivo de ventas, extrae fechas del período, calcula la rotación y la duración, crea o actualiza productos y genera registros de `DetalleInforme` para la carga correspondiente.
6. El resultado se guarda directamente en la base de datos (modelos `Producto`, `CargaReporte`, `DetalleInforme`, `ProductoExcluido`, `Compra`, `Categorizacion` y `Vencimiento`). La vista luego devuelve una respuesta JSON o renderiza una plantilla HTML según el caso.
7. Para exportar reportes, las vistas llaman a utilidades específicas en [app/informe/utils_generar_reportes.py](app/informe/utils_generar_reportes.py), que generan archivos Excel en memoria y los devuelven como respuesta HTTP para descarga.
8. Para los módulos de compras, categorización y vencimientos, el flujo es similar: una vista recibe el archivo, lo valida, lo procesa con una función de utilidades y luego actualiza o crea registros en base de datos. En algunos casos, las vistas también devuelven respuestas JSON para uso de frontend o integraciones.

## 3. Estructura técnica

| Archivo o carpeta | Rol |
|---|---|
| [app](app) | Carpeta raíz del proyecto Django. |
| [app/app](app/app) | Contiene la configuración del proyecto: settings, URLs, ASGI y WSGI. |
| [app/app/settings.py](app/app/settings.py) | Configuración principal de Django, apps instaladas, base de datos, caché, statics y middleware. |
| [app/app/urls.py](app/app/urls.py) | Punto de entrada de URLs del proyecto. |
| [app/informe](app/informe) | App principal del negocio: reportes, productos, compras, categorización y vencimientos. |
| [app/informe/models.py](app/informe/models.py) | Modelos de datos: Producto, CargaReporte, DetalleInforme, ProductoExcluido, Compra, Categorizacion y Vencimiento. |
| [app/informe/views.py](app/informe/views.py) | Lógica de control de cada pantalla y endpoint. |
| [app/informe/urls.py](app/informe/urls.py) | Definición de rutas de la app. |
| [app/informe/forms.py](app/informe/forms.py) | Formularios de subida para informes, exclusiones y compras. |
| [app/informe/utils.py](app/informe/utils.py) | Lógica de procesamiento de archivos Excel, cálculos de rotación/duración y utilidades de negocio. |
| [app/informe/utils_generar_reportes.py](app/informe/utils_generar_reportes.py) | Generación de reportes Excel para inventario, compras, categorizaciones y vencimientos. |
| [app/informe/utils_cache.py](app/informe/utils_cache.py) | Invalidación y gestión de caché de la aplicación. |
| [app/templates](app/templates) | Plantillas HTML usadas por las vistas. |
| [app/static](app/static) | Archivos estáticos y plantilla base de Excel con macros. |
| [app/db.sqlite3](app/db.sqlite3) | Archivo SQLite presente en el repo, pero no se usa por la configuración actual. |
| [api/index.py](api/index.py) | Entry point para despliegue en Vercel. |
| [requirements.txt](requirements.txt) | Dependencias Python del proyecto. |
| [build_files.sh](build_files.sh) | Script de preparación para despliegue/ejecución. |
| [vercel.json](vercel.json) | Configuración de Vercel para servir la app Django. |

## 4. Endpoints principales

| Ruta | Función | Propósito |
|---|---|---|
| `/admin/` | `admin.site.urls` | Panel administrativo de Django. |
| `/` | `informe` | Vista principal del informe con resumen y detalle del último período cargado. |
| `/subir/` | `subir_archivos` | Recibe los archivos de listado básico y ventas para procesar el informe. |
| `/informe/exportar-excel/` | `descargar_reporte_excel` | Descarga el reporte de inventario en Excel. |
| `/productos/` | `lista_productos` | Muestra la lista de productos. |
| `/productos/crear/` | `crear_producto_view` | Crea un producto nuevo. |
| `/productos/actualizar/<uuid:id_random>/` | `actualizar_producto_view` | Actualiza un producto existente. |
| `/excluir-productos/` | `excluir_productos` | Excluye productos manualmente o desde un Excel. |
| `/productos-excluidos/` | `productos_excluidos` | Muestra la lista de productos excluidos. |
| `/productos-excluidos/reactivar/<int:exclusion_id>/` | `reactivar_producto_view` | Reactiva un producto excluido. |
| `/productos-excluidos/reporte/` | `descargar_reporte_productos_excluidos` | Descarga un Excel con los productos excluidos. |
| `/compras/` | `compras_lista` | Muestra el listado de compras. |
| `/compras/cargar/` | `cargar_compras` | Carga compras desde un Excel. |
| `/bajo-duracion/` | `productos_bajo_duracion` | Muestra los productos con duración menor a 1.0 y ofrece un endpoint JSON con `?api=true`. |
| `/categorizacion/` | `categorizacion_view` | Muestra y guarda categorizaciones de productos. |
| `/categorizacion/cargar/` | `cargar_excel_categorizacion` | Importa categorizaciones desde un Excel. |
| `/categorizaciones/reporte/` | `descargar_reporte_categorizaciones` | Descarga un Excel con las categorizaciones. |
| `/reporte-excel/` | `descargar_reporte_bajo_stock` | Descarga un Excel de productos con baja duración. |
| `/api/comparativa/<str:codigo_mantis>/` | `comparativa_producto_view` | Devuelve la comparativa de rotación de un producto específico en JSON. |
| `/api/comparativa-general/` | `comparativa_general_view` | Devuelve una comparativa general en JSON. |
| `/vencimientos/` | `lista_vencimientos` | Muestra los vencimientos registrados. |
| `/vencimientos/cargar/` | `cargar_excel_vencimientos` | Importa vencimientos desde un Excel. |
| `/vencimientos/reporte/` | `descargar_reporte_vencimientos` | Descarga un Excel con los vencimientos filtrados por fechas. |

## 5. Requisitos

- Lenguaje y runtime:
  - Python 3.14.5 en el entorno virtual actual del repo, según [env/pyvenv.cfg](env/pyvenv.cfg).
  - No se usa Node.js en este proyecto.
- Dependencias principales:
  - Django 6.0.7
  - dj-database-url 3.1.2
  - psycopg2-binary 2.9.12
  - openpyxl 3.1.5
  - xlrd 2.0.2
  - whitenoise 6.12.0
  - gunicorn 26.0.0
  - requests 2.34.2
  - python-dateutil 2.9.0.post0
- Base de datos:
  - La configuración actual en [app/app/settings.py](app/app/settings.py) usa `dj_database_url.config(...)` y apunta a PostgreSQL mediante `DATABASE_URL`.
  - Si no se define la variable, el código usa un valor por defecto con una conexión a Neon que ya está escrita en el archivo.
  - El archivo [app/db.sqlite3](app/db.sqlite3) existe, pero el proyecto actual no lo usa porque la configuración apunta a PostgreSQL.
- Servicios externos:
  - No hay integración visible con correo, S3, Azure Blob Storage ni APIs externas en el código actual.
  - El proyecto sí usa caché de base de datos (`django_cache_table`) y archivos Excel como entrada/salida principal.

## 6. Estructura recomendada de ejecución

Estas instrucciones son para Windows con PowerShell desde la raíz del proyecto.

1. Abrir PowerShell en la raíz del proyecto:

   ```powershell
   cd C:\PROYECTO\inventario_PM_V2.0
   ```

2. Crear y activar un entorno virtual:

   ```powershell
   py -3 -m venv .\env
   .\env\Scripts\Activate.ps1
   ```

3. Actualizar `pip` e instalar dependencias:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. Definir la base de datos si se desea usar una distinta a la configuración por defecto. La variable real que lee el proyecto es `DATABASE_URL`:

   ```powershell
   $env:DATABASE_URL = "postgresql://usuario:password@host:5432/nombre_bd"
   ```

   Si no se define, el proyecto usará el valor por defecto que ya está en [app/app/settings.py](app/app/settings.py).

5. Ejecutar migraciones y preparar la caché de base de datos:

   ```powershell
   cd .\app
   python manage.py migrate
   python manage.py createcachetable
   ```

6. Recolectar archivos estáticos si se va a servir la app con WhiteNoise o en un entorno más parecido a producción:

   ```powershell
   python manage.py collectstatic --noinput
   ```

7. Levantar el servidor local:

   ```powershell
   python manage.py runserver 0.0.0.0:8000
   ```

8. Abrir la aplicación en el navegador en:

   ```text
   http://127.0.0.1:8000/
   ```

## 7. Opción con script

El repositorio incluye [build_files.sh](build_files.sh), que prepara el entorno de despliegue con los siguientes pasos:

1. Instala las dependencias desde [requirements.txt](requirements.txt).
2. Ejecuta `collectstatic` para preparar los archivos estáticos.
3. Ejecuta `makemigrations` para generar migraciones si hay cambios en los modelos.
4. Ejecuta `migrate` para aplicar las migraciones a la base de datos.
5. Ejecuta `createcachetable` para crear la tabla de caché usada por la app.

En Windows, el script no está pensado para ejecutarse directamente como `.sh`, pero su lógica es la que debe seguirse manualmente si se quiere preparar el entorno de forma equivalente.

## 8. Notas finales

- El proyecto está pensado principalmente para trabajar con reportes en Excel y operaciones masivas sobre productos, compras, categorizaciones y vencimientos.
- La lógica de negocio más importante está en [app/informe/utils.py](app/informe/utils.py) y [app/informe/views.py](app/informe/views.py).
- La generación de reportes Excel depende de la plantilla [app/static/plantillas/plantilla_reporte_base.xlsm](app/static/plantillas/plantilla_reporte_base.xlsm), que se usa para crear el Excel de baja duración.
- Para despliegues en Vercel, el proyecto está preparado con [vercel.json](vercel.json) y [api/index.py](api/index.py), usando Django como WSGI application.
- Si el proyecto se despliega en un entorno nuevo, conviene revisar la configuración de `DATABASE_URL` y la presencia de la tabla de caché, porque la app usa caché de base de datos y no sólo memoria local.
