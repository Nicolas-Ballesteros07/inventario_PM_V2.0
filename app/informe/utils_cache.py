import logging
from django.core.cache import cache

logger = logging.getLogger(__name__)

CACHE_TIMEOUT = 60 * 15  # 15 minutos

CACHE_KEY_INFORME = 'cache_informe_datos'
CACHE_KEY_LISTA_PRODUCTOS = 'cache_lista_productos'
CACHE_KEY_COMPRAS = 'cache_compras_lista'
CACHE_KEY_EXCLUIDOS = 'cache_productos_excluidos'
CACHE_KEY_BAJO_DURACION = 'cache_productos_bajo_duracion'

# Agrupaciones de invalidación según qué datos cambiaron
CLAVES_DEPENDIENTES_DE_CARGA = [
    CACHE_KEY_INFORME,
    CACHE_KEY_BAJO_DURACION,
]

CLAVES_DEPENDIENTES_DE_PRODUCTO = [
    CACHE_KEY_LISTA_PRODUCTOS,
    CACHE_KEY_EXCLUIDOS,
    CACHE_KEY_INFORME,
    CACHE_KEY_BAJO_DURACION,
]

CLAVES_DEPENDIENTES_DE_COMPRA = [
    CACHE_KEY_COMPRAS,
    CACHE_KEY_BAJO_DURACION,
]


def _invalidar_claves(claves, origen):
    """
    El caché es SOLO un acelerador de lecturas repetidas; nunca debe
    condicionar si los datos se guardaron o no en la base de datos.
    Si falla la invalidación (ej. tabla de caché caída), se registra
    el error pero el flujo de la vista continúa con normalidad —
    la página simplemente seguirá sirviendo datos cacheados hasta el
    próximo TTL, en vez de romper el guardado.
    """
    try:
        cache.delete_many(claves)
    except Exception:
        logger.exception(f"No se pudo invalidar el caché ({origen}); "
                          f"los datos ya se guardaron correctamente en la BD.")


def invalidar_cache_carga():
    """Después de subir un nuevo informe (subir_archivos)."""
    _invalidar_claves(CLAVES_DEPENDIENTES_DE_CARGA, origen="carga de informe")


def invalidar_cache_producto():
    """Después de excluir/reactivar/crear productos."""
    _invalidar_claves(CLAVES_DEPENDIENTES_DE_PRODUCTO, origen="cambio de producto")


def invalidar_cache_compra():
    """Después de cargar_compras."""
    _invalidar_claves(CLAVES_DEPENDIENTES_DE_COMPRA, origen="carga de compras")