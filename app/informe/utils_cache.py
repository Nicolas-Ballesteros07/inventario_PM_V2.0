from django.core.cache import cache

CACHE_TIMEOUT = 60 * 60 * 6  # 6 horas (respaldo; se invalida manualmente al escribir)

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


def invalidar_cache_carga():
    """Después de subir un nuevo informe (subir_archivos)."""
    cache.delete_many(CLAVES_DEPENDIENTES_DE_CARGA)


def invalidar_cache_producto():
    """Después de excluir/reactivar/crear productos."""
    cache.delete_many(CLAVES_DEPENDIENTES_DE_PRODUCTO)


def invalidar_cache_compra():
    """Después de cargar_compras."""
    cache.delete_many(CLAVES_DEPENDIENTES_DE_COMPRA)