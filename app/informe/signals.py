from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver
from .models import Producto, Compra, ProductoExcluido, DetalleInforme, CargaReporte, Categorizacion
from .utils_cache import (
    invalidar_cache_producto,
    invalidar_cache_compra,
    invalidar_cache_carga,
)


@receiver([post_save, post_delete], sender=Producto)
def _invalidar_cache_producto_signal(sender, **kwargs):
    invalidar_cache_producto()


@receiver([post_save, post_delete], sender=ProductoExcluido)
def _invalidar_cache_excluido_signal(sender, **kwargs):
    invalidar_cache_producto()


@receiver([post_save, post_delete], sender=Compra)
def _invalidar_cache_compra_signal(sender, **kwargs):
    invalidar_cache_compra()


@receiver([post_save, post_delete], sender=DetalleInforme)
def _invalidar_cache_detalle_signal(sender, **kwargs):
    invalidar_cache_carga()


@receiver(post_delete, sender=CargaReporte)
def _invalidar_cache_carga_signal(sender, **kwargs):
    invalidar_cache_carga()