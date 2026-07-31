import uuid
from django.db import models


class Producto(models.Model):
    id_random = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    fecha_creacion = models.DateTimeField(auto_now_add=True)
    codigo_mantis = models.CharField(max_length=100, unique=True)
    barras = models.CharField(max_length=100, blank=True, null=True)
    descripcion = models.TextField(blank=True, null=True)
    laboratorio = models.CharField(max_length=200, blank=True, null=True)
    estado = models.BooleanField(default=True)

    class Meta:
        db_table = "productos"
        verbose_name = "Producto"
        verbose_name_plural = "Productos"

    def __str__(self):
        return f"{self.codigo_mantis} - {self.descripcion}"


class CargaReporte(models.Model):
    TIPO_CARGA_CHOICES = [
        ('CARGA_INFORME', 'Carga de Informe (Listado Básico + Ventas)'),
        ('DETALLE_COMPRA', 'Detallado de Compra'),
    ]

    id_carga = models.AutoField(primary_key=True)
    fecha_carga = models.DateTimeField(auto_now_add=True)
    tipo_carga = models.CharField(
        max_length=20,
        choices=TIPO_CARGA_CHOICES,
        default='CARGA_INFORME'
    )

    class Meta:
        db_table = "cargas_reportes"
        verbose_name = "Carga de Reporte"
        verbose_name_plural = "Cargas de Reportes"

    def __str__(self):
        return f"Carga {self.id_carga} - {self.get_tipo_carga_display()}"


class DetalleInforme(models.Model):
    id_formulario = models.AutoField(primary_key=True)

    producto = models.ForeignKey(
        Producto,
        on_delete=models.CASCADE,
        related_name='detalles'
    )

    carga = models.ForeignKey(
        CargaReporte,
        on_delete=models.CASCADE,
        related_name='detalles'
    )

    fecha_reporte = models.DateTimeField(auto_now_add=True)
    stock = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    rentabilidad = models.DecimalField(max_digits=8, decimal_places=4, default=0)
    rotacion = models.DecimalField(max_digits=12, decimal_places=3, default=0)
    costo = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    duracion = models.DecimalField(max_digits=12, decimal_places=1, default=0)
    observacion = models.TextField(blank=True, null=True)
    pedidos = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    solicitar = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    ind_duracion = models.CharField(max_length=50, blank=True, null=True)
    periodo_inicio = models.DateField(null=True, blank=True)
    periodo_fin = models.DateField(null=True, blank=True)
    periodo_meses = models.PositiveIntegerField(null=True, blank=True)

    class Meta:
        db_table = "detalle_informes"
        verbose_name = "Detalle de Informe"
        verbose_name_plural = "Detalles de Informes"

    def __str__(self):
        return f"Detalle {self.id_formulario} - {self.producto.codigo_mantis}"


class ProductoExcluido(models.Model):
    producto = models.ForeignKey(
        Producto,
        on_delete=models.CASCADE,
        related_name='exclusiones',
        db_column='id_producto'
    )

    fecha_creacion = models.DateTimeField(auto_now_add=True)
    activo = models.BooleanField(default=True)
    observacion = models.CharField(max_length=255, blank=True, null=True)

    class Meta:
        db_table = "productos_excluidos"
        verbose_name = "Producto Excluido"
        verbose_name_plural = "Productos Excluidos"
        ordering = ['-fecha_creacion']
        constraints = [
            models.UniqueConstraint(
                fields=['producto'],
                condition=models.Q(activo=True),
                name='unique_producto_activo'
            )
        ]

    def __str__(self):
        return f"{self.producto.codigo_mantis} — {self.producto.descripcion}"

    @property
    def codigo_mantis(self):
        return self.producto.codigo_mantis

    @property
    def descripcion(self):
        return self.producto.descripcion


class Compra(models.Model):
    id_compra = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)

    producto = models.ForeignKey(
        Producto,
        on_delete=models.CASCADE,
        related_name='compras',
        db_column='id_producto'
    )

    fecha_compra = models.DateField()
    factura = models.CharField(max_length=100)
    proveedor = models.CharField(max_length=255)
    valor_compra = models.DecimalField(max_digits=12, decimal_places=2)
    periodo_inicio = models.DateField(null=True, blank=True)
    periodo_fin = models.DateField(null=True, blank=True)

    class Meta:
        db_table = "compras"
        verbose_name = "Compra"
        verbose_name_plural = "Compras"
        ordering = ['-fecha_compra']
        constraints = [
            models.UniqueConstraint(
                fields=['producto', 'factura', 'fecha_compra'],
                name='unique_compra_producto_factura_fecha'
            )
        ]

    def __str__(self):
        return f"{self.factura} - {self.producto.codigo_mantis}"


class Categorizacion(models.Model):
    TIPO_CATEGORIA_CHOICES = [
        ('A', 'Categoría A'),
        ('B', 'Categoría B'),
        ('C', 'Categoría C'),
    ]

    id_categorizacion = models.AutoField(primary_key=True)
    fecha_registro = models.DateTimeField(auto_now_add=True)

    tipo_categoria = models.CharField(
        max_length=1,
        choices=TIPO_CATEGORIA_CHOICES
    )

    analisis = models.TextField(blank=True, null=True)

    producto = models.ForeignKey(
        Producto,
        on_delete=models.CASCADE,
        related_name='categorizaciones',
        db_column='id_producto'
    )

    class Meta:
        db_table = "categorizaciones"
        verbose_name = "Categorización"
        verbose_name_plural = "Categorizaciones"
        ordering = ['-fecha_registro']
        constraints = [
            models.UniqueConstraint(
                fields=['producto'],
                name='unique_categorizacion_producto'
            )
        ]

    def __str__(self):
        return f"{self.producto.codigo_mantis} - Categoría {self.tipo_categoria}"


class Vencimiento(models.Model):
    id_vencimiento = models.AutoField(primary_key=True)
    fecha_carga = models.DateTimeField(auto_now_add=True)
    fecha_vencimiento = models.DateField()

    producto = models.ForeignKey(
        Producto,
        on_delete=models.CASCADE,
        related_name="vencimientos",
        db_column="id_producto"
    )

    lote = models.CharField(max_length=100)
    unidades = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)

    class Meta:
        db_table = "fechas_vencimiento"
        verbose_name = "Vencimiento"
        verbose_name_plural = "Vencimientos"
        ordering = ["fecha_vencimiento"]
        constraints = [
            models.UniqueConstraint(
                fields=["producto", "lote"],
                name="unique_producto_lote"
            )
        ]

    def __str__(self):
        return f"{self.producto.codigo_mantis} - Lote {self.lote}"