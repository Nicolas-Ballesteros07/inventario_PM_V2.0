from django.urls import path
from . import views 

urlpatterns = [
    # PARA VER LOS DATOS DEL INFORME
    path('', views.informe, name='informe'),
    # PARA SUBIR LOS ARCHIVOS DESDE EL MODAL PARA EL INFORME
    path('subir/', views.subir_archivos, name='subir_archivos'),
    # GENERAR EL REPORTE
    path('informe/exportar-excel/', views.descargar_reporte_excel, name='descargar_reporte_excel'),
    # PARA VER LOS PRODUCTOS EN UNA VISTA DE LISTA
    path("productos/", views.lista_productos, name="lista_productos"),
    # PARA CREAR UN PRODUCTO
    path("productos/crear/",views.crear_producto_view,name="crear_producto"),

    # url para editar un producto
    path("productos/actualizar/<uuid:id_random>/",views.actualizar_producto_view,name="actualizar_producto"),

    # PARA EXCLUIR PRODUCTOS (MANUAL Y POR EXCEL)
    path('excluir-productos/', views.excluir_productos, name='excluir_productos'),

    # PARA VER LOS PRODUCTOS EXCLUIDOS
    path('productos-excluidos/', views.productos_excluidos, name='productos_excluidos'),
    # PARA REACTIVAR UN PRODUCTO EXCLUIDO
    path('productos-excluidos/reactivar/<int:exclusion_id>/', views.reactivar_producto_view, name='reactivar_producto'),
    
    # PARA DESCARGAR EL REPORTE DE PRODUCTOS EXCLUIDOS
    path("productos-excluidos/reporte/",views.descargar_reporte_productos_excluidos,name="descargar_reporte_productos_excluidos",),

    # VISUALIZAR LOS DATOS DE DETALLES DE COMPRA
    path('compras/', views.compras_lista, name='compras_lista'),

    # CARGAR EL ARCHIVO DE DETALLES DE COMPRA
    path('compras/cargar/', views.cargar_compras, name='cargar_compras'),

    # url para la visualizacion de productos en bajo stock
    path('bajo-duracion/', views.productos_bajo_duracion, name='productos_bajo_duracion'),

    # url para la visualizacion de productos y su categoria
    path("categorizacion/",views.categorizacion_view,name="categorizacion"),

    # url para cargar el archivo de categorizacion
    path("categorizacion/cargar/",views.cargar_excel_categorizacion,name="cargar_excel_categorizacion"),

    # url para descargar el informe de productos con bajo stock
    path('reporte-excel/', views.descargar_reporte_bajo_stock, name='reporte_excel_productos'),

    # url para la visualizacion y la validacion de los productos como individual y general
    path('api/comparativa/<str:codigo_mantis>/', views.comparativa_producto_view, name='comparativa_producto'),
    path('api/comparativa-general/', views.comparativa_general_view, name='comparativa_general'),

    # url para ver los productos y la fecha de vencimientos
    path('vencimientos/', views.lista_vencimientos, name='lista_vencimientos'),
    # url para cargar los archivos
    path('vencimientos/cargar/', views.cargar_excel_vencimientos, name='cargar_excel_vencimientos'),
    # url para descargar el reporte de productos con fecha de vencimiento
    path("vencimientos/reporte/",views.descargar_reporte_vencimientos,name="descargar_reporte_vencimientos",),


]   