from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.core.cache import cache
from .models import *
from .forms import *
from .utils import *
from .utils_generar_reportes import *
from .utils_cache import (
    CACHE_KEY_INFORME, CACHE_KEY_LISTA_PRODUCTOS, CACHE_KEY_COMPRAS,
    CACHE_KEY_EXCLUIDOS, CACHE_TIMEOUT,
    invalidar_cache_carga, invalidar_cache_producto, invalidar_cache_compra,
)
import tempfile, os
from django.db import transaction
from django.db.models import Count
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_GET, require_POST
from django.urls import reverse
import json
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.cache import never_cache
from django.db.models import Case, When, Value, IntegerField
import logging
logger = logging.getLogger(__name__)
#======================================
# VER EL INFORME (con caché)
#======================================
def informe(request):
    datos = cache.get(CACHE_KEY_INFORME)

    if datos is None:
        ultima_carga = CargaReporte.objects.order_by('-fecha_carga').first()
        detalles = []
        resumen = {
            'total': 0, 'sin_rotacion': 0, 'sin_stock': 0,
            'bajo_stock': 0, 'ok': 0, 'sobre_stock': 0,
        }

        if ultima_carga:
            productos_excluidos_ids = ProductoExcluido.objects.filter(
                activo=True
            ).values_list('producto_id', flat=True)

            detalles_qs = (
                ultima_carga.detalles
                .select_related('producto')
                .exclude(producto_id__in=productos_excluidos_ids)
            )
            detalles = list(detalles_qs)

            conteos = (
                detalles_qs.values('ind_duracion')
                .annotate(cantidad=Count('id_formulario'))
            )

            mapa_estados = {
                'SIN ROTACIÓN': 'sin_rotacion',
                'SIN STOCK': 'sin_stock',
                'BAJO STOCK': 'bajo_stock',
                'OK': 'ok',
                'SOBRE STOCK': 'sobre_stock',
            }

            for item in conteos:
                clave = mapa_estados.get(item['ind_duracion'])
                if clave:
                    resumen[clave] = item['cantidad']
                resumen['total'] += item['cantidad']

        datos = {
            'detalles': detalles,
            'carga': ultima_carga,
            'resumen': resumen,
            'periodo_inicio': detalles[0].periodo_inicio if detalles else None,
            'periodo_fin': detalles[0].periodo_fin if detalles else None,
        }
        cache.set(CACHE_KEY_INFORME, datos, CACHE_TIMEOUT)

    return render(request, 'informe.html', {
        **datos,
        'form': cargar_archivos_informe(),
        'form_manual': ExcluirProductosForm(),
        'form_excel': ExcluirProductosExcelForm(),
    })


# PARA PODER SUBIR LOS ARCHIVOS DESDE EL MODAL
def subir_archivos(request):
    """
    Procesa el formulario de carga vía fetch/AJAX. Siempre retorna JSON
    (status, message, errors, period_status) para que el front-end
    active el modal correspondiente. Todo el procesamiento se hace en
    memoria — nunca se escribe el archivo a disco.
    """
    if request.method != 'POST':
        return JsonResponse({
            'status': 'error',
            'message': 'Método no permitido.',
            'errors': [],
            'period_status': None,
        }, status=405)

    form = cargar_archivos_informe(request.POST, request.FILES)
    if not form.is_valid():
        errores_form = [
            f'{campo}: {", ".join(msgs)}'
            for campo, msgs in form.errors.items()
        ]
        return JsonResponse({
            'status': 'warning',
            'message': 'El formulario tiene campos inválidos o incompletos.',
            'errors': errores_form,
            'period_status': None,
        }, status=400)

    archivo_listado = request.FILES['archivo_listado_basico']
    archivo_ventas = request.FILES['archivo_ventas']

    # 1) Validación previa en memoria (estructura + duplicado de período)
    #    ANTES de procesar el archivo completo — evita trabajo innecesario
    #    y permite avisar al usuario si el período ya existía.
    validacion = validar_archivos_antes_de_procesar(archivo_listado, archivo_ventas)

    if not validacion['valido']:
        return JsonResponse({
            'status': 'error',
            'message': 'No se pudo validar la estructura de los archivos.',
            'errors': validacion['errores'],
            'period_status': None,
        }, status=422)

    # 2) Procesamiento real, siempre en memoria (sin FileField/ImageField,
    #    sin guardar el archivo en disco).
    try:
        carga = procesar_archivos(
            archivo_listado=archivo_listado,
            archivo_ventas=archivo_ventas,
        )
        invalidar_cache_carga()
        invalidar_cache_producto()
    except Exception as e:
        return JsonResponse({
            'status': 'error',
            'message': f'Ocurrió un error al procesar los archivos: {str(e)}',
            'errors': [],
            'period_status': None,
        }, status=500)

    mensaje = (
        'El período ya estaba cargado y se actualizó con los datos nuevos.'
        if validacion['period_status'] == 'actualizado'
        else 'El informe se generó correctamente con un nuevo período.'
    )

    return JsonResponse({
        'status': 'success',
        'message': mensaje,
        'errors': [],
        'period_status': validacion['period_status'],
    })


# GENERAR REPORTE DEL INFORME EXCEL
def descargar_reporte_excel(request):
    ultima_carga = CargaReporte.objects.order_by('-fecha_carga').first()

    if not ultima_carga:
        messages.error(request, 'No hay ningún informe cargado todavía para exportar.')
        return redirect('informe')

    productos_excluidos_ids = ProductoExcluido.objects.filter(
        activo=True
    ).values_list('producto_id', flat=True)

    detalles = (
        ultima_carga.detalles
        .select_related('producto')
        .exclude(producto_id__in=productos_excluidos_ids)
        .order_by('producto__codigo_mantis')
    )

    primer_detalle = detalles.first()
    meses_a_revisar = primer_detalle.periodo_meses if primer_detalle else 1

    buffer = generar_reporte_excel(detalles, meses_a_revisar)
    nombre_archivo = f"informe_inventario_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f'attachment; filename="{nombre_archivo}"'
    return response

#===========================================================================================
# Para la comparacion de periodos entre la rotacion de los produtos con el periodo anterior
#===========================================================================================
# Comparativa de UN producto (para el botón por fila)
def comparativa_producto_view(request, codigo_mantis):
    ultima_carga = CargaReporte.objects.order_by('-fecha_carga').first()
    if not ultima_carga:
        return JsonResponse({'error': 'No hay carga disponible.'}, status=404)

    detalle_actual = (
        DetalleInforme.objects
        .filter(carga=ultima_carga, producto__codigo_mantis=codigo_mantis)
        .select_related('producto')
        .first()
    )
    if not detalle_actual:
        return JsonResponse({'error': 'Producto no encontrado en la carga actual.'}, status=404)

    carga_anterior = obtener_carga_anterior(ultima_carga)
    resultado = comparar_producto(detalle_actual, carga_anterior)
    resultado['codigo'] = codigo_mantis
    resultado['descripcion'] = detalle_actual.producto.descripcion

    return JsonResponse(resultado)


# Comparativa general (para el botón "Ver en general")
def comparativa_general_view(request):
    ultima_carga = CargaReporte.objects.order_by('-fecha_carga').first()
    if not ultima_carga:
        return JsonResponse({'error': 'No hay carga disponible.'}, status=404)

    productos_excluidos_ids = ProductoExcluido.objects.filter(
        activo=True
    ).values_list('producto_id', flat=True)

    resultado = obtener_comparativa_general(ultima_carga, productos_excluidos_ids)
    return JsonResponse(resultado)


#=======================================
# LISTA DE PRODUCTOS (con caché)
#=======================================
def lista_productos(request):
    productos = cache.get(CACHE_KEY_LISTA_PRODUCTOS)

    if productos is None:
        productos_excluidos_ids = set(
            ProductoExcluido.objects.filter(activo=True)
            .values_list('producto_id', flat=True)
        )

        productos = list(
            Producto.objects.all().order_by("codigo_mantis")
        )

        # Anotamos en cada producto si está excluido, para usarlo en el template
        for p in productos:
            p.excluido = p.id_random in productos_excluidos_ids

        cache.set(CACHE_KEY_LISTA_PRODUCTOS, productos, CACHE_TIMEOUT)

    context = {
        "productos": productos,
        "form_manual": ExcluirProductosForm(),
        "form_excel": ExcluirProductosExcelForm(),
    }
    return render(request, "productos.html", context)


# VIEW PARA CREACION DE UN PRODUCTO NUEVO
def crear_producto_view(request):
    if request.method != "POST":
        return redirect("productos")

    data = {
        "codigo_mantis": request.POST.get("codigo_mantis"),
        "barras": request.POST.get("barras"),
        "descripcion": request.POST.get("descripcion"),
        "laboratorio": request.POST.get("laboratorio"),
        "estado": request.POST.get("estado") == "on",
    }

    resultado = crear_producto(data)

    if resultado["success"]:
        invalidar_cache_producto()  # <-- ESTO FALTA
        messages.success(request, resultado["message"])
    else:
        messages.error(request, resultado["message"])

    return redirect("lista_productos")

# view para actualizar un producto existente
def actualizar_producto_view(request, id_random):
    if request.method != "POST":
        return redirect("productos")

    producto = get_object_or_404(Producto, id_random=id_random)

    # ¿Tiene una exclusión activa?
    excluido = ProductoExcluido.objects.filter(
        producto=producto, activo=True
    ).exists()

    nuevo_estado = request.POST.get("estado") == "on"

    if excluido and nuevo_estado != producto.estado:
        messages.error(
            request,
            "Este producto está excluido. Debes quitarlo de la exclusión antes de cambiar su estado."
        )
        return redirect("lista_productos")

    data = {
        "codigo_mantis": request.POST.get("codigo_mantis"),
        "barras": request.POST.get("barras"),
        "descripcion": request.POST.get("descripcion"),
        "laboratorio": request.POST.get("laboratorio"),
        # Si está excluido, forzamos que el estado se mantenga igual (False)
        "estado": producto.estado if excluido else nuevo_estado,
    }

    resultado = actualizar_producto(id_random, data)

    if resultado["success"]:
        invalidar_cache_producto()
        messages.success(request, resultado["message"])
    else:
        messages.error(request, resultado["message"])

    return redirect("lista_productos")


#=======================================
# EXCLUIR PRODUCTOS (invalida caché al modificar)
#=======================================
def excluir_productos(request):
    if request.method == 'POST':
        es_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        next_url = request.POST.get('next', 'lista_productos')

        if 'enviar_manual' in request.POST:
            form_manual = ExcluirProductosForm(request.POST)
            if form_manual.is_valid():
                codigos = form_manual.limpiar_codigos()
                observacion = form_manual.cleaned_data.get('observacion', '')
                creados, no_encontrados, ya_excluidos = _excluir_codigos(codigos, observacion)
                if creados:
                    invalidar_cache_producto()

                if es_ajax:
                    return _respuesta_json_resumen(creados, no_encontrados, ya_excluidos)
                _mostrar_resumen(request, creados, no_encontrados, ya_excluidos)
            else:
                if es_ajax:
                    return JsonResponse({
                        'status': 'error',
                        'message': 'Formulario inválido.',
                        'errors': [str(e) for e in form_manual.errors.values()]
                    }, status=400)
                messages.error(request, 'Formulario inválido.')

        elif 'enviar_excel' in request.POST:
            form_excel = ExcluirProductosExcelForm(request.POST, request.FILES)
            if form_excel.is_valid():
                archivo = form_excel.cleaned_data['archivo_excel']
                observacion_general = form_excel.cleaned_data.get('observacion_general', '')

                ext = os.path.splitext(archivo.name)[1].lower()
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    for chunk in archivo.chunks():
                        tmp.write(chunk)
                    ruta_temporal = tmp.name

                try:
                    filas = leer_excel_exclusiones(ruta_temporal)
                except Exception as e:
                    mensaje = f'Error al leer el Excel: {str(e)}'
                    if es_ajax:
                        return JsonResponse({'status': 'error', 'message': mensaje}, status=400)
                    messages.error(request, mensaje)
                    filas = []
                finally:
                    os.remove(ruta_temporal)

                creados, no_encontrados, ya_excluidos = [], [], []
                for codigo, obs_fila in filas:
                    obs_final = obs_fila or observacion_general
                    c, ne, ye = _excluir_codigos([codigo], obs_final)
                    creados += c
                    no_encontrados += ne
                    ya_excluidos += ye

                if creados:
                    invalidar_cache_producto()

                if es_ajax:
                    return _respuesta_json_resumen(creados, no_encontrados, ya_excluidos)
                _mostrar_resumen(request, creados, no_encontrados, ya_excluidos)
            else:
                if es_ajax:
                    return JsonResponse({
                        'status': 'error',
                        'message': 'Formulario de Excel inválido.',
                        'errors': [str(e) for e in form_excel.errors.values()]
                    }, status=400)
                messages.error(request, 'Formulario de Excel inválido.')

        return redirect(next_url)

    form_manual = ExcluirProductosForm()
    form_excel = ExcluirProductosExcelForm()
    exclusiones_activas = ProductoExcluido.objects.filter(activo=True).select_related('producto')

    return render(request, 'excluir_productos.html', {
        'form_manual': form_manual,
        'form_excel': form_excel,
        'exclusiones_activas': exclusiones_activas,
    })


def _respuesta_json_resumen(creados, no_encontrados, ya_excluidos):
    """Arma la respuesta JSON para peticiones AJAX de exclusión."""
    partes = []
    if creados:
        partes.append(f'{len(creados)} producto(s) excluido(s): {", ".join(creados)}')
    if ya_excluidos:
        partes.append(f'{len(ya_excluidos)} ya estaban excluidos: {", ".join(ya_excluidos)}')
    if no_encontrados:
        partes.append(f'{len(no_encontrados)} código(s) no encontrados: {", ".join(no_encontrados)}')

    mensaje = ' | '.join(partes) if partes else 'No se procesó ningún código.'

    if no_encontrados and not creados:
        status = 'error'
    elif ya_excluidos and not creados:
        status = 'warning'
    else:
        status = 'success'

    return JsonResponse({'status': status, 'message': mensaje})


@transaction.atomic
def _excluir_codigos(codigos, observacion):
    creados, no_encontrados, ya_excluidos = [], [], []

    for codigo in codigos:
        try:
            producto = Producto.objects.get(codigo_mantis=codigo)
        except Producto.DoesNotExist:
            no_encontrados.append(codigo)
            continue

        exclusion_activa = ProductoExcluido.objects.filter(
            producto=producto, activo=True
        ).first()

        if exclusion_activa:
            ya_excluidos.append(codigo)
            continue

        ProductoExcluido.objects.create(
            producto=producto,
            activo=True,
            observacion=observacion or ''
        )

        # Marcar el producto como inactivo en la BD
        if producto.estado:
            producto.estado = False
            producto.save(update_fields=['estado'])

        creados.append(codigo)

    return creados, no_encontrados, ya_excluidos


def _mostrar_resumen(request, creados, no_encontrados, ya_excluidos):
    if creados:
        messages.success(request, f'{len(creados)} producto(s) excluido(s): {", ".join(creados)}')
    if ya_excluidos:
        messages.warning(request, f'{len(ya_excluidos)} ya estaban excluidos: {", ".join(ya_excluidos)}')
    if no_encontrados:
        messages.error(request, f'{len(no_encontrados)} código(s) no encontrados: {", ".join(no_encontrados)}')


#=======================================
# LISTAR PRODUCTOS EXCLUIDOS (con caché)
#=======================================
def productos_excluidos(request):
    exclusiones_activas = cache.get(CACHE_KEY_EXCLUIDOS)

    if exclusiones_activas is None:
        exclusiones_activas = list(
            ProductoExcluido.objects
            .filter(activo=True)
            .select_related('producto')
            .order_by('-fecha_creacion')
        )
        cache.set(CACHE_KEY_EXCLUIDOS, exclusiones_activas, CACHE_TIMEOUT)

    return render(request, 'excluidos.html', {
        'exclusiones_activas': exclusiones_activas,
        'form_manual': ExcluirProductosForm(),
        'form_excel': ExcluirProductosExcelForm(),
    })


#=======================================
# REACTIVAR (DESEXCLUIR) UN PRODUCTO (invalida caché)
#=======================================
def reactivar_producto_view(request, exclusion_id):
    es_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

    if request.method != 'POST':
        if es_ajax:
            return JsonResponse({'status': 'error', 'message': 'Método no permitido.'}, status=405)
        return redirect('productos_excluidos')

    ok = reactivar_producto(exclusion_id)

    if ok:
        invalidar_cache_producto()
        mensaje = 'Producto reincorporado al listado.'
        if es_ajax:
            return JsonResponse({'status': 'success', 'message': mensaje})
        messages.success(request, mensaje)
    else:
        mensaje = 'No se pudo reincorporar el producto (ya no estaba excluido).'
        if es_ajax:
            return JsonResponse({'status': 'error', 'message': mensaje}, status=400)
        messages.error(request, mensaje)

    return redirect('productos_excluidos')





# =======================================
# VIEW: LISTA DE COMPRAS (lectura con caché)
# =======================================
def compras_lista(request):
    """
    El caché aquí SOLO evita repetir la consulta a la BD en cada
    carga de página. Nunca se usa para decidir si algo se guardó:
    eso siempre pasa directo contra la base de datos en cargar_compras()
    (bulk_create / bulk_update dentro de procesar_archivo_compras).
 
    Igual que en las vistas de exclusión, la lectura del caché se
    protege con try/except: si el backend de caché falla por lo que
    sea, la vista sigue funcionando yendo directo a la BD en vez de
    romperse.
    """
    try:
        compras = cache.get(CACHE_KEY_COMPRAS)
    except Exception:
        logger.exception("Fallo al leer caché de compras; se consulta la BD directamente.")
        compras = None
 
    if compras is None:
        compras = list(
            Compra.objects.select_related('producto').order_by('-fecha_compra')
        )
        try:
            cache.set(CACHE_KEY_COMPRAS, compras, CACHE_TIMEOUT)
        except Exception:
            logger.exception("Fallo al guardar caché de compras; los datos igual se muestran desde la BD.")
 
    return render(request, 'detalles_compra.html', {
        'compras': compras,
        'form': CargarComprasForm(),
    })
 
 
# =======================================
# VIEW: CARGA DE EXCEL DE COMPRAS (escritura directa a la BD)
# =======================================
def cargar_compras(request):
    if request.method != 'POST':
        return JsonResponse({'status': 'error', 'message': 'Método no permitido.'}, status=405)

    form = CargarComprasForm(request.POST, request.FILES)
    if not form.is_valid():
        return JsonResponse({
            'status': 'error',
            'message': f'Formulario inválido: {form.errors.as_text()}'
        }, status=400)

    archivo = request.FILES['archivo_excel']
    logger.info(f"Archivo de compras recibido: {archivo.name} ({archivo.size} bytes)")

    try:
        creados, actualizados, productos_creados = procesar_archivo_compras(archivo)

        if creados or actualizados:
            invalidar_cache_compra()
        if productos_creados:
            invalidar_cache_producto()

        partes = []
        if creados:
            partes.append(f'{creados} compra(s) nueva(s) cargada(s).')
        if actualizados:
            partes.append(f'{actualizados} compra(s) actualizada(s).')
        if productos_creados:
            partes.append(f'{len(productos_creados)} producto(s) nuevo(s) creado(s).')
        if not creados and not actualizados:
            partes.append('No se encontraron filas válidas para cargar (revisa la columna D desde la fila 6).')

        return JsonResponse({'status': 'success', 'message': ' '.join(partes)})

    except Exception as e:
        logger.exception("Fallo al procesar archivo de compras")
        return JsonResponse({'status': 'error', 'message': f'Error al procesar el archivo: {str(e)}'}, status=500)


#=======================================
# PRODUCTOS CON DURACIÓN MENOR A 1.0
#=======================================
@never_cache
def productos_bajo_duracion(request):

    # ══════════════════════════════════════════════════════════════════
    # ENDPOINT JSON PARA POWER AUTOMATE — ?api=true envio de alerta
    # ══════════════════════════════════════════════════════════════════
    if request.GET.get("api") == "true":
        fecha_referencia, dia_de_corte, alertas = construir_alertas_baja_duracion(umbral=1.0, solo_categoria_a=True)
        resp = JsonResponse({
            "fecha_referencia": fecha_referencia,
            "dia_de_corte":     dia_de_corte,
            "total_alertas":    len(alertas),
            "alertas":          alertas,
        })
        resp["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp["Pragma"]        = "no-cache"
        resp["Expires"]       = "0"
        return resp

    # ══════════════════════════════════════════════════════════════════
    # VISTA HTML NORMAL
    # ══════════════════════════════════════════════════════════════════
    productos = obtener_productos_bajo_duracion(umbral=1.0)

    if productos:
        ids = [item['detalle'].producto_id for item in productos]
        categorizaciones = Categorizacion.objects.filter(producto_id__in=ids)
        mapa_cat = {c.producto_id: c for c in categorizaciones}
        for item in productos:
            item['categorizacion'] = mapa_cat.get(item['detalle'].producto_id)

        # ── Orden: Categoría A > B > C > Sin categoría ──
        # Dentro de cada categoría: duración ascendente, salvo que
        # duracion == 0 Y solicitar == 0 (nada urgente que pedir),
        # que se manda al final de su categoría.
        ORDEN_CATEGORIA = {'A': 0, 'B': 1, 'C': 2}

        def clave_orden(item):
            detalle = item['detalle']
            cat = item.get('categorizacion')
            orden_cat = ORDEN_CATEGORIA.get(cat.tipo_categoria, 3) if cat else 3

            sin_urgencia = (detalle.duracion == 0 and detalle.solicitar == 0)
            # False (0) ordena antes que True (1)
            return (orden_cat, sin_urgencia, detalle.duracion)

        productos = sorted(productos, key=clave_orden)

    return render(request, 'productos_comprar.html', {
        'productos': productos,
    })

# para generar y descargar un reporte Excel de productos con baja duración
@require_GET
def descargar_reporte_bajo_stock(request):
    """
    Genera y descarga un archivo Excel con el detalle de productos
    con baja duración, incluyendo todas las compras históricas.
    """
    productos = obtener_productos_bajo_duracion(umbral=1.0)

    if productos:
        ids = [item['detalle'].producto_id for item in productos]
        categorizaciones = Categorizacion.objects.filter(producto_id__in=ids)
        mapa_cat = {c.producto_id: c for c in categorizaciones}
        for item in productos:
            item['categorizacion'] = mapa_cat.get(item['detalle'].producto_id)

    return generar_excel_reporte(productos, titulo="Productos con baja duración")
#====================================================
# VIEW PARA LA CARGA DE CATEGORIA DE PRODUCTOS 
#====================================================

def categorizacion_view(request):

    if request.method == "POST":
        es_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        codigo = request.POST.get("codigo")
        categoria = request.POST.get("categoria")
        analisis = request.POST.get("analisis")

        try:
            producto = Producto.objects.get(codigo_mantis=codigo.strip())
            Categorizacion.objects.update_or_create(
                producto=producto,
                defaults={"tipo_categoria": categoria, "analisis": analisis}
            )
            if es_ajax:
                return JsonResponse({"status": "success", "message": "Registro guardado correctamente."})
            messages.success(request, "Registro guardado correctamente.")

        except Producto.DoesNotExist:
            if es_ajax:
                return JsonResponse({"status": "error", "message": "No existe un producto con ese código."}, status=400)
            messages.error(request, "No existe un producto con ese código.")

        return redirect("categorizacion")

    # ── Orden: Categoría A primero, luego B, luego C, luego sin categoría ──
    datos = (
        Categorizacion.objects
        .select_related("producto")
        .annotate(
            orden_categoria=Case(
                When(tipo_categoria='A', then=Value(0)),
                When(tipo_categoria='B', then=Value(1)),
                When(tipo_categoria='C', then=Value(2)),
                default=Value(3),
                output_field=IntegerField(),
            )
        )
        .order_by("orden_categoria", "-fecha_registro")
    )

    return render(request, "categorizacion.html", {"datos": datos})


def cargar_excel_categorizacion(request):
    if request.method == "POST":
        es_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        archivo = request.FILES.get("archivo")

        if archivo:
            mensaje = importar_categorizacion_excel(archivo)
            if es_ajax:
                return JsonResponse({"status": "success", "message": mensaje})
            messages.success(request, mensaje)
        else:
            if es_ajax:
                return JsonResponse({"status": "error", "message": "No se recibió ningún archivo."}, status=400)
            messages.error(request, "No se recibió ningún archivo.")

    return redirect("categorizacion")

# VIEW PARA CARGAR FECHAS DE VENCIMIENTO
#====================================================
# VIEW: LISTA DE VENCIMIENTOS
#====================================================
def lista_vencimientos(request):
    datos = (
        Vencimiento.objects
        .select_related("producto")
        .order_by("fecha_vencimiento")
    )
    return render(request, "vencimientos.html", {"datos": datos})

#====================================================
# VIEW: CARGA DE EXCEL DE VENCIMIENTOS
#====================================================
def cargar_excel_vencimientos(request):
    if request.method == "POST":
        es_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        archivo = request.FILES.get("archivo")
        tipo_archivo = request.POST.get("tipo_archivo")

        if not archivo:
            if es_ajax:
                return JsonResponse({"status": "error", "message": "No se recibió ningún archivo."}, status=400)
            messages.error(request, "No se recibió ningún archivo.")
            return redirect("lista_vencimientos")

        if not archivo.name.lower().endswith((".xlsx", ".xls")):
            msg = "Formato de archivo no válido. Solo se aceptan archivos .xlsx o .xls."
            if es_ajax:
                return JsonResponse({"status": "error", "message": msg}, status=400)
            messages.error(request, msg)
            return redirect("lista_vencimientos")

        if tipo_archivo not in ("detallado_compras", "fecha_vencimiento"):
            if es_ajax:
                return JsonResponse({"status": "error", "message": "Debes seleccionar el tipo de archivo."}, status=400)
            messages.error(request, "Debes seleccionar el tipo de archivo.")
            return redirect("lista_vencimientos")

        mensaje, no_encontrados = importar_vencimientos_excel(archivo, tipo_archivo)

        if es_ajax:
            return JsonResponse({
                "status": "success",
                "message": mensaje,
                "no_encontrados": no_encontrados,
            })

        messages.success(request, mensaje)
        return redirect("lista_vencimientos")

    return redirect("lista_vencimientos")
