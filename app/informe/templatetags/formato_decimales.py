from django import template

register = template.Library()

@register.filter
def sin_ceros(valor, decimales=2):
    """
    Muestra el número con hasta 'decimales' decimales,
    pero sin ceros innecesarios si es un número entero.
    Ej: 23.000 -> 23   |   23.810 -> 23.81
    """
    try:
        decimales = int(decimales)
        numero = float(valor)
    except (TypeError, ValueError):
        return valor

    formateado = f"{numero:.{decimales}f}"
    if '.' in formateado:
        formateado = formateado.rstrip('0').rstrip('.')
    return formateado