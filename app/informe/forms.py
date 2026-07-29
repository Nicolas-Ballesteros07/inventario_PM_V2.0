from django import forms
import os

# FORMULARIO PARA SUBIR LOS ARCHIVOS DEL INFORME
class cargar_archivos_informe(forms.Form):
    archivo_listado_basico = forms.FileField(
        label='Listado Básico (Excel)',
        help_text='Formatos aceptados: .xls, .xlsx'
    )
    archivo_ventas = forms.FileField(
        label='Ventas (Excel)',
        help_text='Formatos aceptados: .xls, .xlsx'
    )

    def clean_archivo_listado_basico(self):
        archivo = self.cleaned_data['archivo_listado_basico']
        ext = os.path.splitext(archivo.name)[1].lower()
        if ext not in ['.xls', '.xlsx']:
            raise forms.ValidationError('Solo se permiten archivos .xls o .xlsx')
        return archivo

    def clean_archivo_ventas(self):
        archivo = self.cleaned_data['archivo_ventas']
        ext = os.path.splitext(archivo.name)[1].lower()
        if ext not in ['.xls', '.xlsx']:
            raise forms.ValidationError('Solo se permiten archivos .xls o .xlsx')
        return archivo

  
# FORMOS PARA FILTRAR PRODUCTOS EXCLUIDOS
class ExcluirProductosForm(forms.Form):
    """
    Permite excluir uno o varios productos escribiendo los códigos
    separados por coma, espacio o salto de línea, con una sola observación.
    """
    codigos = forms.CharField(
        label='Código(s) Mantis',
        widget=forms.Textarea(attrs={
            'rows': 4,
            'placeholder': 'Ej: MD00002, MD00007\nMD00027\nMD00028'
        }),
        help_text='Puedes escribir uno o varios códigos separados por coma, espacio o salto de línea.'
    )
    observacion = forms.CharField(
        label='Observación',
        max_length=255,
        required=False,
        widget=forms.TextInput(attrs={'placeholder': 'Ej: NO VOLVER A COMPRAR'})
    )

    def limpiar_codigos(self):
        """Devuelve una lista de códigos únicos, limpios y en mayúscula."""
        raw = self.cleaned_data['codigos']
        # separa por coma, espacio o salto de línea
        import re
        partes = re.split(r'[,\s]+', raw.strip())
        codigos = [p.strip().upper() for p in partes if p.strip()]
        # elimina duplicados manteniendo el orden
        vistos = set()
        unicos = []
        for c in codigos:
            if c not in vistos:
                vistos.add(c)
                unicos.append(c)
        return unicos

# PARA CARGAR PRODUCTOS EXCLUIDOS DESDE UN EXCEL
class ExcluirProductosExcelForm(forms.Form):
    """
    Permite excluir productos masivamente subiendo un Excel.
    Se espera una columna con el código (REFERENCIA / CODIGO / CODIGO MANTIS)
    y opcionalmente una columna de observación (EXISTENCIA / OBSERVACION).
    """
    archivo_excel = forms.FileField(
        label='Archivo Excel (.xls o .xlsx)',
        help_text='Debe tener una columna con el código del producto.'
    )
    observacion_general = forms.CharField(
        label='Observación (si el Excel no trae una columna de observación)',
        max_length=255,
        required=False,
        widget=forms.TextInput(attrs={'placeholder': 'Ej: NO VOLVER A COMPRAR'})
    )

    def clean_archivo_excel(self):
        archivo = self.cleaned_data['archivo_excel']
        ext = os.path.splitext(archivo.name)[1].lower()
        if ext not in ['.xls', '.xlsx']:
            raise forms.ValidationError('Solo se permiten archivos .xls o .xlsx')
        return archivo

# CARGAR EL ARCHIVO DE COMPRA
class CargarComprasForm(forms.Form):
    archivo_excel = forms.FileField(label='Archivo de Compras (.xlsx)')