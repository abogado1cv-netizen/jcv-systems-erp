import csv
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from django import forms 
from django.urls import path
from django.contrib import admin
from django.http import HttpResponse
from django.utils.html import format_html
from import_export.admin import ImportExportModelAdmin
from import_export import resources, fields
from import_export.widgets import ForeignKeyWidget, DateWidget
from django.contrib import messages
from django.shortcuts import redirect, render
from django.core.mail import send_mail
from django.conf import settings
from django.utils import timezone 
from django.utils.safestring import mark_safe
from .models import ConfiguracionEmail
from django.core.mail import get_connection
from django.core.mail import EmailMultiAlternatives

from .models import (
    OrdenCompra, PartidaCompra, Inventario, DocumentoOrdenCompra,
    CatalogoMedicamento, EstatusProcedimiento, Empresa, SocioComercial,
    Licitacion, PartidaRequerimiento, RegistroUbicacion, Contrato, 
    FianzaContrato, ClaveContrato, OrdenSuministro, RemisionEntrega,
    EntradaAlmacen, Almacen, TraspasoIntercompany
    
)

@admin.action(description='📥 Descargar seleccionados a Excel (CSV)')
def exportar_a_csv(modeladmin, request, queryset):
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="reporte_erp.csv"'
    response.write(u'\ufeff'.encode('utf8')) 
    writer = csv.writer(response)
    campos = [field.name for field in modeladmin.model._meta.fields]
    writer.writerow(campos)
    for objeto in queryset:
        writer.writerow([getattr(objeto, campo) for campo in campos])
    return response

# ==========================================
# --- ANÁLISIS COMERCIAL (EXCEL REAL .XLSX) ---
# ==========================================
@admin.action(description='📊 Descargar Análisis Comercial (Ejecutivo / Sábana) (Excel)')
def exportar_analisis_licitacion(modeladmin, request, queryset):
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    es_individual = queryset.count() == 1
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Análisis Comercial"

    font_cabecera = Font(bold=True, color="000000")
    fill_cabecera = PatternFill("solid", fgColor="D9D9D9")
    fill_datos_grales = PatternFill("solid", fgColor="EFEFEF")
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alineacion_izq = Alignment(horizontal="left", vertical="center", wrap_text=True)
    borde_delgado = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    formato_moneda = '"$"#,##0.00'

    if es_individual:
        licitacion = queryset.first()
        response['Content-Disposition'] = f'attachment; filename="Análisis_Ejecutivo_{licitacion.num_procedimiento}.xlsx"'
        
        f_pub = licitacion.fecha_publicacion.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_publicacion else 'S/D'
        f_ape = licitacion.fecha_apertura.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_apertura else 'S/D'
        f_jun = licitacion.fecha_junta.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_junta else 'S/D'
        f_fal = licitacion.fecha_fallo.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_fallo else 'S/D'
        estatus = licitacion.estatus.estado if licitacion.estatus else 'S/D'

        ws.append(['', 'NÚMERO DE EXPEDIENTE', 'DEPENDENCIA', 'ESTATUS'])
        ws.append(['', licitacion.num_procedimiento, licitacion.dependencia, estatus])
        ws.append([])
        ws.append(['', 'FECHA DE PUBLICACIÓN', 'JUNTA DE ACLARACIONES', 'APERTURA DE PROPUESTAS', 'ACTO DE FALLO'])
        ws.append(['', f_pub, f_jun, f_ape, f_fal])
        ws.append([])
        ws.append([])

        for r in [1, 4]:
            for c in [2, 3, 4, 5]:
                celda = ws.cell(row=r, column=c)
                celda.font = font_cabecera
                celda.fill = fill_datos_grales
                celda.alignment = alineacion_centro
                celda.border = borde_delgado
        for r in [2, 5]:
            for c in [2, 3, 4, 5]:
                celda = ws.cell(row=r, column=c)
                celda.alignment = alineacion_centro
                celda.border = borde_delgado
        
        encabezados = ['PARTIDA', 'CLAVE', 'CANTIDAD MÍNIMA', 'CANTIDAD MÁXIMA', 'PRECIO UNITARIO', 'IMPORTE MÁXIMO', 'RESULTADO', 'FABRICANTE', 'MARCA', 'GENÉRICA']
        ws.append(encabezados)
        fila_encabezados = ws.max_row
        
        for col_num, header in enumerate(encabezados, 1):
            celda = ws.cell(row=fila_encabezados, column=col_num)
            celda.font = font_cabecera
            celda.fill = fill_cabecera
            celda.alignment = alineacion_centro
            celda.border = borde_delgado
        
        for partida in licitacion.partidas.all():
            med = partida.medicamento
            clave = med.clave_sector if med else 'S/C'
            cant_min = partida.cantidad_minima or 0
            cant_max = partida.cantidad_maxima or 0
            precio = float(partida.precio or 0)
            importe_max = cant_max * precio
            resultado_txt = partida.resultado if hasattr(partida, 'resultado') else 'Pendiente'
            
            def escribir_fila(datos):
                ws.append(datos)
                r = ws.max_row
                for c in range(1, 11):
                    celda = ws.cell(row=r, column=c)
                    celda.border = borde_delgado
                    celda.alignment = alineacion_centro
                    if c in [5, 6]:
                        celda.number_format = formato_moneda
                    if c in [8, 9, 10]:
                        celda.alignment = alineacion_izq
            
            if med and clave != 'S/C':
                medicamentos_similares = CatalogoMedicamento.objects.filter(clave_sector=clave)
                if medicamentos_similares.exists():
                    for m in medicamentos_similares:
                        fab = m.fabricante if m.fabricante else 'Sin Laboratorio'
                        marca = m.denominacion_distintiva if m.denominacion_distintiva else 'Genérico'
                        generica = m.denominacion_generica if m.denominacion_generica else 'S/D'
                        escribir_fila([partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, fab, marca, generica])
                else:
                    escribir_fila([partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, 'SIN REGISTRO', 'S/D', 'S/D'])
            else:
                escribir_fila([partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, 'SIN CLAVE', 'S/D', 'S/D'])
        
        ws.column_dimensions['A'].width = 12 
        ws.column_dimensions['B'].width = 20 
        ws.column_dimensions['C'].width = 20 
        ws.column_dimensions['D'].width = 20 
        ws.column_dimensions['E'].width = 20 
        ws.column_dimensions['F'].width = 22 
        ws.column_dimensions['G'].width = 20 
        ws.column_dimensions['H'].width = 35 
        ws.column_dimensions['I'].width = 25 
        ws.column_dimensions['J'].width = 40 
        
    else:
        response['Content-Disposition'] = 'attachment; filename="Sabana_Comercial_Global.xlsx"'
        encabezados = ['EVENTO', 'DEPENDENCIA', 'FECHA DE PUBLICACIÓN', 'FECHA DE JUNTA', 'FECHA DE APERTURA', 'FECHA DE FALLO', 'PARTIDA', 'CLAVE', 'CANTIDAD MÍNIMA', 'CANTIDAD MÁXIMA', 'PRECIO UNITARIO', 'IMPORTE MÁXIMO', 'RESULTADO', 'LABORATORIO FABRICANTE', 'MARCA (DISTINTIVA)', 'SUSTANCIA (GENÉRICA)']
        ws.append(encabezados)
        
        for col_num, header in enumerate(encabezados, 1):
            celda = ws.cell(row=1, column=col_num)
            celda.font = font_cabecera
            celda.fill = fill_cabecera
            celda.alignment = alineacion_centro
            celda.border = borde_delgado
            
        for licitacion in queryset:
            evento = licitacion.num_procedimiento
            dependencia = licitacion.dependencia
            f_pub = licitacion.fecha_publicacion.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_publicacion else 'S/D'
            f_ape = licitacion.fecha_apertura.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_apertura else 'S/D'
            f_jun = licitacion.fecha_junta.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_junta else 'S/D'
            f_fal = licitacion.fecha_fallo.strftime('%d/%m/%Y %H:%M') if licitacion.fecha_fallo else 'S/D'
            
            for partida in licitacion.partidas.all():
                med = partida.medicamento
                clave = med.clave_sector if med else 'S/C'
                cant_min = partida.cantidad_minima or 0
                cant_max = partida.cantidad_maxima or 0
                precio = float(partida.precio or 0)
                importe_max = cant_max * precio
                resultado_txt = partida.resultado if hasattr(partida, 'resultado') else 'Pendiente'
                
                def escribir_fila_sabana(datos):
                    ws.append(datos)
                    r = ws.max_row
                    for c in range(1, 17):
                        celda = ws.cell(row=r, column=c)
                        celda.border = borde_delgado
                        celda.alignment = alineacion_centro
                        if c in [11, 12]:
                            celda.number_format = formato_moneda
                        if c in [14, 15, 16]:
                            celda.alignment = alineacion_izq
                            
                if med and clave != 'S/C':
                    medicamentos_similares = CatalogoMedicamento.objects.filter(clave_sector=clave)
                    if medicamentos_similares.exists():
                        for m in medicamentos_similares:
                            fab = m.fabricante if m.fabricante else 'Sin Laboratorio'
                            marca = m.denominacion_distintiva if m.denominacion_distintiva else 'Genérico'
                            generica = m.denominacion_generica if m.denominacion_generica else 'S/D'
                            escribir_fila_sabana([evento, dependencia, f_pub, f_jun, f_ape, f_fal, partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, fab, marca, generica])
                    else:
                        escribir_fila_sabana([evento, dependencia, f_pub, f_jun, f_ape, f_fal, partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, 'SIN REGISTRO', 'S/D', 'S/D'])
                else:
                    escribir_fila_sabana([evento, dependencia, f_pub, f_jun, f_ape, f_fal, partida.numero_partida, clave, cant_min, cant_max, precio, importe_max, resultado_txt, 'SIN CLAVE', 'S/D', 'S/D'])

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 20
        ws.column_dimensions['G'].width = 10
        ws.column_dimensions['H'].width = 20
        ws.column_dimensions['I'].width = 15
        ws.column_dimensions['J'].width = 15
        ws.column_dimensions['K'].width = 18
        ws.column_dimensions['L'].width = 18
        ws.column_dimensions['M'].width = 20
        ws.column_dimensions['N'].width = 35
        ws.column_dimensions['O'].width = 25
        ws.column_dimensions['P'].width = 40

    wb.save(response)
    return response

# ==========================================
# --- REPORTE POR SOCIO COMERCIAL (EXCEL REAL .XLSX) ---
# ==========================================
@admin.action(description='🏭 Descargar Reporte Agrupado por Socios Comerciales (Excel)')
def exportar_reporte_laboratorios(modeladmin, request, queryset):
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if queryset.count() == 1:
        num_evento = queryset.first().num_procedimiento
        response['Content-Disposition'] = f'attachment; filename="Reporte_Socios_{num_evento}.xlsx"'
    else:
        response['Content-Disposition'] = 'attachment; filename="Reporte_Socios_Global.xlsx"'
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Análisis Comercial"

    font_cabecera = Font(bold=True, color="000000")
    fill_cabecera = PatternFill("solid", fgColor="D9D9D9")
    font_subtotal = Font(bold=True, italic=True)
    fill_subtotal = PatternFill("solid", fgColor="F2F2F2")
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alineacion_izq = Alignment(horizontal="left", vertical="center", wrap_text=True)
    borde_delgado = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    formato_moneda = '"$"#,##0.00'

    encabezados = ['Socio Comercial', 'Partida', 'Clave', 'Descripción', 'Cantidad solicitada', '# de apoyos', 'Cantidad de piezas', 'Ubicación del Inventario', '$ referencia', 'Importe']
    ws.append(encabezados)
    
    for col_num, header in enumerate(encabezados, 1):
        celda = ws.cell(row=1, column=col_num)
        celda.font = font_cabecera
        celda.fill = fill_cabecera
        celda.alignment = alineacion_centro
        celda.border = borde_delgado

    filas = []
    for licitacion in queryset:
        for partida in licitacion.partidas.all():
            med = partida.medicamento
            clave = med.clave_sector if med else 'S/C'
            cant_max = partida.cantidad_maxima or 0
            precio = float(partida.precio or 0)
            importe_max = cant_max * precio
            
            total_piezas = 0
            texto_ubicaciones = "Sin Inventario"
            
            if med and clave != 'S/C':
                stock_disponible = Inventario.objects.filter(
                    medicamento__clave_sector=clave, 
                    cantidad_disponible__gt=0
                ).select_related('almacen')
                
                desglose_almacen = []
                for item in stock_disponible:
                    total_piezas += item.cantidad_disponible
                    nombre_alm = item.almacen.nombre if item.almacen else "Bodega Central"
                    desglose_almacen.append(f"{nombre_alm} ({item.cantidad_disponible})")
                
                if desglose_almacen:
                    texto_ubicaciones = " | ".join(desglose_almacen)

                medicamentos_similares = CatalogoMedicamento.objects.filter(clave_sector=clave)
                if medicamentos_similares.exists():
                    for m in medicamentos_similares:
                        if m.socio_contacto:
                            nombre_socio = m.socio_contacto.nombre
                        else:
                            nombre_socio = 'SIN SOCIO ASIGNADO'
                            
                        desc = m.denominacion_generica if m.denominacion_generica else 'S/D'
                        if m.denominacion_distintiva:
                            desc += f" / Marca: {m.denominacion_distintiva}"
                        
                        filas.append([nombre_socio, partida.numero_partida, clave, desc, cant_max, 0, total_piezas, texto_ubicaciones, precio, importe_max])
                else:
                    filas.append(['SIN REGISTRO', partida.numero_partida, clave, 'S/D', cant_max, 0, total_piezas, texto_ubicaciones, precio, importe_max])
            else:
                filas.append(['SIN CLAVE', partida.numero_partida, clave, 'S/D', cant_max, 0, 0, 'N/A', precio, importe_max])
                
    filas.sort(key=lambda x: (str(x[0]), int(x[1]) if x[1] not in [None, ''] else 0))
    
    current_fab = None
    subtotal_importe = 0
    
    for fila in filas:
        fab = fila[0]
        if current_fab != fab:
            if current_fab is not None:
                ws.append(['', f'Total {current_fab}', '', '', '', '', '', '', '', subtotal_importe])
                fila_subtotal = ws.max_row
                for c in range(1, 11): 
                    ws.cell(row=fila_subtotal, column=c).fill = fill_subtotal
                    ws.cell(row=fila_subtotal, column=c).font = font_subtotal
                    ws.cell(row=fila_subtotal, column=c).border = borde_delgado
                ws.cell(row=fila_subtotal, column=10).number_format = formato_moneda
            
            current_fab = fab
            subtotal_importe = 0
            mostrar_fab = fab
        else:
            mostrar_fab = ''
            
        subtotal_importe += fila[9] 
        
        ws.append([mostrar_fab, fila[1], fila[2], fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9]])
        fila_actual = ws.max_row
        
        for c in range(1, 11):
            celda = ws.cell(row=fila_actual, column=c)
            celda.border = borde_delgado
            if c in [9, 10]: 
                celda.number_format = formato_moneda
            if c in [2, 5, 6, 7]: 
                celda.alignment = alineacion_centro
            if c in [4, 8]: 
                celda.alignment = alineacion_izq

    if current_fab is not None:
        ws.append(['', f'Total {current_fab}', '', '', '', '', '', '', '', subtotal_importe])
        fila_subtotal = ws.max_row
        for c in range(1, 11):
            ws.cell(row=fila_subtotal, column=c).fill = fill_subtotal
            ws.cell(row=fila_subtotal, column=c).font = font_subtotal
            ws.cell(row=fila_subtotal, column=c).border = borde_delgado
        ws.cell(row=fila_subtotal, column=10).number_format = formato_moneda

    ws.column_dimensions['A'].width = 35 
    ws.column_dimensions['B'].width = 10 
    ws.column_dimensions['C'].width = 18 
    ws.column_dimensions['D'].width = 45 
    ws.column_dimensions['E'].width = 15 
    ws.column_dimensions['F'].width = 12 
    ws.column_dimensions['G'].width = 18 
    ws.column_dimensions['H'].width = 35 
    ws.column_dimensions['I'].width = 15 
    ws.column_dimensions['J'].width = 18 

    wb.save(response)
    return response

# ==========================================
# --- REPORTE DE INVENTARIO "CHULO" ---
# ==========================================
@admin.action(description='📊 Descargar Inventario Ejecutivo (Excel Especial)')
def exportar_inventario_personalizado(modeladmin, request, queryset):
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="Inventario_Detallado_Gpharma.xlsx"'
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventario Actual"

    # Estilos (usando el mismo diseño de tus otros reportes)
    font_cabecera = Font(bold=True, color="000000")
    fill_cabecera = PatternFill("solid", fgColor="D9D9D9")
    alineacion_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alineacion_izq = Alignment(horizontal="left", vertical="center", wrap_text=True)
    borde_delgado = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Encabezados solicitados
    encabezados = [
        'ALMACÉN', 'CLAVE SECTOR', 'DESCRIPCIÓN', 'SOCIO COMERCIAL', 
        'FABRICANTE', 'TIPO', 'LOTE', 'CADUCIDAD', 'EXISTENCIA'
    ]
    ws.append(encabezados)

    for col_num, header in enumerate(encabezados, 1):
        celda = ws.cell(row=1, column=col_num)
        celda.font = font_cabecera
        celda.fill = fill_cabecera
        celda.alignment = alineacion_centro
        celda.border = borde_delgado

    # Datos
    for item in queryset:
        med = item.medicamento
        # Traemos al socio comercial desde el catálogo
        socio = med.socio_contacto.nombre if med.socio_contacto else "SIN ASIGNAR"
        
        fila = [
            item.almacen.nombre if item.almacen else "BODEGA CENTRAL",
            med.clave_sector,
            med.denominacion_generica,
            socio,
            med.fabricante,
            item.get_tipo_producto_display(),
            item.lote,
            item.fecha_caducidad.strftime('%d/%m/%Y') if item.fecha_caducidad else "S/F",
            item.cantidad_disponible
        ]
        ws.append(fila)
        
        # Aplicar bordes y formato a la fila
        r = ws.max_row
        for c in range(1, 10):
            celda = ws.cell(row=r, column=c)
            celda.border = borde_delgado
            celda.alignment = alineacion_centro if c != 3 else alineacion_izq

    # Ajustar anchos de columnas
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 12

    wb.save(response)
    return response

# ==========================================
# --- CLASES DEL ADMIN ---
# ==========================================
class SocioComercialWidget(ForeignKeyWidget):
    def clean(self, value, row=None, **kwargs):
        if not value or str(value).strip() in ['#N/A', 'N/A', 'NA', '']:
            return None
        nombre_socio = str(value).strip()
        socio, created = self.model.objects.get_or_create(
            nombre=nombre_socio,
            defaults={'correos': 'pendiente@actualizar.com'}
        )
        return socio

class CatalogoMedicamentoResource(resources.ModelResource):
    socio_contacto = fields.Field(
        column_name='socio_contacto',
        attribute='socio_contacto',
        widget=SocioComercialWidget(SocioComercial, field='nombre')
    )
    fecha_expedicion = fields.Field(
        column_name='fecha_expedicion',
        attribute='fecha_expedicion',
        widget=DateWidget(format='%Y-%m-%d')
    )
    fecha_vigencia = fields.Field(
        column_name='fecha_vigencia',
        attribute='fecha_vigencia',
        widget=DateWidget(format='%Y-%m-%d')
    )

    class Meta:
        model = CatalogoMedicamento
        fields = (
            'id', 'clave_sector', 'descripcion', 'denominacion_generica', 
            'denominacion_distintiva', 'fabricante', 'socio_contacto', 
            'rfc_fabricante', 'pais_fabricacion', 'num_registro_sanitario', 
            'num_prorroga', 'codigo_barras', 'fecha_expedicion', 'fecha_vigencia'
        )
        
    def before_import_row(self, row, **kwargs):
        if str(row.get('fecha_expedicion', '')).strip() in ['None', '#N/A', 'N/A', 'NA', '']:
            row['fecha_expedicion'] = None
        if str(row.get('fecha_vigencia', '')).strip() in ['None', '#N/A', 'N/A', 'NA', '']:
            row['fecha_vigencia'] = None

class CatalogoMedicamentoAdmin(ImportExportModelAdmin): 
    resource_class = CatalogoMedicamentoResource
    list_per_page = 30
    search_fields = ['clave_sector', 'denominacion_distintiva', 'denominacion_generica']
    list_display = ('clave_sector', 'denominacion_generica', 'denominacion_distintiva', 'fabricante', 'registro_sanitario_formato')

    def registro_sanitario_formato(self, obj):
        return obj.num_registro_sanitario
    registro_sanitario_formato.short_description = "Registro Sanitario"

class RegistroUbicacionAdmin(admin.ModelAdmin):
    list_per_page = 30
    list_display = ('usuario', 'fecha_hora', 'latitud', 'longitud', 'ver_mapa')
    list_filter = ('usuario', 'fecha_hora')
    actions = [exportar_a_csv] 

    def ver_mapa(self, obj):
        if obj.latitud and obj.longitud:
            url = f"https://www.google.com/maps?q={obj.latitud},{obj.longitud}"
            return format_html('<a href="{}" target="_blank" style="background-color: #28a745; color: white; padding: 5px 10px; border-radius: 4px; text-decoration: none;">📍 Ver en Mapa</a>', url)
        return "Sin ubicación"
    ver_mapa.short_description = "Google Maps"

@admin.register(Almacen)
class AlmacenAdmin(admin.ModelAdmin):
    list_display = ('nombre', 'descripcion')
    search_fields = ('nombre',)

class LicitacionForm(forms.ModelForm):
    pegar_excel = forms.CharField(label="📥 Carga Masiva (Pegar desde Excel)", required=False, widget=forms.Textarea(attrs={'rows': 6, 'placeholder': 'Ejemplo:\n1\t010.000.4434.00\tPARACETAMOL 500MG\t500\t1000'}), help_text="Copia de Excel 4 o 5 columnas juntas: PARTIDA | CLAVE | DESCRIPCIÓN | (OPCIONAL: PIEZAS MIN) | PIEZAS MAX.")
    fecha_publicacion = forms.DateTimeField(label="Fecha y hora de publicación", required=False, widget=forms.DateTimeInput(format='%Y-%m-%dT%H:%M', attrs={'type': 'datetime-local'}))
    fecha_apertura = forms.DateTimeField(label="Fecha y hora de apertura", required=False, widget=forms.DateTimeInput(format='%Y-%m-%dT%H:%M', attrs={'type': 'datetime-local'}))
    fecha_junta = forms.DateTimeField(label="Fecha y hora de junta de aclaraciones", required=False, widget=forms.DateTimeInput(format='%Y-%m-%dT%H:%M', attrs={'type': 'datetime-local'}))
    fecha_fallo = forms.DateTimeField(label="Fecha y hora del acto del Fallo", required=False, widget=forms.DateTimeInput(format='%Y-%m-%dT%H:%M', attrs={'type': 'datetime-local'}))
    class Meta:
        model = Licitacion
        fields = '__all__'

class PartidaRequerimientoInline(admin.TabularInline):
    model = PartidaRequerimiento
    extra = 0 
    autocomplete_fields = ['medicamento'] 
    fields = ('numero_partida', 'medicamento', 'cantidad_maxima', 'precio', 'resultado', 'motivo_perdida')
    ordering = ['numero_partida']

    def get_readonly_fields(self, request, obj=None):
        if not request.user.is_superuser:
            return ('precio', 'cantidad_maxima', 'licitacion', 'medicamento')
        return super().get_readonly_fields(request, obj)

@admin.register(Licitacion)
class LicitacionAdmin(admin.ModelAdmin):
    list_per_page = 30
    form = LicitacionForm 
    list_display = ('num_procedimiento', 'dependencia', 'estatus_color', 'apertura_semaforo', 'fecha_fallo')
    search_fields = ['num_procedimiento', 'dependencia']
    inlines = [PartidaRequerimientoInline]
    
    actions = [exportar_a_csv, exportar_analisis_licitacion, exportar_reporte_laboratorios]
    
    change_form_template = "admin/licitaciones/licitacion/change_form.html"

    def estatus_color(self, obj):
        if not obj.estatus: return "-"
        estado = obj.estatus.estado.upper()
        if estado == 'EN_PROCESO': color_fondo = '#3498db' 
        elif estado == 'ADJUDICADO': color_fondo = '#2ecc71' 
        elif estado == 'PERDIDO': color_fondo = '#e74c3c' 
        else: color_fondo = '#95a5a6' 
        return format_html('<span style="color: white; background-color: {}; padding: 4px 8px; border-radius: 6px; font-weight: bold; font-size: 11px;">{}</span>', color_fondo, estado)
    estatus_color.short_description = 'Estatus'

    def apertura_semaforo(self, obj):
        if not obj.fecha_apertura: 
            return format_html('<span style="color: #aaa;">{}</span>', 'Sin fecha asignada')
            
        hoy = timezone.now().date()
        dias_faltantes = (obj.fecha_apertura.date() - hoy).days
        
        if dias_faltantes < 0:
            return format_html('<span style="color: gray;"><b>Ya pasó</b> (hace {} días)</span>', abs(dias_faltantes))
        elif dias_faltantes <= 3:
            return format_html('<span style="color: red;"><b>Crítico</b> (Faltan {} días)</span>', dias_faltantes)
        elif dias_faltantes <= 7:
            return format_html('<span style="color: orange;"><b>Atención</b> (Faltan {} días)</span>', dias_faltantes)
        else:
            return format_html('<span style="color: green;"><b>A tiempo</b> (Faltan {} días)</span>', dias_faltantes)
            
    apertura_semaforo.short_description = "Fecha de Apertura"

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('<path:object_id>/exportar-analisis/', self.admin_site.admin_view(self.exportar_analisis_view), name='exportar_analisis_licitacion'),
            path('<path:object_id>/exportar-reporte-laboratorios/', self.admin_site.admin_view(self.exportar_reporte_laboratorios_view), name='exportar_reporte_laboratorios_licitacion'),
            path('<path:object_id>/borrar-partidas/', self.admin_site.admin_view(self.borrar_partidas_view), name='borrar_partidas_licitacion'),
            path('<path:object_id>/notificar-socios/', self.admin_site.admin_view(self.notificar_socios_view), name='notificar_socios_licitacion'),
            path('<path:object_id>/notificar-resultados/', self.admin_site.admin_view(self.notificar_resultados_view), name='notificar_resultados_licitacion'),
        ]
        return custom_urls + urls

    def exportar_analisis_view(self, request, object_id):
        return exportar_analisis_licitacion(self, request, Licitacion.objects.filter(id=object_id))

    def exportar_reporte_laboratorios_view(self, request, object_id):
        return exportar_reporte_laboratorios(self, request, Licitacion.objects.filter(id=object_id))

    def save_model(self, request, obj, form, change):
        super().save_model(request, obj, form, change) 
        datos_excel = form.cleaned_data.get('pegar_excel')
        if datos_excel:
            import io, csv
            from django.db import transaction
            importes_agregados, claves_nuevas_creadas, errores_filas = 0, [], []
            texto_seguro = datos_excel.replace('\r\n', '\n').replace('\r', '\n')
            f = io.StringIO(texto_seguro)
            lector = csv.reader(f, dialect='excel-tab')
            for columnas in lector:
                if not columnas or not "".join(columnas).strip(): continue
                if len(columnas) >= 4:
                    try:
                        with transaction.atomic():
                            partida_val, clave_val, descripcion_val = columnas[0].strip(), columnas[1].strip(), columnas[2].strip().replace('\n', ' ') 
                            if len(columnas) >= 5 and columnas[4].strip():
                                min_str, max_str = columnas[3].strip().replace(',', '').replace(' ', ''), columnas[4].strip().replace(',', '').replace(' ', '')
                            else:
                                min_str, max_str = "0", columnas[3].strip().replace(',', '').replace(' ', '')
                            if not clave_val or not max_str: continue
                            num_partida = int(float(partida_val))
                            num_piezas_min, num_piezas_max = int(float(min_str)) if min_str else 0, int(float(max_str))
                            medicamento_db = CatalogoMedicamento.objects.filter(clave_sector=clave_val).first()
                            if not medicamento_db:
                                medicamento_db = CatalogoMedicamento(clave_sector=clave_val, descripcion=descripcion_val, denominacion_generica=descripcion_val, fabricante='')
                                medicamento_db.save()
                                claves_nuevas_creadas.append(clave_val)
                            PartidaRequerimiento.objects.create(licitacion=obj, numero_partida=num_partida, medicamento=medicamento_db, cantidad_minima=num_piezas_min, cantidad_maxima=num_piezas_max, costo=0, precio=0)
                            importes_agregados += 1
                    except Exception as e:
                        errores_filas.append(f"Partida {columnas[0]}: Error ({e})")
                        continue 
            if importes_agregados > 0: messages.success(request, f"¡Éxito! Se cargaron {importes_agregados} partidas al evento.")
            if claves_nuevas_creadas: messages.warning(request, f"🔔 Se agregaron {len(claves_nuevas_creadas)} CLAVES NUEVAS.")

    def borrar_partidas_view(self, request, object_id):
        licitacion = self.get_object(request, object_id)
        cantidad = licitacion.partidas.count()
        licitacion.partidas.all().delete()
        messages.success(request, f"Se eliminaron {cantidad} partidas.")
        return redirect('admin:licitaciones_licitacion_change', object_id)

    def notificar_socios_view(self, request, object_id):
        from django.core.mail import EmailMultiAlternatives
        from django.template.loader import render_to_string
        from django.utils.html import strip_tags
        from django.utils import timezone
        
        licitacion = self.get_object(request, object_id)
        partidas = licitacion.partidas.all()
        socios_dict = {}
        
        for p in partidas:
            clave_actual = p.medicamento.clave_sector
            medicamentos_relacionados = CatalogoMedicamento.objects.filter(clave_sector=clave_actual)
            for med in medicamentos_relacionados:
                socio = med.socio_contacto
                if socio:
                    if socio.id not in socios_dict: socios_dict[socio.id] = {'socio': socio, 'partidas': []}
                    if p not in socios_dict[socio.id]['partidas']: socios_dict[socio.id]['partidas'].append(p)

        if request.method == 'POST':
            socios_seleccionados = request.POST.getlist('socios')
            archivos_adjuntos = request.FILES.getlist('adjuntos') 
            empresa_id = request.POST.get('empresa_emisora')
            empresa_emisora = Empresa.objects.get(id=empresa_id)
            
            # --- LÓGICA DE RUTEO MÚLTIPLE (REPLY-TO) ---
            nombre_empresa_up = empresa_emisora.nombre.upper()
            
            # 1. Metemos el correo principal en una caja (lista)
            if "SAGO" in nombre_empresa_up: 
                lista_respuesta = ["sagomedical.licitaciones@gmail.com"]
            elif "GSM" in nombre_empresa_up: 
                lista_respuesta = ["gsm.licitaciones@gmail.com"]
            else: 
                lista_respuesta = ["licitaciones2@gpharma.com"]

            # 2. Revisamos si en el panel escribiste correos de empleados
            if empresa_emisora.correos_notificacion:
                empleados = [c.strip() for c in empresa_emisora.correos_notificacion.split(',') if c.strip()]
                lista_respuesta.extend(empleados)

            from django.core.mail import get_connection
            conexion_dinamica = get_connection()
            
            correos_enviados = 0
            
            for socio_id in socios_seleccionados:
                data = socios_dict.get(int(socio_id))
                if not data: continue
                socio = data['socio']
                asunto = f"Requerimiento de Cotización - Evento {licitacion.num_procedimiento}"
                
                partidas_html = []
                for p in data['partidas']:
                    partidas_html.append({
                        'partida': p.numero_partida,
                        'clave': p.medicamento.clave_sector,
                        'descripcion': p.medicamento.descripcion,
                        'cantidad': f"{p.cantidad_maxima:,}" 
                    })
                
                if "SAGO" in nombre_empresa_up: color_empresa = "#8B0000"
                elif "GAMS" in nombre_empresa_up: color_empresa = "#005b96"
                elif "GSM" in nombre_empresa_up: color_empresa = "#218838"
                else: color_empresa = "#333333"
                
                contexto_email = {
                    'socio_nombre': socio.nombre,
                    'evento_num': licitacion.num_procedimiento,
                    'dependencia': licitacion.dependencia,
                    'empresa_emisora': empresa_emisora.nombre,
                    'url_logo': empresa_emisora.url_logo if hasattr(empresa_emisora, 'url_logo') and empresa_emisora.url_logo else None,
                    'color_empresa': color_empresa,
                    'fecha_actual': timezone.now().strftime('%d/%m/%Y'),
                    'items': partidas_html,
                }
                
                html_content = render_to_string('admin/licitaciones/licitacion/emails/cotizacion_email.html', contexto_email)
                text_content = strip_tags(html_content) 

                destinatarios = [c.strip() for c in socio.correos.split(',') if c.strip()]
                
                try:
                    correo = EmailMultiAlternatives(
                        subject=asunto,
                        body=text_content,
                        from_email=f'"{empresa_emisora.nombre}" <{empresa_emisora.correo_remitente}>',
                        to=destinatarios,
                        connection=conexion_dinamica,
                        reply_to=lista_respuesta # AQUI VA LA LISTA MAESTRA
                    )
                    correo.attach_alternative(html_content, "text/html")
                    for archivo in archivos_adjuntos: 
                        correo.attach(archivo.name, archivo.read(), archivo.content_type)
                    correo.send(fail_silently=False)
                    correos_enviados += 1
                except Exception as e: 
                    messages.error(request, f"Error al enviar a {socio.nombre}: {e}")
            
            if correos_enviados > 0: messages.success(request, f"¡Se enviaron {correos_enviados} requerimientos!")
            return redirect('admin:licitaciones_licitacion_change', object_id)

        context = {'title': f'Notificar Socios: {licitacion.num_procedimiento}', 'licitacion': licitacion, 'socios_data': socios_dict.values(), 'opts': self.model._meta, 'empresas_grupo': Empresa.objects.all(), 'has_view_permission': self.has_view_permission(request, licitacion)}
        return render(request, 'admin/licitaciones/licitacion/notificar_socios.html', context)

    def notificar_resultados_view(self, request, object_id):
        from django.core.mail import EmailMultiAlternatives
        from django.template.loader import render_to_string
        from django.utils.html import strip_tags
        from django.utils import timezone
        
        licitacion = self.get_object(request, object_id)
        partidas = licitacion.partidas.select_related('medicamento__socio_contacto')
        socios_dict = {}
        
        for p in partidas:
            socio = p.medicamento.socio_contacto
            if socio:
                if socio.id not in socios_dict: socios_dict[socio.id] = {'socio': socio, 'asignadas': [], 'perdidas_precio': [], 'perdidas_tecnica': [], 'pendientes': []}
                if p.resultado == 'Asignada': socios_dict[socio.id]['asignadas'].append(p)
                elif p.resultado == 'Perdida por precio': socios_dict[socio.id]['perdidas_precio'].append(p)
                elif p.resultado == 'Perdida técnicamente': socios_dict[socio.id]['perdidas_tecnica'].append(p)
                else: socios_dict[socio.id]['pendientes'].append(p)

        if request.method == 'POST':
            socios_seleccionados = request.POST.getlist('socios')
            archivos_adjuntos = request.FILES.getlist('adjuntos') 
            empresa_id = request.POST.get('empresa_emisora')
            empresa_emisora = Empresa.objects.get(id=empresa_id)
            
            # --- LÓGICA DE RUTEO MÚLTIPLE (REPLY-TO) ---
            nombre_empresa_up = empresa_emisora.nombre.upper()
            
            if "SAGO" in nombre_empresa_up: 
                lista_respuesta = ["sagomedical.licitaciones@gmail.com"]
            elif "GSM" in nombre_empresa_up: 
                lista_respuesta = ["gsm.licitaciones@gmail.com"]
            else: 
                lista_respuesta = ["licitaciones2@gpharma.com"]

            if empresa_emisora.correos_notificacion:
                empleados = [c.strip() for c in empresa_emisora.correos_notificacion.split(',') if c.strip()]
                lista_respuesta.extend(empleados)

            from django.core.mail import get_connection
            conexion_dinamica = get_connection()
            
            correos_enviados = 0
            
            for socio_id in socios_seleccionados:
                data = socios_dict.get(int(socio_id))
                if not data: continue
                socio = data['socio']
                asunto = f"Resultados Oficiales (Evento {licitacion.num_procedimiento}) - {empresa_emisora.nombre}"
                
                if "SAGO" in nombre_empresa_up: color_empresa = "#8B0000"
                elif "GAMS" in nombre_empresa_up: color_empresa = "#005b96"
                elif "GSM" in nombre_empresa_up: color_empresa = "#218838"
                else: color_empresa = "#333333"

                def formato_item(p):
                    return {
                        'partida': p.numero_partida,
                        'clave': p.medicamento.clave_sector,
                        'descripcion': p.medicamento.descripcion,
                        'cantidad': f"{p.cantidad_maxima:,}",
                        'motivo': p.motivo_perdida if p.motivo_perdida else "No especificado en el sistema."
                    }

                contexto_email = {
                    'socio_nombre': socio.nombre,
                    'evento_num': licitacion.num_procedimiento,
                    'empresa_emisora': empresa_emisora.nombre,
                    'url_logo': empresa_emisora.url_logo if hasattr(empresa_emisora, 'url_logo') and empresa_emisora.url_logo else None,
                    'color_empresa': color_empresa,
                    'fecha_actual': timezone.now().strftime('%d/%m/%Y'),
                    'url_drive': licitacion.url_carpeta_drive if hasattr(licitacion, 'url_drive') and licitacion.url_carpeta_drive else None,
                    'asignadas': [formato_item(p) for p in data['asignadas']],
                    'perdidas_precio': [formato_item(p) for p in data['perdidas_precio']],
                    'perdidas_tecnica': [formato_item(p) for p in data['perdidas_tecnica']],
                }
                
                html_content = render_to_string('admin/licitaciones/licitacion/emails/resultados_email.html', contexto_email)
                text_content = strip_tags(html_content)

                destinatarios = [c.strip() for c in socio.correos.split(',') if c.strip()]
                
                try:
                    correo = EmailMultiAlternatives(
                        subject=asunto, 
                        body=text_content, 
                        from_email=f'"{empresa_emisora.nombre}" <{empresa_emisora.correo_remitente}>', 
                        to=destinatarios, 
                        connection=conexion_dinamica,
                        reply_to=lista_respuesta # AQUI VA LA LISTA MAESTRA
                    )
                    correo.attach_alternative(html_content, "text/html")
                    for archivo in archivos_adjuntos: correo.attach(archivo.name, archivo.read(), archivo.content_type)
                    correo.send(fail_silently=False)
                    correos_enviados += 1
                except Exception as e: 
                    messages.error(request, f"Error al enviar a {socio.nombre}: {e}")
            
            if correos_enviados > 0: messages.success(request, f"Resultados enviados.")
            return redirect('admin:licitaciones_licitacion_change', object_id)

        context = {'title': f'Notificar Resultados: {licitacion.num_procedimiento}', 'licitacion': licitacion, 'socios_data': socios_dict.values(), 'opts': self.model._meta, 'empresas_grupo': Empresa.objects.all(), 'has_view_permission': self.has_view_permission(request, licitacion)}
        return render(request, 'admin/licitaciones/licitacion/notificar_resultados.html', context)

@admin.register(Empresa)
class EmpresaAdmin(ImportExportModelAdmin):
    list_per_page = 30
    list_display = ('nombre', 'rfc', 'representante', 'telefono')
    search_fields = ('nombre', 'rfc')
    
    # --- CANDADO DE SEGURIDAD (Solo Superusuario) ---
    def has_module_permission(self, request):
        return request.user.is_superuser
        
    def has_view_permission(self, request, obj=None):
        return request.user.is_superuser


@admin.register(SocioComercial)
class SocioComercialAdmin(admin.ModelAdmin):
    list_per_page = 30
    list_display = ('nombre', 'correos', 'telefono')
    search_fields = ('nombre', 'correos')

@admin.register(PartidaRequerimiento)
class PartidaRequerimientoAdmin(admin.ModelAdmin):
    list_per_page = 50
    list_display = ('licitacion', 'numero_partida', 'clave_medicamento', 'cantidad_maxima', 'precio', 'importe_total', 'resultado')
    list_filter = ('resultado', 'licitacion__num_procedimiento')
    search_fields = ('numero_partida', 'medicamento__clave_sector', 'medicamento__denominacion_generica', 'licitacion__num_procedimiento')
    ordering = ('licitacion', 'numero_partida')
    list_editable = ('precio', 'resultado')

    # --- Ocultar del menú principal pero mantenerla funcional ---
    def has_module_permission(self, request):
        return False

    def clave_medicamento(self, obj):
        return f"{obj.medicamento.clave_sector} - {obj.medicamento.denominacion_generica[:30]}..."
    clave_medicamento.short_description = "Clave / Medicamento"
    clave_medicamento.admin_order_field = 'medicamento__clave_sector'

    def importe_total(self, obj):
        total = (obj.cantidad_maxima or 0) * (obj.precio or 0)
        return format_html('<b>${:,.2f}</b>', total)
    importe_total.short_description = "Importe Máximo"
    importe_total.admin_order_field = 'cantidad_maxima'


# admin.site.register(EstatusProcedimiento) # <-- Oculto del panel
admin.site.register(CatalogoMedicamento, CatalogoMedicamentoAdmin)
admin.site.register(RegistroUbicacion, RegistroUbicacionAdmin)

# ==========================================
# --- FASE 2: MÓDULO DE CONTRATOS Y LOGÍSTICA ---
# ==========================================

class FianzaContratoInline(admin.TabularInline):
    model = FianzaContrato
    extra = 1

class ClaveContratoInline(admin.TabularInline):
    model = ClaveContrato
    extra = 0
    autocomplete_fields = ['medicamento']
    fields = ('medicamento', 'cantidad_minima', 'cantidad_maxima', 'precio_neto', 'importe_maximo')
    readonly_fields = ('importe_maximo',)

    def importe_maximo(self, obj):
        if obj.cantidad_maxima and obj.precio_neto:
            total = obj.cantidad_maxima * obj.precio_neto
            total_formateado = f"{total:,.2f}"
            return format_html('<b>${}</b>', total_formateado)
        return "$0.00"
    importe_maximo.short_description = "Importe Máx."

# ==========================================
# RECURSO PARA IMPORTAR CONTRATOS MASIVAMENTE
# ==========================================

class EmpresaSeguraWidget(ForeignKeyWidget):
    def clean(self, value, row=None, **kwargs):
        if value: value = str(value).strip()
        return super().clean(value, row, **kwargs)

class LicitacionOpcionalWidget(ForeignKeyWidget):
    def clean(self, value, row=None, **kwargs):
        if not value or str(value).strip() in ['', 'None', 'N/A', 'NA']: return None
        return self.model.objects.filter(num_procedimiento=str(value).strip()).first()

class CargaMaestraContratoResource(resources.ModelResource):
    empresa = fields.Field(column_name='empresa', attribute='empresa', widget=EmpresaSeguraWidget(Empresa, field='nombre'))
    licitacion = fields.Field(column_name='licitacion', attribute='licitacion_origen', widget=LicitacionOpcionalWidget(Licitacion, field='num_procedimiento'))
    fecha_inicio = fields.Field(column_name='fecha_inicio', attribute='fecha_inicio', widget=DateWidget(format='%Y-%m-%d'))
    fecha_fin = fields.Field(column_name='fecha_fin', attribute='fecha_fin', widget=DateWidget(format='%Y-%m-%d'))

    class Meta:
        model = Contrato
        fields = ('numero_contrato', 'dependencia', 'empresa', 'fecha_inicio', 'fecha_fin', 'licitacion')
        import_id_fields = ('numero_contrato',)
        skip_unchanged = False

    def before_import_row(self, row, **kwargs):
        for key in list(row.keys()):
            if isinstance(row[key], str):
                row[key] = row[key].strip()
                
        for key in list(row.keys()):
            k_lower = str(key).strip().lower()
            val = row[key]
            
            if k_lower == 'empresa': row[key] = str(val).strip() if val else None
            elif k_lower == 'licitacion': row[key] = str(val).strip() if val else None
            elif k_lower in ['fecha_inicio', 'fecha_fin']:
                if str(val).strip() in ['', 'None', 'N/A', 'NA']: row[key] = None

        lic_str = None
        dep_str = 'S/D'
        for key, val in row.items():
            k_lower = str(key).strip().lower()
            if k_lower == 'licitacion':
                lic_str = str(val).strip() if val else None
            if k_lower == 'dependencia':
                dep_str = str(val).strip() if val else 'S/D'

        if lic_str:
            lic_obj = Licitacion.objects.filter(num_procedimiento=lic_str).first()
            if not lic_obj:
                Licitacion.objects.bulk_create([
                    Licitacion(num_procedimiento=lic_str, dependencia=dep_str)
                ])

        clave_sec = None
        for key, val in row.items():
            if str(key).strip().lower() == 'clave_sector':
                clave_sec = str(val).strip() if val else None
                break

        if clave_sec:
            medicamento = CatalogoMedicamento.objects.filter(clave_sector=clave_sec).first()
            if not medicamento:
                CatalogoMedicamento.objects.create(
                    clave_sector=clave_sec,
                    descripcion='Agregado por carga de Contrato', 
                    denominacion_generica='PENDIENTE ASIGNAR MARCA'
                )

    def after_import_row(self, row, row_result, **kwargs):
        if kwargs.get('dry_run'):
            return 
            
        num_contrato = None
        clave_sec = None
        cant_min, cant_max, precio = 0, 0, 0.0
        
        for key, val in row.items():
            k_lower = str(key).strip().lower()
            if k_lower == 'numero_contrato': num_contrato = str(val).strip() if val else None
            elif k_lower == 'clave_sector': clave_sec = str(val).strip() if val else None
            elif k_lower == 'cantidad_minima':
                try: cant_min = int(float(val or 0))
                except: cant_min = 0
            elif k_lower == 'cantidad_maxima':
                try: cant_max = int(float(val or 0))
                except: cant_max = 0
            elif k_lower == 'precio_neto':
                try: precio = float(val or 0.0)
                except: precio = 0.0

        if num_contrato and clave_sec:
            contrato_obj = Contrato.objects.filter(numero_contrato=num_contrato).first()
            med_obj = CatalogoMedicamento.objects.filter(clave_sector=clave_sec).first()
            
            if contrato_obj and med_obj:
                ClaveContrato.objects.update_or_create(
                    contrato=contrato_obj,
                    medicamento=med_obj,
                    defaults={
                        'cantidad_minima': cant_min,
                        'cantidad_maxima': cant_max,
                        'precio_neto': precio
                    }
                )

# ==========================================
# 2. PANTALLA PRINCIPAL: Contratos Maestros
# ==========================================
@admin.register(Contrato)
class ContratoAdmin(ImportExportModelAdmin): 
    resource_class = CargaMaestraContratoResource 
    
    list_per_page = 30
    list_display = ('numero_contrato',  'licitacion_origen', 'dependencia', 'empresa', 'fecha_inicio', 'fecha_fin', 'fianzas_badges', 'mostrar_monto', 'mostrar_avance', 'mostrar_abasto', 'boton_expediente', 'monto_penalizado')
    
    search_fields = ('numero_contrato', 'dependencia', 'fianzas__numero_fianza', 'licitacion_origen__num_procedimiento') 
    list_filter = ('empresa', 'fecha_fin', 'licitacion_origen')
    
    inlines = [FianzaContratoInline, ClaveContratoInline]
    autocomplete_fields = ['licitacion_origen']
    
    change_form_template = "admin/licitaciones/licitacion/boton_contrato.html"

    fieldsets = (
        ('🏢 Datos Generales del Contrato', {
            'fields': ('empresa', 'numero_contrato', 'dependencia')
        }),
        ('📅 Vigencia', {
            'fields': ('fecha_inicio', 'fecha_fin') 
        }),
        ('🤖 Vinculación', {
            'fields': ('licitacion_origen',),
            'description': 'Selecciona la Licitación de origen.'
        }),
        ('📂 Expediente Físico/Servidor', {
            'fields': ('ruta_carpeta_servidor',),
            'description': 'Pega la ruta de la carpeta compartida de tu red (Ej: V:\\REPORTES CONTRATOS...).'
        }),
    )

    def fianzas_badges(self, obj):
        fianzas = obj.fianzas.all()
        if not fianzas:
            return format_html('<span style="color: #ccc;">{}</span>', 'Sin Fianzas')
        
        badges = []
        for f in fianzas:
            badges.append(f'<span style="background-color: #17a2b8; color: white; padding: 3px 8px; border-radius: 10px; font-size: 10px; display: inline-block; margin-bottom: 3px;">{f.get_tipo_display()}: {f.numero_fianza}</span>')
        
        return format_html(''.join(badges))
    fianzas_badges.short_description = "Pólizas / Fianzas"

    def mostrar_monto(self, obj):
        monto = obj.monto_total_contrato
        monto_formateado = f"{monto:,.2f}"
        return format_html('<b style="color: #28a745;">${}</b>', monto_formateado)
    mostrar_monto.short_description = "Monto Total Máximo"

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('<path:object_id>/seleccionar-claves/', self.admin_site.admin_view(self.seleccionar_claves_view), name='seleccionar_claves_contrato'),
        ]
        return custom_urls + urls
    
    def boton_expediente(self, obj):
        if hasattr(obj, 'ruta_carpeta_servidor') and obj.ruta_carpeta_servidor:
            ruta_segura = obj.ruta_carpeta_servidor.replace('\\', '\\\\')
            return format_html(
                '<button type="button" onclick="navigator.clipboard.writeText(\'{}\'); alert(\'¡Ruta copiada! Abre tu explorador de archivos y pégala.\');" style="background-color: #6c757d; color: white; padding: 4px 10px; border: none; border-radius: 4px; font-size: 11px; font-weight: bold; cursor: pointer;">📋 Copiar Ruta</button>',
                ruta_segura
            )
        return mark_safe('<span style="color: #ccc;">-</span>')
    boton_expediente.short_description = "Expediente"

    def mostrar_avance(self, obj):
        avance = obj.porcentaje_avance
        color = 'green' if avance < 80 else ('orange' if avance <= 95 else 'red')
        return format_html('<b style="color: {};">{}%</b>', color, avance)
    mostrar_avance.short_description = 'Avance Financiero'

    def mostrar_abasto(self, obj):
        abasto = obj.porcentaje_abasto
        color = 'green' if abasto >= 95 else ('orange' if abasto >= 80 else 'red')
        return format_html('<b style="color: {};">{}%</b>', color, abasto)
    mostrar_abasto.short_description = 'Nivel de Abasto'

    def seleccionar_claves_view(self, request, object_id):
        contrato = self.get_object(request, object_id)
        if not contrato.licitacion_origen:
            messages.error(request, "Este contrato no tiene una Licitación de origen vinculada.")
            return redirect('admin:licitaciones_contrato_change', object_id)

        claves_ya_en_contrato = contrato.claves.values_list('medicamento_id', flat=True)
        partidas_ganadas = contrato.licitacion_origen.partidas.filter(resultado='Asignada').exclude(medicamento_id__in=claves_ya_en_contrato)

        if request.method == 'POST':
            partidas_seleccionadas = request.POST.getlist('partidas')
            claves_creadas = 0
            for partida_id in partidas_seleccionadas:
                partida = PartidaRequerimiento.objects.get(id=partida_id)
                cant_min = partida.cantidad_minima if partida.cantidad_minima else 0
                cant_max = partida.cantidad_maxima if partida.cantidad_maxima else 0
                precio = partida.precio if partida.precio else 0

                ClaveContrato.objects.create(
                    contrato=contrato,
                    medicamento=partida.medicamento,
                    precio_neto=precio,
                    cantidad_minima=cant_min,
                    cantidad_maxima=cant_max
                )
                claves_creadas += 1
            
            if claves_creadas > 0:
                messages.success(request, f"¡Éxito! Se agregaron {claves_creadas} claves al contrato {contrato.numero_contrato}.")
            return redirect('admin:licitaciones_contrato_change', object_id)

        context = {
            'title': f'Elegir Claves Ganadas: {contrato.numero_contrato}',
            'contrato': contrato,
            'partidas': partidas_ganadas,
            'opts': self.model._meta,
            'has_view_permission': self.has_view_permission(request, contrato)
        }
        return render(request, 'admin/licitaciones/licitacion/seleccionar_claves_contrato.html', context)

    def monto_penalizado(self, obj):
        ordenes = OrdenSuministro.objects.filter(clave_contrato__contrato=obj)
        try:
            total_multas = sum(float(orden.penalizacion_estimada) for orden in ordenes)
        except (ValueError, TypeError):
            total_multas = 0.0
            
        if total_multas > 0:
            monto_formateado = f"${total_multas:,.2f}"
            return format_html('<span style="color: #dc3545; font-weight: bold;">- {}</span>', monto_formateado)
            
        return mark_safe('<span style="color: #28a745; font-weight: bold;">$0.00</span>')
        
    monto_penalizado.short_description = "Penalizaciones (Est.)"

# ==========================================
# RECURSO PARA IMPORTAR ÓRDENES MASIVAMENTE
# ==========================================
class ClaveContratoWidget(ForeignKeyWidget):
    def clean(self, value, row=None, **kwargs):
        num_contrato = str(row.get('numero_contrato', '')).strip()
        clave_sec = str(row.get('clave_medicamento', '')).strip()
        
        if num_contrato and clave_sec:
            return ClaveContrato.objects.filter(
                contrato__numero_contrato__iexact=num_contrato,
                medicamento__clave_sector__iexact=clave_sec
            ).first()
            
        return None

class OrdenSuministroResource(resources.ModelResource):
    actions = ['marcar_como_canceladas']
    clave_contrato = fields.Field(
        column_name='clave_medicamento', 
        attribute='clave_contrato',
        widget=ClaveContratoWidget(ClaveContrato, field='id')
    )
    
    clave_medicamento_historico = fields.Field(
        column_name='clave_medicamento',
        attribute='clave_medicamento_historico'
    )
    numero_contrato_historico = fields.Field(
        column_name='numero_contrato',
        attribute='numero_contrato_historico'
    )

    fecha_recepcion = fields.Field(column_name='fecha_recepcion', attribute='fecha_recepcion', widget=DateWidget(format='%Y-%m-%d'))
    fecha_limite = fields.Field(column_name='fecha_limite', attribute='fecha_limite', widget=DateWidget(format='%Y-%m-%d'))
    fecha_entrega_real = fields.Field(column_name='fecha_entrega_real', attribute='fecha_entrega_real', widget=DateWidget(format='%Y-%m-%d'))

    class Meta:
        model = OrdenSuministro
        fields = (
            'razon_social', 'numero_orden_suministro', 'numero_procedimiento_extra', 
            'numero_contrato_historico', 'clave_medicamento_historico', 'descripcion_medicamento', 
            'cantidad_solicitada', 'cantidad_entregada', 'precio_unitario', 'clues_destino', 
            'entidad_destino', 'nombre_unidad', 'fecha_recepcion', 'fecha_limite', 
            'fecha_entrega_real', 'clave_contrato', 'estatus' 
        )
        import_id_fields = ('numero_orden_suministro', 'clave_medicamento_historico')

    def before_import_row(self, row, **kwargs):
        for key in list(row.keys()):
            if isinstance(row[key], str):
                row[key] = row[key].strip()
        
        row['numero_contrato_historico'] = str(row.get('numero_contrato', '')).strip()
        row['clave_medicamento_historico'] = str(row.get('clave_medicamento', '')).strip()

        try:
            solicitada = int(float(row.get('cantidad_solicitada', 0) or 0))
            entregada = int(float(row.get('cantidad_entregada', 0) or 0))
        except (ValueError, TypeError):
            solicitada = 0
            entregada = 0

        if entregada > 0:
            if entregada >= solicitada:
                row['estatus'] = 'ENTREGADA'
            else:
                row['estatus'] = 'PARCIAL'
        else:
            row['estatus'] = 'PENDIENTE'

# ==========================================
# 3. PANTALLA DE LOGÍSTICA: Órdenes de Suministro (OPMs)
# ==========================================
@admin.register(OrdenSuministro)
class OrdenSuministroAdmin(ImportExportModelAdmin):
    resource_class = OrdenSuministroResource
    
    list_per_page = 50
    list_display = (
        'razon_social', 'numero_orden_suministro', 'mostrar_contrato', 'mostrar_clave', 'dependencia', 'nombre_unidad',
        'cantidad_solicitada', 'piezas_entregadas', 'piezas_pendientes', 'alerta_inventario', 'fecha_limite', 
        'estatus_logistico', 'monto_penalizacion', 'estatus', 'btn_surtir'
    )
    search_fields = ('numero_orden_suministro', 'clave_medicamento_historico', 'numero_contrato_historico', 'nombre_unidad', 'clues_destino')
    list_filter = ('razon_social', 'dependencia', 'estatus', 'fecha_limite')
    list_editable = ('estatus',)

    change_list_template = "admin/ordenes_changelist.html"
    readonly_fields = ('alerta_inventario',)

    def piezas_entregadas(self, obj):
        from django.utils.html import format_html
        
        enviadas = sum((r.cantidad_entregada or 0) for r in obj.remisiones.exclude(estatus_viaje='RECHAZO'))
        
        enviadas_texto = f"{enviadas:,}"
        
        if enviadas >= obj.cantidad_solicitada and obj.cantidad_solicitada > 0:
            return format_html('<b style="color: #28a745;">{}</b>', enviadas_texto) 
        else:
            return format_html('<b style="color: #007bff;">{}</b>', enviadas_texto)
            
    piezas_entregadas.short_description = "Cant. Entregada"

    def piezas_pendientes(self, obj):
        from django.utils.html import format_html
        from django.utils.safestring import mark_safe 
        
        enviadas = sum((r.cantidad_entregada or 0) for r in obj.remisiones.exclude(estatus_viaje='RECHAZO'))
        pendientes = obj.cantidad_solicitada - enviadas
        
        pendientes_texto = f"{pendientes:,}"
        
        if pendientes <= 0:
            return mark_safe('<b style="color: #28a745;">0</b>') 
        else:
            return format_html('<b style="color: #dc3545;">{}</b>', pendientes_texto)
            
    piezas_pendientes.short_description = "Cant. Pendiente"
    
    def mostrar_contrato(self, obj):
        if obj.clave_contrato:
            return obj.clave_contrato.contrato.numero_contrato
        return f"{obj.numero_contrato_historico} (No cruzó DB)"
    mostrar_contrato.short_description = "No. Contrato"

    def mostrar_clave(self, obj):
        if obj.clave_contrato:
            return obj.clave_contrato.medicamento.clave_sector
        return f"{obj.clave_medicamento_historico} (No cruzó DB)"
    mostrar_clave.short_description = "Clave"

    def estatus_logistico(self, obj):
        from django.utils.html import format_html
        
        color = "#6c757d"
        texto = obj.get_estatus_display()
        
        if obj.estatus == 'ENTREGADA': color = "#28a745"
        elif obj.estatus == 'PENDIENTE': color = "#ffc107"
        elif obj.estatus == 'DEVUELTA': color = "#dc3545"
        elif obj.estatus == 'PARCIAL': color = "#17a2b8"
        elif obj.estatus == 'NO_ATENDIDA': color = "#343a40"

        if obj.estatus in ['CANCELADA', 'CANCELADA_EVIDENCIA']:
            piezas_entregadas = sum((r.cantidad_entregada or 0) for r in obj.remisiones.exclude(estatus_viaje='RECHAZO'))
            
            if piezas_entregadas > 0:
                return format_html(
                    '<div style="text-align: center; line-height: 1.2;">'
                    '<span style="background-color: #000000; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 10px; display: block; margin-bottom: 2px;">🚨 CANCELADA (TIENE EVIDENCIA)</span>'
                    '<span style="font-size: 10px; color: #dc3545; font-weight: bold;">Se entregaron {:,} pzas</span>'
                    '</div>', piezas_entregadas
                )
            else:
                return format_html('<span style="background-color: #6f42c1; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; font-size: 11px;">🚫 CANCELADA (Sin entregas)</span>')

        if obj.estatus in ['PENDIENTE', 'PARCIAL'] and obj.dias_atraso > 0:
            return format_html('<span style="background-color: #dc3545; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; font-size: 11px;">⚠️ ATRASADA ({} DÍAS)</span>', obj.dias_atraso)

        return format_html('<span style="background-color: {}; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; font-size: 11px;">{}</span>', color, texto)
    estatus_logistico.short_description = "Semáforo"

    def monto_penalizacion(self, obj):
        try:
            penalizacion = float(obj.penalizacion_estimada)
        except (ValueError, TypeError):
            penalizacion = 0.0

        if penalizacion > 0:
            monto_formateado = f"${penalizacion:,.2f}"
            return format_html('<span style="color: #dc3545; font-weight: bold;">- {}</span>', monto_formateado)
            
        return mark_safe('<span style="color: #ccc;">$0.00</span>')
    monto_penalizacion.short_description = "Penalización Acumulada"

    def alerta_inventario(self, obj):
        from django.utils.html import format_html
        from django.utils.safestring import mark_safe
        from django.db.models import Sum
        from .models import Inventario

        if obj.estatus == 'ENTREGADA':
            return mark_safe('<span style="color: #28a745; font-weight: bold;">✔ STOCK ENTREGADO</span>')

        clave_buscar = None
        if obj.clave_contrato and obj.clave_contrato.medicamento:
            clave_buscar = obj.clave_contrato.medicamento.clave_sector
        elif obj.clave_medicamento_historico:
            clave_buscar = obj.clave_medicamento_historico
            
        if not clave_buscar: 
            return mark_safe('<span style="color: #ccc;">Sin Clave</span>')
            
        stock = Inventario.objects.filter(
            medicamento__clave_sector=clave_buscar, 
            cantidad_disponible__gt=0
        ).aggregate(total=Sum('cantidad_disponible'))['total'] or 0

        if stock == 0:
            return mark_safe('<b style="color: #dc3545;">❌ 0 en Stock (Agotado)</b>') 
        elif stock >= obj.cantidad_solicitada:
            return format_html('<b style="color: #28a745;">✅ {} en Stock</b>', f"{stock:,}")
        else:
            faltan = obj.cantidad_solicitada - stock
            return format_html('<b style="color: #ffc107;">⚠️ {} en Stock<br><small>(Faltan {})</small></b>', f"{stock:,}", f"{faltan:,}")
            
    alerta_inventario.short_description = "Validación de Almacén"
    
    def btn_surtir(self, obj):
        from django.utils.html import format_html
        from django.utils.safestring import mark_safe 
        
        if obj.estatus == 'ENTREGADA':
            return mark_safe('<span style="color: green; font-weight:bold;">✔ Completada</span>')
            
        return format_html(
            '<a class="button" href="{}/surtir/" style="background-color: #007bff; color:white; padding: 5px 10px; border-radius: 4px; text-decoration: none; font-weight:bold;">🚚 Surtir</a>',
            obj.id
        )
    btn_surtir.short_description = "Logística"
    
    def get_urls(self):
        urls = super().get_urls()
        from django.urls import path
        custom_urls = [
            path('<path:object_id>/surtir/', self.admin_site.admin_view(self.surtir_orden_view), name='surtir_orden_logistica'),
        ]
        return custom_urls + urls

    def surtir_orden_view(self, request, object_id):
        from django.shortcuts import render, redirect
        from django.contrib import messages
        from .models import Inventario, RemisionEntrega
        
        orden = self.get_object(request, object_id)
        
        clave_buscar = orden.clave_contrato.medicamento.clave_sector if (orden.clave_contrato and orden.clave_contrato.medicamento) else orden.clave_medicamento_historico
        lotes_disponibles = Inventario.objects.filter(medicamento__clave_sector=clave_buscar, cantidad_disponible__gt=0).order_by('fecha_caducidad')
        ya_enviado = sum(r.cantidad_entregada for r in orden.remisiones.all())
        falta_enviar = orden.cantidad_solicitada - ya_enviado

        if request.method == 'POST':
            lote_id = request.POST.get('lote_id')
            cantidad_a_enviar = int(request.POST.get('cantidad', 0))
            folio_remision = request.POST.get('folio_remision')
            
            if cantidad_a_enviar > 0 and lote_id and folio_remision:
                lote_seleccionado = Inventario.objects.get(id=lote_id)
                
                if cantidad_a_enviar > lote_seleccionado.cantidad_disponible:
                    messages.error(request, "No puedes enviar más piezas de las que hay en ese lote.")
                else:
                    lote_seleccionado.cantidad_disponible -= cantidad_a_enviar
                    lote_seleccionado.save()
                    
                    RemisionEntrega.objects.create(
                        orden=orden,
                        folio_remision_factura=folio_remision,
                        cantidad_entregada=cantidad_a_enviar,
                        lote=lote_seleccionado.lote,
                        caducidad=lote_seleccionado.fecha_caducidad
                    )
                    
                    orden.estatus = 'EN_RUTA'
                    orden.save()
                    
                    messages.success(request, f"¡Camión en ruta! Se despacharon {cantidad_a_enviar} piezas con la remisión {folio_remision}.")
                    return redirect('admin:licitaciones_ordensuministro_changelist')

        context = {
            'title': f'Surtir Orden: {orden.numero_orden_suministro}',
            'orden': orden,
            'lotes': lotes_disponibles,
            'falta_enviar': falta_enviar,
            'opts': self.model._meta,
        }
        return render(request, 'admin/licitaciones/ordensuministro/surtir_orden.html', context)

    @admin.action(description='🚫 Marcar seleccionadas como CANCELADAS (Audita evidencia automático)')
    def marcar_como_canceladas(self, request, queryset):
        ordenes_procesadas = 0
        con_evidencia = 0
        
        for orden in queryset:
            piezas_entregadas = sum((r.cantidad_entregada or 0) for r in orden.remisiones.exclude(estatus_viaje='RECHAZO'))
            
            if piezas_entregadas > 0:
                orden.estatus = 'CANCELADA_EVIDENCIA'
                con_evidencia += 1
            else:
                orden.estatus = 'CANCELADA'
            
            orden.save()
            ordenes_procesadas += 1
            
        from django.contrib import messages
        if con_evidencia > 0:
            messages.warning(request, f"¡Atención! Se cancelaron {ordenes_procesadas} órdenes, pero {con_evidencia} se marcaron CON EVIDENCIA para protección legal porque ya se habían entregado piezas.")
        else:
            messages.success(request, f"Se cancelaron {ordenes_procesadas} órdenes correctamente.")

    def get_actions(self, request):
        actions = super().get_actions(request)
        if not request.user.is_superuser:
            if 'marcar_como_canceladas' in actions:
                del actions['marcar_como_canceladas']
                
        return actions


# ==========================================
# REGISTRO DEL NUEVO MÓDULO DE INVENTARIO
# ==========================================

# 1. Creamos el "Mapa" (Resource) para que el botón de Excel sepa de dónde jalar los datos extra
class InventarioResource(resources.ModelResource):
    clave_sector = fields.Field(attribute='medicamento__clave_sector', column_name='CLAVE SECTOR')
    descripcion = fields.Field(attribute='medicamento__denominacion_generica', column_name='DESCRIPCIÓN')
    fabricante = fields.Field(attribute='medicamento__fabricante', column_name='FABRICANTE')
    socio_comercial = fields.Field(attribute='medicamento__socio_contacto__nombre', column_name='SOCIO COMERCIAL')

    class Meta:
        model = Inventario
        # 2. Le dictamos el orden exacto de las columnas que queremos en el Excel
        fields = ('almacen__nombre', 'clave_sector', 'descripcion', 'socio_comercial', 'fabricante', 'tipo_producto', 'lote', 'fecha_caducidad', 'cantidad_disponible')
        export_order = fields

# 3. Le inyectamos el "Mapa" al panel de administración
@admin.register(Inventario)
class InventarioAdmin(ImportExportModelAdmin):
    resource_class = InventarioResource # <--- ESTA ES LA LÍNEA MÁGICA QUE FALTABA
    list_per_page = 50
    list_display = ('almacen', 'medicamento', 'tipo_producto', 'lote', 'fecha_caducidad', 'cantidad_disponible', 'fecha_ingreso')
    search_fields = ('medicamento__clave_sector', 'medicamento__denominacion_generica', 'lote')
    list_filter = ('tipo_producto', 'almacen', 'fecha_caducidad',)
    autocomplete_fields = ['medicamento']
    actions = [exportar_inventario_personalizado]

# 4. PANTALLA FÍSICA: Remisiones de Almacén (Viajes)
@admin.register(RemisionEntrega)
class RemisionEntregaAdmin(admin.ModelAdmin):
    list_per_page = 50
    list_display = ('folio_remision_factura', 'orden_vinculada', 'cantidad_entregada', 'estatus_viaje')
    search_fields = ('folio_remision_factura', 'lote', 'orden__numero_orden_suministro')
    list_filter = ('estatus_viaje',)

    fieldsets = (
        ('📦 Datos del Viaje', {
            'fields': ('orden', 'folio_remision_factura', 'cantidad_entregada', 'lote', 'caducidad')
        }),
        ('🚚 Logística y Comprobación', {
            'fields': ('estatus_viaje', 'archivo_evidencia')
        }),
        ('🔴 Incidencias: Rechazo', {
            'fields': ('motivo_rechazo', 'evidencia_rechazo'),
            'description': 'Llena esto si el Instituto rechazó las cajas por caducidad, maltrato, etc.'
        }),
    )

    def orden_vinculada(self, obj):
        return obj.orden.numero_orden_suministro
    orden_vinculada.short_description = "OPM Vinculada"

# ==========================================
# RECURSO PARA IMPORTAR CLAVES A LOS CONTRATOS (CARGA HISTÓRICA)
# ==========================================
class ClaveContratoResource(resources.ModelResource):
    contrato = fields.Field(
        column_name='numero_contrato',
        attribute='contrato',
        widget=ForeignKeyWidget(Contrato, field='numero_contrato')
    )
    medicamento = fields.Field(
        column_name='clave_sector',
        attribute='medicamento',
        widget=ForeignKeyWidget(CatalogoMedicamento, field='clave_sector')
    )

    class Meta:
        model = ClaveContrato
        fields = ('contrato', 'medicamento', 'cantidad_minima', 'cantidad_maxima', 'precio_neto')
        import_id_fields = ('contrato', 'medicamento')

# ==========================================
# 🛒 MÓDULO DE COMPRAS Y RECEPCIÓN
# ==========================================
class PartidaCompraInline(admin.TabularInline):
    model = PartidaCompra
    extra = 1
    autocomplete_fields = ['medicamento']
    fields = ('medicamento', 'cantidad', 'precio_unitario', 'importe_visual', 'cantidad_recibida')
    readonly_fields = ('importe_visual',)

    def importe_visual(self, obj):
        if obj.cantidad and obj.precio_unitario:
            total = obj.cantidad * obj.precio_unitario
            total_formateado = f"{total:,.2f}"
            return format_html('<b>${}</b>', total_formateado)
        return "$0.00"
    importe_visual.short_description = "Importe"

    def formfield_for_foreignkey(self, db_field, request, **kwargs):
        formfield = super().formfield_for_foreignkey(db_field, request, **kwargs)
        if db_field.name == "medicamento":
            formfield.widget.attrs.update({'style': 'width: 450px; min-width: 450px;'})
        return formfield

class DocumentoOrdenCompraInline(admin.TabularInline):
    model = DocumentoOrdenCompra
    extra = 1
    fields = ('descripcion', 'archivo')

@admin.register(OrdenCompra)
class OrdenCompraAdmin(admin.ModelAdmin):
    list_display = ('folio', 'proveedor', 'fecha_entrega_esperada', 'mostrar_total', 'penalizacion_calculada', 'auditoria_ciega', 'estatus_badge', 'enviar_oc_link')
    list_filter = ('estatus', 'empresa_compradora', 'proveedor')
    
    search_fields = ('folio', 'proveedor__nombre', 'destino') 
    
    inlines = [PartidaCompraInline, DocumentoOrdenCompraInline]
    autocomplete_fields = ['proveedor'] 
    list_per_page = 30
    
    actions = ['marcar_recibida_y_crear_inventario', 'notificar_proveedor_masivo']

    def penalizacion_calculada(self, obj):
        return f"${obj.penalizacion_calculada:,.2f}"
    penalizacion_calculada.short_description = "📉 Penalización por Atraso"

    def enviar_oc_link(self, obj):
        return format_html(
            '<a class="button" href="{}" style="background-color: #5e35b1; color:white; padding: 5px 10px; border-radius: 4px; text-decoration: none;">✉️ Enviar OC</a>',
            f"{obj.id}/notificar-proveedor/"
        )
    enviar_oc_link.short_description = "Notificar"

    def get_urls(self):
        urls = super().get_urls()
        from django.urls import path
        custom_urls = [
            path('<path:object_id>/notificar-proveedor/', self.admin_site.admin_view(self.notificar_proveedor_view), name='notificar_proveedor_oc'),
            path('<path:object_id>/vista-previa-pdf/', self.admin_site.admin_view(self.vista_previa_pdf_view), name='vista_previa_pdf_oc'),
        ]
        return custom_urls + urls

    def vista_previa_pdf_view(self, request, object_id):
        from django.http import HttpResponse
        from django.template.loader import render_to_string
        from django.shortcuts import redirect
        from io import BytesIO
        
        try:
            from xhtml2pdf import pisa
        except ImportError:
            self.message_user(request, "⚠️ Falta instalar xhtml2pdf. Ejecuta: pip install xhtml2pdf", level='error')
            return redirect('admin:licitaciones_ordencompra_changelist')

        oc = self.get_object(request, object_id)
        empresa = oc.empresa_compradora

        nombre_empresa = empresa.nombre.upper()
        if "SAGO" in nombre_empresa: color_empresa = "#8B0000"
        elif "GAMS" in nombre_empresa: color_empresa = "#005b96"
        elif "GSM" in nombre_empresa: color_empresa = "#218838"
        else: color_empresa = "#333333"

        contexto = {
            'oc': oc,
            'proveedor': oc.proveedor,
            'empresa': empresa,
            'partidas': oc.partidas_compra.all(),
            'color_empresa': color_empresa,
            'url_logo': empresa.url_logo if hasattr(empresa, 'url_logo') else None,
        }

        html_pdf = render_to_string('admin/licitaciones/ordencompra/pdf/orden_compra_pdf.html', contexto)
        result_file = BytesIO()
        
        pisa_status = pisa.CreatePDF(html_pdf, dest=result_file)
        
        if not pisa_status.err:
            response = HttpResponse(result_file.getvalue(), content_type='application/pdf')
            response['Content-Disposition'] = f'inline; filename="Borrador_OC_{oc.folio}.pdf"'
            return response
        else:
            self.message_user(request, "⚠️ Hubo un error al procesar el PDF de vista previa.", level='error')
            return redirect('admin:licitaciones_ordencompra_changelist')

    def notificar_proveedor_view(self, request, object_id):
        from django.core.mail import EmailMultiAlternatives, get_connection
        from django.template.loader import render_to_string
        from django.utils.html import strip_tags
        from django.shortcuts import render, redirect
        from django.contrib import messages 
        import mimetypes
        from io import BytesIO
        
        try:
            from xhtml2pdf import pisa
        except ImportError:
            messages.error(request, "❌ ERROR: No tienes instalado 'xhtml2pdf'. Corre 'pip install xhtml2pdf' en la terminal.")
            return redirect('admin:licitaciones_ordencompra_changelist')

        oc = self.get_object(request, object_id)
        proveedor = oc.proveedor
        
        if request.method == 'POST':
            empresa = oc.empresa_compradora
            
            if not empresa.correo_remitente:
                messages.error(request, f"❌ ERROR DE CONFIGURACIÓN: La empresa '{empresa.nombre}' necesita un correo remitente registrado.")
                return redirect(request.path)

            try:
                # --- LÓGICA DE RUTEO MÚLTIPLE (REPLY-TO) ---
                nombre_empresa_up = empresa.nombre.upper()
                
                if "SAGO" in nombre_empresa_up: 
                    lista_respuesta = ["sagomedical.licitaciones@gmail.com"]
                elif "GSM" in nombre_empresa_up: 
                    lista_respuesta = ["gsm.licitaciones@gmail.com"]
                else: 
                    lista_respuesta = ["licitaciones2@gpharma.com"]

                if empresa.correos_notificacion:
                    empleados = [c.strip() for c in empresa.correos_notificacion.split(',') if c.strip()]
                    lista_respuesta.extend(empleados)

                conexion_dinamica = get_connection()
                
                asunto = f"ORDEN DE COMPRA OFICIAL: {oc.folio} - {empresa.nombre}"
                partidas = oc.partidas_compra.all()
                
                color_empresa = "#5e35b1" 
                if "SAGO" in nombre_empresa_up: color_empresa = "#8B0000"
                elif "GAMS" in nombre_empresa_up: color_empresa = "#005b96"
                elif "GSM" in nombre_empresa_up: color_empresa = "#218838"

                contexto = {
                    'oc': oc, 'proveedor': proveedor, 'empresa': empresa,
                    'partidas': partidas, 'total': f"{oc.total_compra:,.2f}",
                    'color_empresa': color_empresa,
                }
                
                html_content = render_to_string('admin/licitaciones/ordencompra/emails/orden_compra_email.html', contexto)
                text_content = strip_tags(html_content)
                
                destinatarios = [c.strip() for c in proveedor.correos.split(',') if c.strip()]
                
                msg = EmailMultiAlternatives(
                    subject=asunto, 
                    body=text_content, 
                    from_email=f'"{empresa.nombre}" <{empresa.correo_remitente}>', 
                    to=destinatarios, 
                    connection=conexion_dinamica,
                    reply_to=lista_respuesta # AQUI VA LA LISTA MAESTRA
                )
                msg.attach_alternative(html_content, "text/html")
                
                html_pdf = render_to_string('admin/licitaciones/ordencompra/pdf/orden_compra_pdf.html', contexto)
                result_file = BytesIO()
                pisa_status = pisa.CreatePDF(html_pdf, dest=result_file)
                
                if not pisa_status.err:
                    msg.attach(f"Orden_de_Compra_{oc.folio}.pdf", result_file.getvalue(), 'application/pdf')

                for documento in oc.documentos.all():
                    if documento.archivo:
                        content_type, _ = mimetypes.guess_type(documento.archivo.name)
                        msg.attach(documento.archivo.name.split('/')[-1], documento.archivo.read(), content_type or 'application/octet-stream')

                msg.send()
                
                if oc.estatus == 'BORRADOR':
                    oc.estatus = 'AUTORIZADA'
                    oc.save()
                
                messages.success(request, f"🚀 ¡ENVIADO! La OC {oc.folio} ya va en camino a {proveedor.nombre}.")
                return redirect('admin:licitaciones_ordencompra_changelist')
            
            except Exception as e:
                messages.error(request, f"❌ ERROR DE RED/RESEND: {str(e)}")

        context = {
            'title': f'Enviar Orden de Compra: {oc.folio}',
            'oc': oc, 'proveedor': proveedor, 'opts': self.model._meta,
        }
        return render(request, 'admin/licitaciones/ordencompra/confirmar_envio_oc.html', context)

    @admin.action(description='✉️ Enviar OC seleccionadas a Proveedores')
    def notificar_proveedor_masivo(self, request, queryset):
        self.message_user(request, "Las órdenes seleccionadas han sido procesadas.")

    def estatus_badge(self, obj):
        colores = {
            'BORRADOR': '#6c757d',   
            'AUTORIZADA': '#007bff', 
            'TRANSITO': '#fd7e14',   
            'RECIBIDA': '#28a745',   
            'CANCELADA': '#dc3545'   
        }
        color = colores.get(obj.estatus, '#000')
        return format_html(
            '<span style="background-color: {}; color: white; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 11px;">{}</span>',
            color, obj.get_estatus_display()
        )
    estatus_badge.short_description = "Estatus"

    def mostrar_total(self, obj):
        total_formateado = f"{obj.total_compra:,.2f}"
        return format_html('<b style="color: #5e35b1;">${}</b>', total_formateado)
    mostrar_total.short_description = "Total OC"

    @admin.action(description='📦 Marcar como RECIBIDA e Ingresar al Inventario')
    def marcar_recibida_y_crear_inventario(self, request, queryset):
        ordenes_procesadas = 0
        for orden in queryset:
            if orden.estatus == 'RECIBIDA':
                continue 
                
            orden.estatus = 'RECIBIDA'
            orden.save()
            
            import datetime
            for partida in orden.partidas_compra.all():
                partida.cantidad_recibida = partida.cantidad
                partida.save()
                
                Inventario.objects.create(
                    medicamento=partida.medicamento,
                    cantidad_disponible=partida.cantidad,
                    lote=f"LOT-{orden.folio}", 
                    fecha_ingreso=datetime.date.today(),
                    fecha_caducidad=datetime.date.today() + datetime.timedelta(days=730) 
                )
            ordenes_procesadas += 1
            
        if ordenes_procesadas > 0:
            messages.success(request, f"¡Éxito! Se recibieron {ordenes_procesadas} órdenes de compra y el STOCK SE ACTUALIZÓ en el Almacén.")
        else:
            messages.warning(request, "Las órdenes seleccionadas ya estaban recibidas previamente.")

    def auditoria_ciega(self, obj):
        from django.utils.html import format_html
        
        reclamado_compras = sum(p.cantidad_recibida for p in obj.partidas_compra.all() if p.cantidad_recibida)
        recibido_almacen = sum(e.cantidad_recibida for e in obj.entradaalmacen_set.all() if e.cantidad_recibida)
        
        if reclamado_compras == 0 and recibido_almacen == 0:
            from django.utils.safestring import mark_safe
            return mark_safe('<span style="color: #95a5a6; font-weight: bold;">➖ Pendiente</span>')
            
        if reclamado_compras == recibido_almacen:
            return format_html('<span style="color: #28a745; font-weight: bold;">✅ CUADRA PERFECTO ({})</span>', recibido_almacen)
        else:
            diferencia = recibido_almacen - reclamado_compras
            texto_dif = f"Faltan {-diferencia}" if diferencia < 0 else f"Sobran {diferencia}"
            return format_html(
                '<div style="line-height: 1.2;">'
                '<span style="color: #dc3545; font-weight: bold;">❌ DISCREPANCIA</span><br>'
                '<span style="font-size: 10px; color: #6c757d;">Compras: <b>{}</b> | Almacén: <b>{}</b></span><br>'
                '<span style="font-size: 10px; color: #dc3545; font-weight: bold;">{}</span>'
                '</div>', 
                reclamado_compras, recibido_almacen, texto_dif
            )
    auditoria_ciega.short_description = "Auditoría de Recepción"

# ==========================================
# 📦 FORMULARIO INTELIGENTE DE ALMACÉN (AJAX)
# ==========================================
class EntradaAlmacenForm(forms.ModelForm):
    class Meta:
        model = EntradaAlmacen
        fields = '__all__'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        if self.instance and self.instance.pk and self.instance.orden_id:
            self.fields['medicamento'].queryset = CatalogoMedicamento.objects.filter(
                id__in=PartidaCompra.objects.filter(orden=self.instance.orden).values('medicamento')
            )
        elif self.data.get('orden'):
            self.fields['medicamento'].queryset = CatalogoMedicamento.objects.filter(
                id__in=PartidaCompra.objects.filter(orden_id=self.data.get('orden')).values('medicamento')
            )
        else:
            self.fields['medicamento'].queryset = CatalogoMedicamento.objects.none()

        if 'orden' in self.fields:
            js = """
            <script type="text/javascript">
                document.addEventListener('DOMContentLoaded', function() {
                    var $ = django.jQuery || window.jQuery;
                    
                    $('#id_orden').on('change', function() {
                        var ordenId = $(this).val();
                        var $medSelect = $('#id_medicamento');
                        
                        if (!ordenId) {
                            $medSelect.empty();
                            $medSelect.append('<option value="">--- Primero selecciona una Orden ---</option>');
                            $medSelect.trigger('change');
                            return;
                        }
                        
                        $medSelect.html('<option value="">🔄 Buscando claves en la OC...</option>');
                        
                        var url = '/admin/licitaciones/entradaalmacen/ajax/load-medicamentos/?orden_id=' + ordenId;
                        
                        $.ajax({
                            url: url,
                            success: function(data) {
                                $medSelect.empty();
                                if(data.length === 0){
                                    $medSelect.append('<option value="">⚠️ Esta Orden no tiene partidas guardadas</option>');
                                } else {
                                    $medSelect.append('<option value="">--- Selecciona la clave recibida ---</option>');
                                    $.each(data, function(index, item) {
                                        $medSelect.append('<option value="' + item.id + '">' + item.text + '</option>');
                                    });
                                }
                                $medSelect.trigger('change');
                            },
                            error: function() {
                                alert("Error al conectar. Revisa tu internet.");
                            }
                        });
                    });
                });
            </script>
            """
            self.fields['orden'].help_text = mark_safe((self.fields['orden'].help_text or '') + js)

# ==========================================
# 📦 MÓDULO DE ADUANA / RECEPCIÓN ALMACÉN
# ==========================================
@admin.register(EntradaAlmacen)
class EntradaAlmacenAdmin(admin.ModelAdmin):
    form = EntradaAlmacenForm 
    
    list_per_page = 30
    list_display = ('orden', 'almacen_destino', 'medicamento', 'cantidad_recibida', 'lote', 'ver_acuse', 'ver_factura', 'documentacion_ok', 'fecha_ingreso')
    search_fields = ('orden__folio', 'medicamento__clave_sector', 'lote', 'ubicacion')
    list_filter = ('almacen_destino', 'documentacion_completa', 'fecha_ingreso')
    
    autocomplete_fields = ['orden']

    fieldsets = (
        ('🔗 Vinculación', {
            'fields': ('almacen_destino', 'orden', 'medicamento')
        }),
        ('📦 Datos Físicos', {
            'fields': ('cantidad_recibida', 'lote', 'fecha_caducidad')
        }),
        ('📋 Calidad y Logística', {
            'fields': ('documentacion_completa', 'ubicacion', 'observaciones_calidad', 'acuse_recibo', 'factura_proveedor')
        }),
    )

    def get_urls(self):
        urls = super().get_urls()
        from django.urls import path
        custom_urls = [
            path('ajax/load-medicamentos/', self.admin_site.admin_view(self.load_medicamentos), name='ajax_load_medicamentos_almacen'),
        ]
        return custom_urls + urls

    def load_medicamentos(self, request):
        from django.http import JsonResponse
        from .models import PartidaCompra
        orden_id = request.GET.get('orden_id')
        
        if not orden_id:
            return JsonResponse([])
            
        partidas = PartidaCompra.objects.filter(orden_id=orden_id).select_related('medicamento')
        data = []
        for p in partidas:
            faltantes = (p.cantidad or 0) - (p.cantidad_recibida or 0)
            texto = f"Clave: {p.medicamento.clave_sector} | {p.medicamento.denominacion_generica[:40]}... (Esperando: {faltantes} pz)"
            data.append({'id': p.medicamento.id, 'text': texto})
            
        return JsonResponse(data, safe=False)

    def ver_acuse(self, obj):
        from django.utils.html import format_html
        if obj.acuse_recibo:
            return format_html('<a href="{}" target="_blank" style="background-color: #17a2b8; color: white; padding: 4px 8px; border-radius: 4px; text-decoration: none; font-size: 11px; font-weight: bold;">📄 Acuse</a>', obj.acuse_recibo.url)
        return "-"
    ver_acuse.short_description = "Acuse"

    def ver_factura(self, obj):
        from django.utils.html import format_html
        if obj.factura_proveedor:
            return format_html('<a href="{}" target="_blank" style="background-color: #6c757d; color: white; padding: 4px 8px; border-radius: 4px; text-decoration: none; font-size: 11px; font-weight: bold;">🧾 Factura</a>', obj.factura_proveedor.url)
        return "-"
    ver_factura.short_description = "Factura"

    def documentacion_ok(self, obj):
        from django.utils.safestring import mark_safe 
        if obj.documentacion_completa:
            return mark_safe('<span style="color: #28a745; font-weight: bold;">✔ Completa</span>')
        return mark_safe('<span style="color: #dc3545; font-weight: bold;">❌ Faltante</span>')
    documentacion_ok.short_description = "Documentación"

# ==========================================
# PANTALLA DE TRASPASOS ENTRE ALMACENES
# ==========================================
@admin.register(TraspasoIntercompany)
class TraspasoIntercompanyAdmin(admin.ModelAdmin):
    list_display = ('folio_factura', 'almacen_origen', 'almacen_destino', 'medicamento', 'lote', 'cantidad', 'mostrar_importe', 'estatus')
    search_fields = ('folio_factura', 'medicamento__clave_sector', 'lote')
    list_filter = ('estatus', 'almacen_origen', 'almacen_destino', 'fecha_operacion')
    autocomplete_fields = ['medicamento']
    
    def get_readonly_fields(self, request, obj=None):
        if obj and obj.procesado:
            return ('almacen_origen', 'almacen_destino', 'medicamento', 'lote', 'cantidad', 'precio_unitario', 'folio_factura', 'estatus')
        return ()

    def mostrar_importe(self, obj):
        from django.utils.html import format_html
        total = obj.cantidad * obj.precio_unitario
        total_formateado = f"{total:,.2f}"
        return format_html('<b>${}</b>', total_formateado)
        
    mostrar_importe.short_description = "Valor Fiscal Total"

# @admin.register(ConfiguracionEmail) # <-- Oculto porque la configuración se hace en el modelo Empresa
# class ConfiguracionEmailAdmin(admin.ModelAdmin):
#     list_display = ('empresa', 'email_host_user')

from django.http import HttpResponseRedirect
from django.urls import reverse
from .models import EscanerKardex # Asegúrate de importar el modelo que acabamos de crear

@admin.register(EscanerKardex)
class EscanerKardexAdmin(admin.ModelAdmin):
    # Esto intercepta el clic y te manda a la ruta secreta
    def changelist_view(self, request, extra_context=None):
        return HttpResponseRedirect(reverse('buscar_kardex'))