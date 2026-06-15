import os
from decimal import Decimal
from django.db import models
from django.db.models import Sum, F 
from django.conf import settings
from django.contrib.auth.models import User
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from django.utils import timezone
from datetime import date, timedelta
from django.core.exceptions import ValidationError

TIPO_PRODUCTO_CHOICES = [
    ('GENERICO_SEC', 'Genérico (Sectorizado)'),
    ('COMERCIAL_SEC', 'Comercial (Sectorizado)'),
    ('GENERICO', 'Genérico'),
    ('COMERCIAL', 'Comercial'),
]

# ==========================================
# 🏛️ CATÁLOGO MAESTRO DE DEPENDENCIAS
# ==========================================
DEPENDENCIAS_MAESTRAS = [
    ('HOSPITALES_GENERALES', (
        ('HGM', 'Hospital General de México "Dr. Eduardo Liceaga"'),
        ('HGEA', 'Hospital General Dr. Manuel Gea González'),
        ('HJUAREZ', 'Hospital Juárez de México'),
        ('HJUAREZ_CENTRO', 'Hospital Juárez del Centro'),
        ('HMUJER', 'Hospital de la Mujer'),
        ('HHOMEOPATICO', 'Hospital Homeopático'),
    )),
    ('INSTITUTOS_NACIONALES', (
        ('INCAN', 'Instituto Nacional de Cancerología'),
        ('INCARDIO', 'Instituto Nacional de Cardiología Ignacio Chávez'),
        ('INCMNSZ', 'Instituto Nacional de Ciencias Médicas y Nutrición Salvador Zubirán'),
        ('INER', 'Instituto Nacional de Enfermedades Respiratorias'),
        ('INGER', 'Instituto Nacional de Geriatría'),
        ('INMEGEN', 'Instituto Nacional de Medicina Genómica'),
        ('INNN', 'Instituto Nacional de Neurología y Neurocirugía'),
        ('INP', 'Instituto Nacional de Pediatría'),
        ('INPER', 'Instituto Nacional de Perinatología'),
        ('INP_RAMON', 'Instituto Nacional de Psiquiatría Ramón de la Fuente'),
        ('INR', 'Instituto Nacional de Rehabilitación'),
        ('INSP', 'Instituto Nacional de Salud Pública'),
        ('HIMFG', 'Hospital Infantil de México Federico Gómez'),
    )),
    ('INSTITUCIONES_FEDERALES', (
        ('IMSS_BIENESTAR', 'IMSS-Bienestar'),
        ('IMSS_ORDINARIO', 'Instituto Mexicano del Seguro Social (IMSS)'),
        ('ISSSTE', 'ISSSTE'),
        ('PEMEX', 'PEMEX'),
        ('SEMAR', 'Secretaría de Marina'),
        ('SEDENA', 'Secretaría de la Defensa Nacional'),
        ('CONASAMA', 'Comisión Nacional de Salud Mental y Adicciones'),
        ('CENAPRECE', 'Centro Nacional de Programas Preventivos y Control de Enfermedades'),
        ('CENSIDA', 'Centro Nacional para la Prevención y el Control del VIH y el SIDA'),
        ('BIRMEX', 'Laboratorios de Biológicos y Reactivos de México'),
    )),
    ('HRAE_REGIONALES', (
        ('HRAE_BAJIO', 'HRAE del Bajío'),
        ('HRAE_IXTAPALUCA', 'HRAE Ixtapaluca'),
        ('HRAE_VICTORIA', 'HRAE de Cd. Victoria'),
        ('HRAE_OAXACA', 'HRAE de Oaxaca'),
        ('HRAE_YUCATAN', 'HRAE de la Península de Yucatán'),
        ('HEP_CHIAPAS', 'Hospital de Especialidades Pediátricas Chiapas'),
        ('HRAE_CIUDAD_SALUD', 'Hospital Regional de Alta Especialidad Ciudad Salud')
    )),
    ('OTRA', 'Otra Dependencia'),
]


class EstatusProcedimiento(models.Model):
    ESTATUS_CHOICES = [
        ('ADJUDICADO', 'Adjudicado'),
        ('PERDIDO', 'Perdido'),
        ('EN_PROCESO', 'En proceso'),
    ]
    estado = models.CharField(max_length=20, choices=ESTATUS_CHOICES, default='EN_PROCESO')

    def __str__(self):
        return self.estado

class Licitacion(models.Model):
    num_procedimiento = models.CharField(max_length=100, unique=True, verbose_name="Número de procedimiento")
    empresa = models.ForeignKey('Empresa', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Empresa Participante")
    fecha_publicacion = models.DateTimeField(verbose_name="Fecha y hora de publicación", null=True, blank=True)
    fecha_apertura = models.DateTimeField(verbose_name="Fecha y hora de apertura", null=True, blank=True)
    fecha_junta = models.DateTimeField(verbose_name="Fecha y hora de junta de aclaraciones", null=True, blank=True)
    fecha_fallo = models.DateTimeField(verbose_name="Fecha y hora del acto del Fallo", null=True, blank=True)
    dependencia = models.CharField(max_length=100, choices=DEPENDENCIAS_MAESTRAS, verbose_name="Dependencia")
    estatus = models.ForeignKey(EstatusProcedimiento, on_delete=models.SET_NULL, null=True)
    url_carpeta_drive = models.URLField(max_length=500, blank=True, null=True, verbose_name="URL Carpeta Drive")

    class Meta:
        verbose_name = "Licitación"
        verbose_name_plural = "Licitaciones"

    def __str__(self):
        return f"{self.num_procedimiento} - {self.get_dependencia_display()}"

    def save(self, *args, **kwargs):
        es_nuevo = self.pk is None 
        if es_nuevo and not self.url_carpeta_drive:
            try:
                ruta_token = os.path.join(settings.BASE_DIR, 'token.json')
                if os.path.exists(ruta_token):
                    creds = Credentials.from_authorized_user_file(ruta_token, ['https://www.googleapis.com/auth/drive'])
                    service = build('drive', 'v3', credentials=creds)
                    resultados = service.files().list(q="name='ERP_Licitaciones_Maestra' and mimeType='application/vnd.google-apps.folder'", spaces='drive', fields='files(id, name)').execute()
                    carpetas_maestras = resultados.get('files', [])
                    if carpetas_maestras:
                        id_maestra = carpetas_maestras[0]['id']
                        nombre_nueva_carpeta = f"{self.num_procedimiento} - {self.dependencia}"
                        metadatos = {'name': nombre_nueva_carpeta, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [id_maestra]}
                        carpeta_creada = service.files().create(body=metadatos, fields='id, webViewLink').execute()
                        self.url_carpeta_drive = carpeta_creada.get('webViewLink')
            except Exception as e:
                print(f"Error Drive: {e}")
        super().save(*args, **kwargs)

class CatalogoMedicamento(models.Model):
    clave_sector = models.CharField(max_length=50)
    descripcion = models.TextField()
    denominacion_generica = models.CharField(max_length=150)
    denominacion_distintiva = models.CharField(max_length=150)
    fabricante = models.CharField(max_length=150)
    socio_contacto = models.ForeignKey('SocioComercial', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Contacto para Notificaciones")
    rfc_fabricante = models.CharField(max_length=20)
    pais_fabricacion = models.CharField(max_length=50)
    num_registro_sanitario = models.CharField(max_length=50)
    num_prorroga = models.CharField(max_length=50, default="NO APLICA")
    codigo_barras = models.CharField(max_length=50)
    fecha_expedicion = models.DateField(null=True, blank=True)
    fecha_vigencia = models.DateField(null=True, blank=True)

    class Meta:
        verbose_name = "Clave"
        verbose_name_plural = "Claves"

    def __str__(self):
        marca = self.denominacion_distintiva if self.denominacion_distintiva else "Genérico"
        lab = self.fabricante if self.fabricante else "Sin Laboratorio"
        return f"Clave: {self.clave_sector} | Marca: {marca} | Lab: {lab}"

class PartidaRequerimiento(models.Model):
    licitacion = models.ForeignKey(Licitacion, on_delete=models.CASCADE, related_name='partidas')
    numero_partida = models.IntegerField(verbose_name="No. Partida", null=True, blank=True) 
    medicamento = models.ForeignKey(CatalogoMedicamento, on_delete=models.PROTECT, verbose_name="Medicamento")
    cantidad_minima = models.PositiveIntegerField(null=True, blank=True, default=0)
    cantidad_maxima = models.PositiveIntegerField(null=True, blank=True, default=0)
    costo = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, default=0)
    precio = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, default=0)

    RESULTADOS_OPCIONES = [
        ('Pendiente', 'Pendiente'),
        ('Asignada', 'Asignada'),
        ('Perdida por precio', 'Perdida por precio'),
        ('Perdida técnicamente', 'Perdida técnicamente'),
    ]
    resultado = models.CharField(max_length=30, choices=RESULTADOS_OPCIONES, default='Pendiente', verbose_name="Resultado")
    motivo_perdida = models.CharField(max_length=255, blank=True, null=True, verbose_name="Motivo Técnico", help_text="Escribe por qué se perdió (solo si fue pérdida técnica).")
    
    @property
    def valor_minimo_costo(self): return (self.costo * self.cantidad_minima) if self.costo and self.cantidad_minima else 0
    @property
    def valor_maximo_costo(self): return (self.costo * self.cantidad_maxima) if self.costo and self.cantidad_maxima else 0
    @property
    def valor_minimo_precio(self): return (self.precio * self.cantidad_minima) if self.precio and self.cantidad_minima else 0
    @property
    def valor_maximo_precio(self): return (self.precio * self.cantidad_maxima) if self.precio and self.cantidad_maxima else 0

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs) # Primero guardamos la partida normal
        
        # Si ganamos la partida y tiene precio, la mandamos al historial
        if self.resultado == 'Asignada' and self.precio and self.precio > 0:
            proveedor_nombre = 'Sin Laboratorio'
            if self.medicamento.socio_contacto:
                proveedor_nombre = self.medicamento.socio_contacto.nombre
            elif self.medicamento.fabricante:
                proveedor_nombre = self.medicamento.fabricante

            HistorialPrecio.objects.update_or_create(
                procedimiento=self.licitacion.num_procedimiento,
                clave=self.medicamento.clave_sector,
                defaults={
                    'medicamento': self.medicamento,
                    'descripcion': self.medicamento.descripcion,
                    'tipo_procedimiento': 'Licitación Pública',
                    'proveedor': proveedor_nombre,
                    'institucion': self.licitacion.get_dependencia_display() if hasattr(self.licitacion, 'get_dependencia_display') else self.licitacion.dependencia,
                    'institucion_desagregada': self.licitacion.dependencia,
                    'precio': self.precio,
                    'cantidad': self.cantidad_maxima or 0,
                    'valor': (self.cantidad_maxima or 0) * self.precio
                }
            )

class RegistroUbicacion(models.Model):
    usuario = models.ForeignKey(User, on_delete=models.CASCADE, verbose_name="Empleado")
    latitud = models.CharField(max_length=50)
    longitud = models.CharField(max_length=50)
    fecha_hora = models.DateTimeField(auto_now_add=True, verbose_name="Fecha y hora del Check-in")

    class Meta:
        verbose_name = "Registro de Ubicación"
        verbose_name_plural = "Registros de Ubicaciones"

    def __str__(self):
        return f"{self.usuario.username} - {self.fecha_hora.strftime('%d/%m/%Y %H:%M')}"

class Empresa(models.Model):
    nombre = models.CharField(max_length=200, verbose_name="Nombre de la Empresa")
    rfc = models.CharField(max_length=20, blank=True, null=True, verbose_name="RFC")
    representante = models.CharField(max_length=150, blank=True, null=True, verbose_name="Representante Legal")
    telefono = models.CharField(max_length=20, blank=True, null=True, verbose_name="Teléfono")
    PROVEEDOR_CHOICES = [('smtp.gmail.com', 'Gmail'), ('smtp.office365.com', 'Outlook / Office 365')]
    servidor_correo = models.CharField(max_length=100, default='smtp.resend.com', verbose_name="Servidor SMTP", help_text="Ej: smtp.resend.com", blank=True, null=True)
    correo_remitente = models.EmailField(blank=True, null=True, verbose_name="Correo Emisor (Gmail/Outlook)")
    password_aplicacion = models.CharField(max_length=100, blank=True, null=True, verbose_name="Contraseña de Aplicación (16 letras)")
    url_logo = models.URLField(max_length=500, null=True, blank=True, verbose_name="URL del Logo (Público)", help_text="Pega aquí el link directo a la imagen del logo (debe terminar en .png o .jpg). Ej: https://tudominio.com/logo_sago.png")
    correos_notificacion = models.TextField(blank=True, null=True, help_text="Escribe aquí los correos de los empleados para ESTA empresa, separados por coma.")
    
    class Meta:
        verbose_name = "Empresa"
        verbose_name_plural = "Empresas"

    def __str__(self):
        return self.nombre

class SocioComercial(models.Model):
    nombre = models.CharField(max_length=200, verbose_name="Nombre del Laboratorio / Socio")
    correos = models.CharField(max_length=250, verbose_name="Correos Electrónicos", help_text="Puedes poner varios correos separados por coma (Ej: ventas@pisa.com, dir@pisa.com)")
    telefono = models.CharField(max_length=50, blank=True, null=True, verbose_name="Teléfono de Contacto")

    class Meta:
        verbose_name = "Socio Comercial"
        verbose_name_plural = "Directorio de Socios"

    def __str__(self):
        return self.nombre


# ==========================================
# FASE 2: MÓDULO DE CONTRATOS Y ABASTO LOGÍSTICO
# ==========================================

class Contrato(models.Model):
    empresa = models.ForeignKey('Empresa', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Empresa GPHARMA")
    numero_contrato = models.CharField(max_length=100, unique=True, verbose_name="Número de Contrato")
    dependencia = models.CharField(max_length=100, choices=DEPENDENCIAS_MAESTRAS, verbose_name="Dependencia / Cliente")
    fecha_inicio = models.DateField(verbose_name="Inicio de Vigencia", null=True, blank=True)
    fecha_fin = models.DateField(verbose_name="Fin de Vigencia", null=True, blank=True)
    licitacion_origen = models.ForeignKey('Licitacion', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Viene de la Licitación")
    
    tiene_convenio_modificatorio = models.BooleanField(default=False, verbose_name="¿Tiene Convenio Modificatorio?")
    porcentaje_ampliacion = models.DecimalField(max_digits=5, decimal_places=2, default=0.00, help_text="Ej: 20.00 si se amplió el 20%")

    ruta_carpeta_servidor = models.CharField(max_length=500, blank=True, null=True, verbose_name="Ruta en Servidor Local (Ej: V:\\REPORTES...)")
    porcentaje_penalizacion_diaria = models.DecimalField("Penalización Diaria (%)", max_digits=5, decimal_places=2, default=2.5, help_text="Ej. 2.5% por día de atraso")
    tope_penalizacion = models.DecimalField("Tope de Penalización (%)", max_digits=5, decimal_places=2, default=10.0, help_text="Límite de la penalización (Ej. 10%)")
    monto_fianza = models.DecimalField("Monto de Fianza (10%)", max_digits=15, decimal_places=2, blank=True, null=True, help_text="Se calcula en automático (10% del monto total).")

    def save(self, *args, **kwargs):
        if hasattr(self, 'monto_total') and self.monto_total: 
            self.monto_fianza = self.monto_total * Decimal('0.10')
        super().save(*args, **kwargs)

    class Meta:
        verbose_name = "Contrato"
        verbose_name_plural = "Contratos"
        
    def __str__(self):
        return f"{self.numero_contrato} - {self.get_dependencia_display()}"

    @property
    def monto_total_contrato(self):
        res = self.claves.aggregate(total=Sum(F('cantidad_maxima') * F('precio_neto')))
        return res['total'] or Decimal('0.00')

    @property
    def porcentaje_avance(self):
        maximo = self.monto_total_contrato
        if maximo == 0: return 0.0
        
        res_nuevo = PartidaOrden.objects.filter(clave_contrato__contrato=self).aggregate(
            consumido=Sum(F('cantidad_solicitada') * F('clave_contrato__precio_neto'))
        )
        consumido_nuevo = res_nuevo['consumido'] or Decimal('0.00')
        
        res_hist = self.claves.aggregate(
            consumido_hist=Sum(F('piezas_historicas_solicitadas') * F('precio_neto'))
        )
        consumido_hist = res_hist['consumido_hist'] or Decimal('0.00')
        
        return round((float(consumido_nuevo + consumido_hist) / float(maximo)) * 100, 2)

    @property
    def porcentaje_abasto(self):
        res_sol_nuevo = PartidaOrden.objects.filter(clave_contrato__contrato=self).aggregate(tot=Sum('cantidad_solicitada'))
        res_sol_hist = self.claves.aggregate(tot=Sum('piezas_historicas_solicitadas'))
        sol_total = (res_sol_nuevo['tot'] or 0) + (res_sol_hist['tot'] or 0)
        
        if sol_total == 0: return 0.0
        
        res_ent_nuevo = RemisionEntrega.objects.filter(orden__partidas__clave_contrato__contrato=self).distinct().aggregate(tot=Sum('cantidad_entregada'))
        res_ent_hist = self.claves.aggregate(tot=Sum('piezas_historicas_entregadas'))
        ent_total = (res_ent_nuevo['tot'] or 0) + (res_ent_hist['tot'] or 0)
        
        return round((ent_total / float(sol_total)) * 100, 2)

    @property
    def piezas_totales_maximas(self):
        res = self.claves.aggregate(total=Sum('cantidad_maxima'))
        return res['total'] or 0

    @property
    def piezas_totales_minimas(self):
        res = self.claves.aggregate(total=Sum('cantidad_minima'))
        return res['total'] or 0

class FianzaContrato(models.Model):
    TIPO_FIANZA = [
        ('ANTICIPO', 'Fianza de Anticipo'),
        ('LICITACION', 'Fianza de Licitación'),
        ('CUMPLIMIENTO', 'Fianza de Cumplimiento'),
        ('VICIOS', 'Fianza de Vicios Ocultos / Buena Calidad'),
        ('OTRA', 'Otra...'),
    ]
    contrato = models.ForeignKey(Contrato, on_delete=models.CASCADE, related_name='fianzas', verbose_name="Contrato")
    tipo = models.CharField(max_length=20, choices=TIPO_FIANZA, verbose_name="Tipo de Fianza")
    numero_fianza = models.CharField(max_length=100, verbose_name="No. de Póliza / Fianza")
    afianzadora = models.CharField(max_length=150, blank=True, null=True, verbose_name="Aseguradora / Afianzadora", help_text="Ej. Sofimex, Aserta, etc.")
    monto = models.DecimalField(max_digits=12, decimal_places=2, blank=True, null=True, verbose_name="Monto Asegurado")

    class Meta:
        verbose_name = "Fianza / Garantía"
        verbose_name_plural = "Fianzas del Contrato"

    def __str__(self):
        return f"{self.get_tipo_display()} - {self.numero_fianza}"

class ClaveContrato(models.Model):
    contrato = models.ForeignKey(Contrato, on_delete=models.CASCADE, related_name='claves')
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave / Medicamento")
    cantidad_minima = models.IntegerField(default=0, verbose_name="Cant. Mínima")
    cantidad_maxima = models.IntegerField(default=0, verbose_name="Cant. Máxima")
    precio_neto = models.DecimalField(max_digits=10, decimal_places=2, default=0.00, verbose_name="Precio Neto")
    piezas_historicas_solicitadas = models.IntegerField(default=0, verbose_name="Histórico Solicitado")
    piezas_historicas_entregadas = models.IntegerField(default=0, verbose_name="Histórico Entregado")
    
    class Meta:
        verbose_name = "Clave Asignada"
        verbose_name_plural = "Claves del Contrato"

    def __str__(self):
        return f"{self.medicamento.clave_sector} - {self.contrato.numero_contrato}"

class OrdenSuministro(models.Model):
    TIPO_CHOICES = [
        ('SUMINISTRO', 'Orden de Suministro (Gobierno)'),
        ('PEDIDO', 'Pedido Privado (Cliente)'),
    ]
    tipo_documento = models.CharField(max_length=20, choices=TIPO_CHOICES, default='SUMINISTRO', verbose_name='Tipo de Documento')
    razon_social = models.CharField(max_length=200, blank=True, null=True, verbose_name="Razón Social (Empresa)")
    dependencia = models.CharField(max_length=100, choices=DEPENDENCIAS_MAESTRAS, blank=True, null=True, verbose_name="Dependencia / Institución")
    numero_orden_suministro = models.CharField(max_length=150, verbose_name="Folio (No. Orden / Pedido)")
    
    clues_destino = models.CharField(max_length=100, blank=True, null=True, verbose_name="CLUES Destino")
    entidad_destino = models.CharField(max_length=250, blank=True, null=True, verbose_name="Entidad / Lugar de Entrega")
    nombre_unidad = models.CharField(max_length=250, blank=True, null=True, verbose_name="Nombre de Unidad / Dependencia")
    
    fecha_recepcion = models.DateField(verbose_name="Fecha en que se recibió", default=timezone.now)
    fecha_limite = models.DateField(verbose_name="Fecha Límite de Entrega")
    fecha_entrega_real = models.DateField(null=True, blank=True, verbose_name='Fecha real de entrega')

    ESTATUS_ORDEN = [
        ('PENDIENTE', 'Pendiente'),
        ('PARCIAL', 'Entrega Parcial'),
        ('ENTREGADA', 'Entregada'),
        ('DEVUELTA', 'Devuelta / Rechazada'),
        ('NO_ATENDIDA', 'No Atendida'),
        ('CANCELADA', 'Cancelada (Sin entrega)'),
        ('CANCELADA_EVIDENCIA', 'Cancelada (CON EVIDENCIA DE ENTREGA)'),
    ]
    estatus = models.CharField(max_length=20, choices=ESTATUS_ORDEN, default='PENDIENTE', verbose_name="Estatus Logístico")
    motivo_incidencia = models.TextField(blank=True, null=True, verbose_name="Motivo (Rechazo / No Atención)")

    class Meta:
        verbose_name = "Orden de Suministro"
        verbose_name_plural = "Órdenes de Suministro"
        ordering = ['fecha_limite'] 

    def __str__(self):
        tipo = dict(self.TIPO_CHOICES).get(self.tipo_documento, "Documento")
        cliente = self.razon_social or self.get_dependencia_display() or "Sin Cliente"
        return f"{tipo}: {self.numero_orden_suministro} | {cliente}"

    @property
    def dias_atraso(self):
        from django.utils import timezone
        if not self.fecha_limite: return 0
        if self.estatus in ['ENTREGADA', 'CANCELADA_EVIDENCIA']:
            fecha_cierre = self.fecha_entrega_real or timezone.now().date()
            dias = (fecha_cierre - self.fecha_limite).days
        else:
            hoy = timezone.now().date()
            dias = (hoy - self.fecha_limite).days
        return dias if dias > 0 else 0

    @property
    def valor_total(self):
        return sum(partida.importe for partida in self.partidas.all())

    @property
    def penalizacion_estimada(self):
        dias = self.dias_atraso
        if dias <= 0: return 0.0
        importe_total = float(self.valor_total)
        porcentaje_multa = min(dias * 0.02, 0.10)
        return importe_total * porcentaje_multa


class PartidaOrden(models.Model):
    orden = models.ForeignKey(OrdenSuministro, on_delete=models.CASCADE, related_name='partidas')
    clave_contrato = models.ForeignKey(ClaveContrato, on_delete=models.CASCADE, verbose_name="Clave del Medicamento", null=True, blank=True)
    medicamento = models.ForeignKey(CatalogoMedicamento, on_delete=models.CASCADE, verbose_name="Catálogo Libre (Privados)", null=True, blank=True)
    cantidad_solicitada = models.IntegerField(verbose_name="Cant. Solicitada")
    cantidad_entregada = models.IntegerField(default=0, verbose_name="Cant. Entregada")
    precio_unitario = models.DecimalField(max_digits=12, decimal_places=2, default=0.00, verbose_name="Precio Unitario")
    
    clave_historica = models.CharField(max_length=50, blank=True, null=True, verbose_name="Clave Histórica")
    
    @property
    def importe(self):
        return float(self.cantidad_solicitada or 0) * float(self.precio_unitario or 0)

    class Meta:
        verbose_name = "Medicamento / Clave"
        verbose_name_plural = "Lista de Medicamentos"

    def __str__(self):
        return str(self.clave_contrato) if self.clave_contrato else "Partida Manual"
    
class RemisionEntrega(models.Model):
    orden = models.ForeignKey(OrdenSuministro, on_delete=models.CASCADE, related_name='remisiones', verbose_name="Orden de Suministro")
    folio_remision_factura = models.CharField(max_length=100, verbose_name="Folio Remisión / Factura")
    cantidad_entregada = models.IntegerField(verbose_name="Piezas Enviadas en este viaje")
    lote = models.CharField(max_length=50, verbose_name="Lote del Medicamento")
    caducidad = models.DateField(verbose_name="Fecha de Caducidad")
    
    fecha_salida = models.DateTimeField(auto_now_add=True, verbose_name="Fecha y Hora de Salida")
    archivo_evidencia = models.FileField(upload_to='logistica/evidencias/', blank=True, null=True, verbose_name="PDF Remisión Sellada/Firmada", help_text="Sube aquí el acuse firmado para dar la orden por Entregada.")
    
    ESTATUS_VIAJE = [
        ('EN_RUTA', '🚚 En Ruta hacia el Instituto'), 
        ('ENTREGADA', '✅ Entregada y Comprobada'),
        ('PARCIAL', '🔵 Entrega Parcial (Faltan piezas)'),
        ('RECHAZO', '🔴 Rechazada por el Instituto')
    ]
    estatus_viaje = models.CharField(max_length=20, choices=ESTATUS_VIAJE, default='EN_RUTA', verbose_name="Estatus del Viaje")

    evidencia_rechazo = models.FileField(upload_to='logistica/rechazos/', blank=True, null=True, verbose_name="Evidencia del Rechazo (Acta/Foto)")
    motivo_rechazo = models.TextField("Motivo del Rechazo", blank=True, null=True)

    class Meta:
        verbose_name = "Remisión Física"
        verbose_name_plural = "Remisiones de Almacén"

    def __str__(self):
        return f"Remisión {self.folio_remision_factura} - Orden: {self.orden.numero_orden_suministro}"
        
    def save(self, *args, **kwargs):
        es_nuevo = self.pk is None
        
        # 👇 1. Capturamos el estatus anterior para saber si apenas lo acaban de rechazar
        estatus_anterior = None
        if not es_nuevo:
            estatus_anterior = RemisionEntrega.objects.get(pk=self.pk).estatus_viaje

        if self.archivo_evidencia and self.estatus_viaje == 'EN_RUTA':
            self.estatus_viaje = 'ENTREGADA'
            
        super().save(*args, **kwargs)

        orden = self.orden
        piezas_comprobadas = sum(r.cantidad_entregada for r in orden.remisiones.all() if r.estatus_viaje == 'ENTREGADA')
        total_solicitado_orden = sum(p.cantidad_solicitada for p in orden.partidas.all())
        
        if self.estatus_viaje == 'ENTREGADA':
            if piezas_comprobadas >= total_solicitado_orden:
                orden.estatus = 'ENTREGADA'
            else:
                orden.estatus = 'PARCIAL'
        elif self.estatus_viaje == 'PARCIAL':
            orden.estatus = 'PARCIAL'
        elif self.estatus_viaje == 'RECHAZO':
            if piezas_comprobadas == 0:
                orden.estatus = 'DEVUELTA'
            elif piezas_comprobadas < total_solicitado_orden:
                orden.estatus = 'PARCIAL' # Si rechazaron una parte pero otra sí pasó
            
        orden.save()

        from .models import MovimientoKardex, Inventario

        # 🔥 LÓGICA DE SALIDA (Cuando el camión sale de almacén hacia el hospital)
        if es_nuevo:
            stock = Inventario.objects.filter(lote__iexact=self.lote).first()

            if stock:
                stock.cantidad_disponible -= self.cantidad_entregada
                stock.save()

                destino_final = orden.nombre_unidad or orden.entidad_destino or "Instituto"
                MovimientoKardex.objects.create(
                    almacen=stock.almacen,
                    medicamento=stock.medicamento,
                    lote=self.lote,
                    tipo='SALIDA_OPM',
                    cantidad=self.cantidad_entregada,
                    saldo_restante=stock.cantidad_disponible,
                    folio_documento=self.folio_remision_factura,
                    observaciones=f"Enviado con Orden {orden.numero_orden_suministro} a {destino_final}"
                )
                
        # 🔥 LÓGICA DE DEVOLUCIÓN (Cuando el hospital rechaza y la mercancía regresa al inventario)
        if not es_nuevo and self.estatus_viaje == 'RECHAZO' and estatus_anterior != 'RECHAZO':
            stock = Inventario.objects.filter(lote__iexact=self.lote).first()
            
            if stock:
                # Regresamos las piezas a la bodega
                stock.cantidad_disponible += self.cantidad_entregada
                stock.save()
                
                # Dejamos la huella legal en el Kardex
                MovimientoKardex.objects.create(
                    almacen=stock.almacen,
                    medicamento=stock.medicamento,
                    lote=self.lote,
                    tipo='ENTRADA_DEVOLUCION',
                    cantidad=self.cantidad_entregada,
                    saldo_restante=stock.cantidad_disponible,
                    folio_documento=self.folio_remision_factura,
                    observaciones=f"Devolución por Rechazo (Orden {orden.numero_orden_suministro}). Motivo: {self.motivo_rechazo}"
                )

class Almacen(models.Model):
    nombre = models.CharField("Nombre del Almacén", max_length=100, unique=True)
    descripcion = models.TextField("Descripción / Ubicación Físca", blank=True, null=True)

    class Meta:
        verbose_name = "Almacén"
        verbose_name_plural = "Almacenes Físicos"

    def __str__(self):
        return self.nombre

class Inventario(models.Model):
    almacen = models.ForeignKey('Almacen', on_delete=models.CASCADE, verbose_name="Almacén", null=True, blank=True)
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave / Medicamento")
    lote = models.CharField(max_length=50, verbose_name="Lote del Producto")
    fecha_caducidad = models.DateField(verbose_name="Fecha de Caducidad")
    tipo_producto = models.CharField(max_length=20, choices=TIPO_PRODUCTO_CHOICES, default='COMERCIAL', verbose_name="Tipo de Producto")
    codigo_barras = models.CharField(max_length=150, blank=True, null=True, verbose_name="Código de Barras", help_text="Escanea aquí con la pistola de códigos")
    cantidad_disponible = models.IntegerField(default=0, verbose_name="Piezas Físicas Disponibles")
    fecha_ingreso = models.DateTimeField(auto_now_add=True, verbose_name="Fecha de Ingreso al Sistema")

    class Meta:
        verbose_name = "Inventario Físico"
        verbose_name_plural = "Inventarios"
        ordering = ['fecha_caducidad']
        unique_together = ('almacen', 'medicamento', 'lote', 'fecha_caducidad')

    def __str__(self):
        nombre_almacen = self.almacen.nombre if self.almacen else "Sin Almacén"
        return f"[{nombre_almacen}] {self.medicamento.clave_sector} | Lote: {self.lote} | Disp: {self.cantidad_disponible}"
    
    def clean(self):
        super().clean()
        if self.lote and hasattr(self, 'medicamento') and self.medicamento:
            choque = Inventario.objects.filter(lote__iexact=self.lote).exclude(medicamento=self.medicamento).first()
            if choque:
                lab_original = choque.medicamento.fabricante or "Otro Laboratorio"
                raise ValidationError({
                    'lote': f"🚨 ALERTA DE TRAZABILIDAD: El lote '{self.lote}' ya pertenece al fabricante '{lab_original}'."
                })

class MovimientoKardex(models.Model):
    TIPO_MOVIMIENTO = [
        ('ENTRADA_COMPRA', '📥 Entrada por Compra'),
        ('SALIDA_OPM', '📤 Salida por OPM'),
        ('ENTRADA_DEVOLUCION', '↩️ Entrada por Devolución (Rechazo)'),
        ('TRASPASO_IN', '🔄 Entrada por Traspaso'),
        ('TRASPASO_OUT', '🔄 Salida por Traspaso'),
        ('MERMA', '❌ Merma / Caducidad'),
        ('AJUSTE', '⚙️ Ajuste Manual'),
    ]
    
    fecha = models.DateTimeField("Fecha y Hora", auto_now_add=True)
    almacen = models.ForeignKey('Almacen', on_delete=models.CASCADE, verbose_name="Almacén")
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave")
    lote = models.CharField("Lote Físico", max_length=50)
    
    tipo = models.CharField("Tipo de Movimiento", max_length=20, choices=TIPO_MOVIMIENTO)
    cantidad = models.PositiveIntegerField("Cantidad Movida")
    saldo_restante = models.IntegerField("Saldo Posterior", help_text="Cuántas piezas quedaron después de este movimiento")
    
    folio_documento = models.CharField("Folio Relacionado", max_length=100, help_text="Ej. OC-123, REM-456, TRAS-789")
    observaciones = models.TextField("Observaciones", blank=True, null=True)

    class Meta:
        verbose_name = "Movimiento de Kardex"
        verbose_name_plural = "Kardex (Historial Exacto)"
        ordering = ['-fecha'] 

    def __str__(self):
        signo = "+" if "ENTRADA" in self.tipo or "IN" in self.tipo else "-"
        return f"{self.fecha.strftime('%d/%m/%Y %H:%M')} | {self.tipo} | {self.lote} | {signo}{self.cantidad} pzs (Quedan: {self.saldo_restante})"

class TraspasoIntercompany(models.Model):
    almacen_origen = models.ForeignKey('Almacen', on_delete=models.PROTECT, related_name='traspasos_salida', verbose_name="Almacén Origen (El que Vende)")
    almacen_destino = models.ForeignKey('Almacen', on_delete=models.PROTECT, related_name='traspasos_entrada', verbose_name="Almacén Destino (El que Compra)")
    folio_factura = models.CharField("Folio de Orden / Factura Fiscal", max_length=100, help_text="Folio del documento fiscal que ampara este movimiento.")
    
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave a Transferir")
    lote = models.CharField("Lote Físico", max_length=50)
    cantidad = models.PositiveIntegerField("Cantidad de Piezas")
    precio_unitario = models.DecimalField("Precio Unitario Fiscal", max_digits=10, decimal_places=2, default=0.00, help_text="A cuánto se lo está vendiendo internamente.")
    
    fecha_operacion = models.DateTimeField("Fecha del Traspaso", auto_now_add=True)
    
    ESTATUS_TRASPASO = [
        ('BORRADOR', 'Borrador 📝 (No afecta inventario)'), 
        ('COMPLETADO', 'Completado ✅ (Inventario movido)')
    ]
    estatus = models.CharField(max_length=20, choices=ESTATUS_TRASPASO, default='BORRADOR', verbose_name="Estado del Traspaso")
    procesado = models.BooleanField(default=False, editable=False)

    class Meta:
        verbose_name = "Traspaso Interno"
        verbose_name_plural = "Traspasos Inter-Compañías"

    def __str__(self):
        return f"Traspaso {self.folio_factura}: {self.almacen_origen.nombre} ➔ {self.almacen_destino.nombre}"

    def clean(self):
        super().clean()
        if self.almacen_origen == self.almacen_destino:
            raise ValidationError("El almacén de origen y destino no pueden ser el mismo.")
        
        if self.estatus == 'COMPLETADO' and not self.procesado:
            stock_origen = Inventario.objects.filter(almacen=self.almacen_origen, medicamento=self.medicamento, lote__iexact=self.lote).first()
            if not stock_origen or stock_origen.cantidad_disponible < self.cantidad:
                raise ValidationError(f"🚨 IMPOSIBLE: El almacén '{self.almacen_origen.nombre}' no tiene {self.cantidad} piezas del lote '{self.lote}'.")

    def save(self, *args, **kwargs):
        if self.estatus == 'COMPLETADO' and not self.procesado:
            stock_origen = Inventario.objects.get(almacen=self.almacen_origen, medicamento=self.medicamento, lote__iexact=self.lote)
            stock_origen.cantidad_disponible -= self.cantidad
            stock_origen.save()
            
            stock_destino, creado = Inventario.objects.get_or_create(
                almacen=self.almacen_destino,
                medicamento=self.medicamento,
                lote=self.lote,
                defaults={
                    'fecha_caducidad': stock_origen.fecha_caducidad,
                    'cantidad_disponible': 0
                }
            )
            stock_destino.cantidad_disponible += self.cantidad
            stock_destino.save()
            
            self.procesado = True
            
        super().save(*args, **kwargs)

class Proveedor(models.Model):
    nombre = models.CharField("Razón Social", max_length=200)
    rfc = models.CharField("RFC", max_length=20, blank=True, null=True)
    contacto_principal = models.CharField(max_length=150, blank=True, null=True)
    correo_ventas = models.EmailField(blank=True, null=True)
    telefono = models.CharField(max_length=50, blank=True, null=True)
    dias_credito = models.PositiveIntegerField("Días de Crédito", default=30)

    class Meta:
        verbose_name = "Proveedor"
        verbose_name_plural = "Proveedores"

    def __str__(self):
        return self.nombre

class OrdenCompra(models.Model):
    ESTATUS_COMPRA = [
        ('BORRADOR', '📝 Borrador'),
        ('AUTORIZADA', '✅ Autorizada (Enviada)'),
        ('TRANSITO', '🚚 En Tránsito'),
        ('RECIBIDA', '📦 Recibida en Almacén'),
        ('CANCELADA', '❌ Cancelada'),
    ]

    folio = models.CharField("Folio OC", max_length=50, unique=True, help_text="Ej. OC-GAMS-2026-001")
    empresa_compradora = models.ForeignKey('Empresa', on_delete=models.CASCADE, verbose_name="Empresa que Compra")
    
    proveedor = models.ForeignKey('SocioComercial', on_delete=models.PROTECT, verbose_name="Proveedor (Socio Comercial)")
    
    fecha_emision = models.DateField(default=timezone.now)
    fecha_entrega_esperada = models.DateField()
    # 👇 NUEVOS CAMPOS PARA KPIs DE COMPRAS 👇
    fecha_necesidad = models.DateField("Fecha de Solicitud (Necesidad)", default=timezone.now, help_text="¿Cuándo se originó la necesidad interna?")
    fecha_recepcion_real = models.DateField("Fecha de Recepción Real", null=True, blank=True)
    es_compra_no_planeada = models.BooleanField("Gasto Maverick (Fuera de Contrato)", default=False, help_text="Marcar si es urgencia fuera de acuerdos/catálogos preaprobados.")
    costo_administrativo = models.DecimalField("Costo Administrativo de OC", max_digits=8, decimal_places=2, default=500.00, help_text="Costo interno de procesamiento ($)")
    
    estatus = models.CharField(max_length=20, choices=ESTATUS_COMPRA, default='BORRADOR')
    destino = models.CharField("Destino de Entrega", max_length=255, default="ALMACÉN GENERAL GPHARMA", help_text="Dirección o unidad donde se recibirá el medicamento")
    observaciones = models.TextField(blank=True, null=True)

    class Meta:
        verbose_name = "Orden de Compra"
        verbose_name_plural = "Órdenes de Compra (OC)"

    def __str__(self):
        return f"{self.folio} - {self.proveedor.nombre}"

    @property
    def total_compra(self):
        return sum(partida.importe for partida in self.partidas_compra.all())
    
    @property
    def penalizacion_calculada(self):
        from decimal import Decimal
        from django.utils import timezone
        
        if not hasattr(self, 'contrato') or not self.contrato:
            return Decimal('0.00')
            
        if not self.fecha_entrega_esperada:
            return Decimal('0.00')
            
        fecha_cierre = self.fecha_emision 
        if hasattr(self, 'fecha_recepcion_real') and self.fecha_recepcion_real:
            fecha_cierre = self.fecha_recepcion_real.date() if hasattr(self.fecha_recepcion_real, 'date') else self.fecha_recepcion_real
        else:
            fecha_cierre = timezone.now().date()
            
        dias_atraso = (fecha_cierre - self.fecha_entrega_esperada).days
        
        if dias_atraso <= 0:
            return Decimal('0.00') 
            
        porc_diario = self.contrato.porcentaje_penalizacion_diaria / Decimal('100')
        tope_maximo = self.contrato.tope_penalizacion / Decimal('100')
        
        castigo_porcentaje = Decimal(str(dias_atraso)) * porc_diario
        
        if castigo_porcentaje > tope_maximo:
            castigo_porcentaje = tope_maximo
            
        descuento = self.total_compra * castigo_porcentaje
        return round(descuento, 2)

class PartidaCompra(models.Model):
    orden = models.ForeignKey(OrdenCompra, related_name='partidas_compra', on_delete=models.CASCADE)
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.PROTECT)
    cantidad = models.PositiveIntegerField("Cantidad (Piezas)")
    precio_unitario = models.DecimalField("Costo Unitario", max_digits=12, decimal_places=2)
    precio_referencia = models.DecimalField("Precio Presupuestado/Histórico", max_digits=12, decimal_places=2, default=0.00, help_text="Para medir el ahorro (PPV).")
    piezas_rechazadas = models.PositiveIntegerField("Pzas Rechazadas (Calidad)", default=0, help_text="Piezas devueltas por defecto del proveedor.")
    
    cantidad_recibida = models.PositiveIntegerField("Pzas Recibidas", default=0, help_text="Se llena al llegar al almacén")

    class Meta:
        verbose_name = "Partida"
        verbose_name_plural = "Partidas"

    @property
    def importe(self):
        return self.cantidad * self.precio_unitario

class DocumentoOrdenCompra(models.Model):
    orden = models.ForeignKey(OrdenCompra, related_name='documentos', on_delete=models.CASCADE)
    archivo = models.FileField(upload_to='ordenes_compra/evidencias/', verbose_name="Archivo Adjunto")
    descripcion = models.CharField("Descripción / Título", max_length=150, help_text="Ej. Cotización del proveedor, Autorización del jefe, etc.")
    fecha_subida = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Documento de Evidencia"
        verbose_name_plural = "📂 Documentos y Evidencias"

    def __str__(self):
        return f"{self.descripcion} - {self.orden.folio}"

class EntradaAlmacen(models.Model):
    orden = models.ForeignKey(OrdenCompra, on_delete=models.CASCADE, verbose_name="Orden de Compra Vinculada")
    almacen_destino = models.ForeignKey('Almacen', on_delete=models.PROTECT, verbose_name="Almacén Destino", null=True, blank=True)
    folio_remision = models.CharField("Folio de Remisión / Factura", max_length=100, default="S/F", help_text="Anota el número de ticket o factura.")
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave / Medicamento")
    tipo_producto = models.CharField(max_length=20, choices=TIPO_PRODUCTO_CHOICES, default='COMERCIAL', verbose_name="Tipo de Producto")
    
    cantidad_recibida = models.PositiveIntegerField("Piezas Físicas Recibidas")
    
    # 👇 ESTE ES EL NUEVO CAMPO DE CALIDAD 👇
    piezas_rechazadas = models.PositiveIntegerField("Piezas Rechazadas (Mal estado)", default=0, help_text="Anota si llegaron cajas rotas, caducadas o diferentes.") 
    
    lote = models.CharField("Lote Impreso", max_length=50)
    fecha_caducidad = models.DateField("Fecha de Caducidad")
    
    documentacion_completa = models.BooleanField("¿Llegó con Remisión y Certificados Analíticos?", default=True)
    ubicacion = models.CharField("Ubicación Física (Ej. Pasillo 3, Anaquel B)", max_length=150, help_text="¿En qué parte del almacén se guardará?")
    observaciones_calidad = models.TextField("Observaciones de Calidad", blank=True, null=True)
    acuse_recibo = models.FileField("Acuse de Recibo / Remisión Firmada", upload_to='almacen/acuses/', blank=True, null=True, help_text="Sube el PDF o Foto del acuse firmado al recibir la mercancía.")
    factura_proveedor = models.FileField("Factura del Proveedor (PDF/XML)", upload_to='almacen/facturas/', blank=True, null=True, help_text="Sube la factura correspondiente a esta entrega.")
    fecha_ingreso = models.DateTimeField("Fecha y Hora de Recepción", auto_now_add=True)

    class Meta:
        verbose_name = "Entrada de Almacén"
        verbose_name_plural = "Recepciones de Almacén"

    def __str__(self):
        return f"Remisión: {self.folio_remision} | OC-{self.orden.folio} | {self.cantidad_recibida} pzas"

    def clean(self):
        super().clean()
        if self.lote and hasattr(self, 'medicamento') and self.medicamento:
            query = EntradaAlmacen.objects.filter(lote__iexact=self.lote).exclude(medicamento=self.medicamento)
            if self.pk:
                query = query.exclude(pk=self.pk)
            choque = query.first()
            if choque:
                lab_original = choque.medicamento.fabricante or "Otro Laboratorio"
                marca_original = choque.medicamento.denominacion_distintiva or "Genérico"
                raise ValidationError({
                    'lote': f"🚨 BARRERA DE SEGURIDAD: El lote '{self.lote}' ya fue ingresado antes como '{marca_original}' ({lab_original})."
                })

    def save(self, *args, **kwargs):
        es_nuevo = self.pk is None
        super().save(*args, **kwargs)

        if es_nuevo:
            inventario_obj, creado = Inventario.objects.get_or_create(
                almacen=self.almacen_destino,
                medicamento=self.medicamento,
                lote=self.lote,
                fecha_caducidad=self.fecha_caducidad,
                tipo_producto=self.tipo_producto,
                defaults={
                    'cantidad_disponible': 0,
                    'fecha_ingreso': self.fecha_ingreso.date()
                }
            )
            # Solo metemos al almacén lo que SÍ se recibió
            inventario_obj.cantidad_disponible += self.cantidad_recibida
            inventario_obj.save()

            if self.orden.estatus not in ['RECIBIDA', 'CANCELADA']:
                self.orden.estatus = 'PARCIAL' 
                self.orden.save()

            # 👇 MAGIA DE CALIDAD: REPORTARLE A COMPRAS EL DEFECTO 👇
            if self.piezas_rechazadas > 0:
                partida_oc = self.orden.partidas_compra.filter(medicamento=self.medicamento).first()
                if partida_oc:
                    partida_oc.piezas_rechazadas += self.piezas_rechazadas
                    partida_oc.save()

            from .models import MovimientoKardex
            MovimientoKardex.objects.create(
                almacen=self.almacen_destino,
                medicamento=self.medicamento,
                lote=self.lote,
                tipo='ENTRADA_COMPRA',
                cantidad=self.cantidad_recibida,
                saldo_restante=inventario_obj.cantidad_disponible,
                folio_documento=self.folio_remision,
                observaciones=f"Recepcionada desde OC-{self.orden.folio}"
            )

class ConfiguracionEmail(models.Model):
    EMPRESA_CHOICES = [
        ('SAGO', 'Sago Medical'),
        ('GSM', 'GSM'),
        ('GAMS', 'GAMS'),
    ]
    empresa = models.CharField(max_length=20, choices=EMPRESA_CHOICES, unique=True)
    email_host_user = models.EmailField()
    email_host_password = models.CharField(max_length=100, help_text="Usa la Contraseña de Aplicación de 16 letras")

    class Meta:
        verbose_name = "Configuración de Correo"
        verbose_name_plural = "Configuraciones de Correos"

    def __str__(self):
        return f"Correo para {self.empresa}"
    
class EscanerKardex(Inventario):
    class Meta:
        proxy = True
        verbose_name = "📱 Escáner Kardex (Cámara)"
        verbose_name_plural = "Escáner Kardex"

# ==========================================
# 🚀 MÓDULO: COTIZACIONES Y VENTAS DIRECTAS
# ==========================================
class Cotizacion(models.Model):
    TIPO_PROCEDIMIENTO = [
        ('COTIZACION_PRIVADA', 'Cotización a Cliente Privado'),
        ('ADJUDICACION', 'Adjudicación Directa (Gobierno)'),
        ('INVESTIGACION', 'Investigación de Mercado (Gobierno)'),
        ('INVITACION_3', 'Invitación a 3 Personas (Gobierno)'),
    ]
    
    ESTATUS_COTIZACION = [
        ('BORRADOR', 'Borrador (En edición)'),
        ('ENVIADA', 'Enviada al Cliente / Dependencia'),
        ('GANADA', 'Ganada / Aprobada (Lista para Pedido)'),
        ('PERDIDA', 'Perdida / Rechazada'),
        ('CANCELADA', 'Cancelada'),
    ]

    tipo_procedimiento = models.CharField(max_length=30, choices=TIPO_PROCEDIMIENTO, default='COTIZACION_PRIVADA', verbose_name="Tipo de Venta")
    
    # 👇 NUEVO: Empresa emisora
    empresa = models.ForeignKey('Empresa', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Empresa Emisora")
    
    # 👇 MODIFICADO: Folio ahora permite blancos para autogenerarse
    folio = models.CharField(max_length=50, unique=True, blank=True, verbose_name="Folio de Cotización / Evento", help_text="Déjalo en blanco para auto-generar el consecutivo (Ej: Sago.Cot-001).")
    
    razon_social = models.CharField(max_length=200, blank=True, null=True, verbose_name="Razón Social (Cliente Privado)")
    dependencia = models.CharField(max_length=100, choices=DEPENDENCIAS_MAESTRAS, blank=True, null=True, verbose_name="Dependencia (Gobierno)")
    
    fecha_apertura = models.DateTimeField(verbose_name="Fecha y hora de apertura (Propuestas)", null=True, blank=True)
    fecha_emision = models.DateField(default=timezone.now, verbose_name="Fecha de Emisión")
    vigencia_dias = models.IntegerField(default=15, verbose_name="Días de Vigencia")
    
    estatus = models.CharField(max_length=20, choices=ESTATUS_COTIZACION, default='BORRADOR', verbose_name="Estatus Comercial")

    class Meta:
        verbose_name = "Cotización / Venta Directa"
        verbose_name_plural = "Cotizaciones y Ventas Directas"
        ordering = ['-fecha_emision']

    def __str__(self):
        cliente = self.razon_social or self.get_dependencia_display() or "Sin Cliente"
        return f"{self.folio} | {cliente}"

    # 🧠 MAGIA: Autogenerador de Folios Consecutivos
    def save(self, *args, **kwargs):
        if not self.folio and self.empresa:
            nombre_empresa = self.empresa.nombre.upper()
            
            if "SAGO" in nombre_empresa: prefijo = "Sago"
            elif "GSM" in nombre_empresa: prefijo = "Gsm"
            elif "GAMS" in nombre_empresa: prefijo = "Gams"
            else: prefijo = "Gen"

            ultimo = Cotizacion.objects.filter(folio__startswith=f"{prefijo}.Cot-").order_by('-folio').first()
            
            if ultimo and '-' in ultimo.folio:
                try:
                    num = int(ultimo.folio.split('-')[-1])
                    siguiente = num + 1
                except ValueError:
                    siguiente = 1
            else:
                siguiente = 1
                
            self.folio = f"{prefijo}.Cot-{siguiente:03d}"
            
        elif not self.folio and not self.empresa:
            import random
            self.folio = f"COT-TEMP-{random.randint(1000,9999)}"

        super().save(*args, **kwargs)

    @property
    def total_cotizacion(self):
        return sum(partida.importe for partida in self.partidas_cotizacion.all())

class PartidaCotizacion(models.Model):
    cotizacion = models.ForeignKey(Cotizacion, on_delete=models.CASCADE, related_name='partidas_cotizacion')
    medicamento = models.ForeignKey(CatalogoMedicamento, on_delete=models.CASCADE, verbose_name="Medicamento (Catálogo Libre)")
    cantidad = models.IntegerField(verbose_name="Cantidad Solicitada")
    precio_unitario = models.DecimalField(max_digits=12, decimal_places=2, default=0.00, verbose_name="Precio Unitario Ofertado")

    @property
    def importe(self):
        return float(self.cantidad or 0) * float(self.precio_unitario or 0)

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs)
        
        if self.cotizacion.estatus == 'GANADA' and self.precio_unitario and self.precio_unitario > 0:
            proveedor_nombre = 'Sin Laboratorio'
            if self.medicamento.socio_contacto:
                proveedor_nombre = self.medicamento.socio_contacto.nombre
            elif self.medicamento.fabricante:
                proveedor_nombre = self.medicamento.fabricante

            institucion_cliente = self.cotizacion.get_dependencia_display() if self.cotizacion.dependencia else (self.cotizacion.razon_social or 'Privado')

            HistorialPrecio.objects.update_or_create(
                procedimiento=self.cotizacion.folio,
                clave=self.medicamento.clave_sector,
                defaults={
                    'medicamento': self.medicamento,
                    'descripcion': self.medicamento.descripcion,
                    'tipo_procedimiento': self.cotizacion.get_tipo_procedimiento_display(),
                    'proveedor': proveedor_nombre,
                    'institucion': institucion_cliente,
                    'institucion_desagregada': self.cotizacion.dependencia or self.cotizacion.razon_social or 'Privado',
                    'precio': self.precio_unitario,
                    'cantidad': self.cantidad or 0,
                    'valor': (self.cantidad or 0) * self.precio_unitario
                }
            )


    class Meta:
        verbose_name = "Partida Cotizada"
        verbose_name_plural = "Medicamentos Cotizados"

    def __str__(self):
        return str(self.medicamento.clave_sector)
    
# ==========================================
# 🚀 3.1 MÓDULO: PEDIDOS DIRECTOS (PROXY)
# ==========================================
class PedidoDirecto(OrdenSuministro):
    class Meta:
        proxy = True
        verbose_name = "Pedido Directo (10 días)"
        verbose_name_plural = "Pedidos Directos"

    def save(self, *args, **kwargs):
        self.tipo_documento = 'PEDIDO'
        if not self.fecha_recepcion:
            self.fecha_recepcion = timezone.now().date()
        if not self.fecha_limite:
            self.fecha_limite = self.fecha_recepcion + timedelta(days=10)
            
        super().save(*args, **kwargs)

        # ==========================================
# 🛑 MÓDULO DE CUARENTENA, MERMAS Y DEVOLUCIONES
# ==========================================
class IncidenciaInventario(models.Model):
    MOTIVO_CHOICES = [
        ('PROXIMO_CADUCAR', '⏳ Próximo a Caducar'),
        ('VICIO_OCULTO', '🕵️‍♂️ Vicio Oculto / Defecto de Fábrica'),
        ('DANADO', '💥 Dañado en Almacén o Transporte'),
    ]
    RESOLUCION_CHOICES = [
        ('EN_CUARENTENA', '🔒 En Cuarentena (Esperando decisión)'),
        ('DONATIVO', '🎁 Donativo (Salida sin cobro)'),
        ('DEVOLUCION', '🚚 Devolución a Proveedor (Garantía)'),
        ('DESTRUCCION', '🔥 Destrucción / Merma (Pérdida)'),
    ]
    
    inventario_origen = models.ForeignKey('Inventario', on_delete=models.SET_NULL, null=True, verbose_name="Lote Original en Inventario", help_text="Selecciona el lote del inventario que vas a mandar a Cuarentena.")
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.CASCADE, verbose_name="Clave / Medicamento")
    lote = models.CharField("Lote Afectado", max_length=50)
    cantidad_afectada = models.PositiveIntegerField("Piezas Afectadas (Se restarán del stock)")
    
    motivo = models.CharField("Motivo del Problema", max_length=20, choices=MOTIVO_CHOICES)
    resolucion = models.CharField("Decisión Final (Destino)", max_length=20, choices=RESOLUCION_CHOICES, default='EN_CUARENTENA')
    
    # FINANZAS Y GARANTÍAS
    socio_comercial = models.ForeignKey('SocioComercial', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Proveedor (Socio Comercial)", help_text="¿A quién le vamos a reclamar?")
    genera_nota_credito = models.BooleanField("¿Genera Nota de Crédito a nuestro favor?", default=False)
    monto_nota_credito = models.DecimalField("Monto Nota de Crédito ($)", max_digits=12, decimal_places=2, default=0.00)
    
    aplica_penalizacion = models.BooleanField("¿Aplicar Multa/Penalización al Proveedor?", default=False)
    monto_penalizacion = models.DecimalField("Monto Penalización ($)", max_digits=12, decimal_places=2, default=0.00)
    
    observaciones = models.TextField("Observaciones / Detalles del defecto", blank=True, null=True)
    fecha_registro = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Control de Cuarentena y Merma"
        verbose_name_plural = "Cuarentenas y Devoluciones"

    def __str__(self):
        return f"Incidencia {self.id} | {self.medicamento.clave_sector} | {self.get_motivo_display()}"

    def save(self, *args, **kwargs):
        es_nuevo = self.pk is None
        
        # Si no han llenado el medicamento o lote, lo jalamos del inventario origen
        if self.inventario_origen:
            if not self.medicamento_id:
                self.medicamento = self.inventario_origen.medicamento
            if not self.lote:
                self.lote = self.inventario_origen.lote
                
        super().save(*args, **kwargs)
        
        # 👇 MAGIA: DESCONTAR DEL INVENTARIO SANO AUTOMÁTICAMENTE 👇
        if es_nuevo and self.inventario_origen:
            self.inventario_origen.cantidad_disponible -= self.cantidad_afectada
            if self.inventario_origen.cantidad_disponible < 0:
                self.inventario_origen.cantidad_disponible = 0
            self.inventario_origen.save()
            
            # Dejamos la huella de auditoría en el Kardex
            from .models import MovimientoKardex
            MovimientoKardex.objects.create(
                almacen=self.inventario_origen.almacen,
                medicamento=self.medicamento,
                lote=self.lote,
                tipo='MERMA',
                cantidad=self.cantidad_afectada,
                saldo_restante=self.inventario_origen.cantidad_disponible,
                folio_documento=f"INC-{self.id}",
                observaciones=f"Aislado en CUARENTENA. Motivo: {self.get_motivo_display()}"
            )

from django.contrib.auth.models import User, Group # <--- Agregamos Group aquí

class PerfilEquipo(models.Model):
    ROLES = [
        # --- DIRECCIÓN ---
        ('DIRECCION_GRAL', 'Dirección General'),

        # --- ÁREA COMERCIAL Y LICITACIONES ---
        ('DIRECCION_COMER', 'Dirección Comercial'),
        ('GERENCIA_COMERCIAL', 'Gerencia Comercial'),
        ('COORD_COMERCIAL', 'Coordinador Comercial'),
        ('ANALISTA_LICITA', 'Analista de Licitaciones'),
        ('ATENCION_COMERCIAL', 'Atención Comercial / Ventas'),
        ('ANALISTA_ATENCION', 'Analista Atención Comercial'),
        ('COMPRADOR', 'Ejecutivo de Compras'), # <--- Lo movimos a su casa correcta

        # --- ÁREA LOGÍSTICA Y OPERACIONES ---
        ('GERENTE_LOGISTICA', 'Gerente de Logística'),
        ('AUX_ALMACEN', 'Auxiliar de Almacén / Chofer'),

        # --- ÁREA COMPRAS Y FINANZAS ---
        ('DIRECCION_ADMIN', 'Dirección Administrativa'),


        # --- ÁREA COBRANZA ---
        ('DIRECCION_COBR', 'Dirección Cobranza'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='perfil')
    whatsapp = models.CharField(max_length=20, help_text="Ej: +5215540741751")
    rol = models.CharField(max_length=20, choices=ROLES, default='COORD_COMERCIAL')
    activo = models.BooleanField(default=True)

    def __str__(self):
        return f"{self.user.get_full_name()} - {self.rol}"

    # 👇 AQUÍ EMPIEZA LA MAGIA DEL PILOTO AUTOMÁTICO 👇
    def save(self, *args, **kwargs):
        # 1. Guardamos el perfil normalmente
        super().save(*args, **kwargs)
        
        # 2. El "Diccionario" que mapea tu organigrama con los Megagrupos
        mapa_megagrupos = {
            'DIRECCION_GRAL': 'DIRECCION_GLOBAL',
            
            'DIRECCION_COMER': 'COMERCIAL',
            'GERENCIA_COMERCIAL': 'COMERCIAL',
            'COORD_COMERCIAL': 'COMERCIAL',
            'ANALISTA_LICITA': 'COMERCIAL',
            'ATENCION_COMERCIAL': 'COMERCIAL',
            'ANALISTA_ATENCION': 'COMERCIAL',
            'COMPRADOR': 'COMERCIAL',
            
            'GERENTE_LOGISTICA': 'LOGISTICA',
            'AUX_ALMACEN': 'LOGISTICA',
            
            'DIRECCION_ADMIN': 'COMPRAS_FINANZAS',
            
            'DIRECCION_COBR': 'COBRANZA'
        }
        
        nombre_megagrupo = mapa_megagrupos.get(self.rol)
        
        if nombre_megagrupo:
            # 3. Le quitamos permisos viejos (por si lo cambiaste de puesto)
            self.user.groups.clear()
            
            # 4. Buscamos el Megagrupo (o lo creamos si no existe)
            grupo, creado = Group.objects.get_or_create(name=nombre_megagrupo)
            
            # 5. Metemos al usuario a su nueva "casa" y le damos acceso al sistema
            self.user.groups.add(grupo)
            self.user.is_staff = True 
            self.user.save()

            # ==========================================
# 📊 MÓDULO DE INTELIGENCIA: HISTORIAL DE PRECIOS
# ==========================================
class HistorialPrecio(models.Model):
    medicamento = models.ForeignKey('CatalogoMedicamento', on_delete=models.SET_NULL, null=True, blank=True)
    clave = models.CharField("Clave", max_length=50)
    descripcion = models.TextField("Descripción")
    grupo_terapeutico = models.CharField("Grupo Terapéutico", max_length=150, default="No Especificado")
    
    procedimiento = models.CharField("Procedimiento", max_length=100)
    tipo_procedimiento = models.CharField("Tipo de Procedimiento", max_length=100)
    
    proveedor = models.CharField("Proveedor / Laboratorio", max_length=200)
    institucion = models.CharField("Institución", max_length=150)
    institucion_desagregada = models.CharField("Institución Desagregada", max_length=150, blank=True, null=True)
    
    precio = models.DecimalField("Precio Unitario", max_digits=12, decimal_places=2)
    cantidad = models.PositiveIntegerField("Cantidad", default=0)
    valor = models.DecimalField("Valor Total", max_digits=15, decimal_places=2)
    
    fecha_captura = models.DateTimeField("Fecha de Registro", auto_now_add=True)

    class Meta:
        verbose_name = "Historial de Precio"
        verbose_name_plural = "Historial de Precios"
        ordering = ['-fecha_captura']

    def __str__(self):
        return f"{self.clave} - ${self.precio}"