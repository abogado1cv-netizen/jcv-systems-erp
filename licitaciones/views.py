import json
import csv
from datetime import timedelta
from django.http import HttpResponse
from django.shortcuts import render
from django.db.models import Sum, F, Q, Count, Avg
from django.db.models.functions import TruncMonth, Coalesce
from django.utils import timezone
from django.contrib.admin.views.decorators import staff_member_required
from .models import Inventario
from .services import DashboardService
from django.db.models import Q

# Traemos todos los modelos
from .models import (
    Contrato, Licitacion, PartidaRequerimiento, Empresa, 
    OrdenSuministro, RemisionEntrega, ClaveContrato, CatalogoMedicamento
)

# ==========================================
# 1. DASHBOARD DE CONTRATOS (COMERCIAL VS LOGÍSTICO)
# ==========================================
@staff_member_required
def dashboard_contratos(request):
    contratos = Contrato.objects.all()
    
    # 1. LEEMOS LA BÚSQUEDA Y LOS FILTROS
    busqueda = request.GET.get('q', '').strip() 
    filtro_dependencia = request.GET.get('dependencia', '')
    filtro_empresa_id = request.GET.get('empresa', '')
    filtro_contrato = request.GET.get('contrato', '')

    # 1.5 APLICAMOS LA BÚSQUEDA GLOBAL
    if busqueda:
        contratos = contratos.filter(
            Q(numero_contrato__icontains=busqueda) |
            Q(dependencia__icontains=busqueda) |
            Q(licitacion_origen__num_procedimiento__icontains=busqueda)
        ).distinct()

    # APLICAMOS LOS FILTROS DESPLEGABLES
    if filtro_dependencia:
        contratos = contratos.filter(dependencia=filtro_dependencia)
    if filtro_empresa_id.isdigit():
        contratos = contratos.filter(empresa_id=filtro_empresa_id)
    if filtro_contrato:
        contratos = contratos.filter(numero_contrato=filtro_contrato)

    # 2. OBTENEMOS TODAS LAS CLAVES DE ESOS CONTRATOS FILTRADOS
    claves_qs = ClaveContrato.objects.filter(contrato__in=contratos)

    # 3. CALCULAMOS MONTOS Y PIEZAS MÁXIMAS DEL CONTRATO
    agregados = claves_qs.aggregate(
        min_tot=Sum(F('cantidad_minima') * F('precio_neto')),
        max_tot=Sum(F('cantidad_maxima') * F('precio_neto')),
        pzas_min=Sum('cantidad_minima'),
        pzas_max=Sum('cantidad_maxima')
    )
    
    monto_minimo = agregados.get('min_tot') or 0
    monto_maximo = agregados.get('max_tot') or 0
    piezas_minimas = agregados.get('pzas_min') or 0
    piezas_maximas = agregados.get('pzas_max') or 0

    # =========================================================
    # 4. CALCULAMOS EL AVANCE COMERCIAL Y LOGÍSTICO (Piezas y Montos)
    # =========================================================
    
    # 4.1 Piezas y Monto Solicitado (OPMs)
    agg_solicitadas = OrdenSuministro.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
        tot_pzas=Sum('cantidad_solicitada'),
        tot_dinero=Sum(F('cantidad_solicitada') * F('clave_contrato__precio_neto'))
    )
    piezas_solicitadas = agg_solicitadas.get('tot_pzas') or 0
    monto_solicitado = agg_solicitadas.get('tot_dinero') or 0

    # 4.2 Piezas y Monto Entregado (Leemos directo de las OPMs)
    agg_entregadas = OrdenSuministro.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
        tot_pzas=Sum('cantidad_entregada'),
        tot_dinero=Sum(F('cantidad_entregada') * F('clave_contrato__precio_neto'))
    )
    piezas_entregadas = agg_entregadas.get('tot_pzas') or 0
    monto_entregado = agg_entregadas.get('tot_dinero') or 0

    # Cálculos para gráficas y porcentajes
    piezas_pendientes_solicitar = piezas_maximas - piezas_solicitadas
    if piezas_pendientes_solicitar < 0: piezas_pendientes_solicitar = 0

    piezas_pendientes_entregar = piezas_solicitadas - piezas_entregadas
    if piezas_pendientes_entregar < 0: piezas_pendientes_entregar = 0

    avance_min_pct = (piezas_solicitadas / piezas_minimas * 100) if piezas_minimas > 0 else 0
    avance_max_pct = (piezas_solicitadas / piezas_maximas * 100) if piezas_maximas > 0 else 0

    # 4.3 CALCULAMOS PENALIZACIONES
    ordenes_vinculadas = OrdenSuministro.objects.filter(clave_contrato__contrato__in=contratos)
    total_penalizado = sum(float(o.penalizacion_estimada) for o in ordenes_vinculadas)


    # 5. TOP 5 CLAVES CON MAYOR ASIGNACIÓN
    top_claves = claves_qs.annotate(
        importe_max=F('cantidad_maxima') * F('precio_neto')
    ).values('medicamento__clave_sector').annotate(
        total_asignado=Sum('importe_max')
    ).order_by('-total_asignado')[:5]
    
    nombres_top = [c['medicamento__clave_sector'] for c in top_claves]
    montos_top = [float(c['total_asignado']) for c in top_claves]

    # 6. DETALLE DE CLAVES PARA LA TABLA INFERIOR
    detalle_claves_qs = claves_qs.select_related('medicamento', 'contrato').annotate(
        importe_max=F('cantidad_maxima') * F('precio_neto')
    ).order_by('-importe_max')[:100]
    
    detalle_claves = []
    for clave in detalle_claves_qs:
        # 6.1 Sumamos las piezas ya entregadas (Leemos de las OPMs directamente)
        entregas_dict = OrdenSuministro.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_entregada'))
        clave.pzas_entregadas = entregas_dict.get('tot') or 0

        # 6.2 Sumamos las piezas solicitadas (OPMs)
        solicitadas_dict = OrdenSuministro.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_solicitada'))
        clave.pzas_solicitadas = solicitadas_dict.get('tot') or 0
        
        # 6.3 --- NUEVO: CALCULAMOS LAS PIEZAS FALTANTES ---
        faltantes = clave.pzas_solicitadas - clave.pzas_entregadas
        clave.pzas_faltantes = faltantes if faltantes > 0 else 0
        # --------------------------------------------------

        detalle_claves.append(clave)

    # 7. FILTROS EN CASCADA CORREGIDOS
    deps_list = Contrato.objects.values_list('dependencia', flat=True).distinct()
    
    empresas_qs = Empresa.objects.filter(contrato__isnull=False).distinct()
    if filtro_dependencia:
        empresas_qs = empresas_qs.filter(contrato__dependencia=filtro_dependencia)
        
    contratos_qs = Contrato.objects.all()
    if filtro_dependencia: contratos_qs = contratos_qs.filter(dependencia=filtro_dependencia)
    if filtro_empresa_id.isdigit(): contratos_qs = contratos_qs.filter(empresa_id=filtro_empresa_id)
    contratos_list = contratos_qs.values_list('numero_contrato', flat=True).distinct()

    context = {
        # KPIs Financieros Principales
        'monto_minimo_str': f"${monto_minimo:,.2f}",
        'monto_maximo_str': f"${monto_maximo:,.2f}",
        'monto_solicitado_str': f"${monto_solicitado:,.2f}",
        'monto_entregado_str': f"${monto_entregado:,.2f}",
        'piezas_minimas_str': f"{piezas_minimas:,}",
        'piezas_maximas_str': f"{piezas_maximas:,}",
        'total_penalizado_str': f"-${total_penalizado:,.2f}",
        
        # Nuevas variables para las Donas y Tarjetas
        'piezas_solicitadas': piezas_solicitadas,
        'piezas_pendientes_solicitar': piezas_pendientes_solicitar,
        'piezas_entregadas': piezas_entregadas,
        'piezas_pendientes_entregar': piezas_pendientes_entregar,
        
        'avance_min_pct': f"{avance_min_pct:,.1f}%",
        'avance_max_pct': f"{avance_max_pct:,.1f}%",
        
        # Gráficas
        'nombres_top_json': json.dumps(nombres_top),
        'montos_top_json': json.dumps(montos_top),
        
        # Tablas
        'detalle_claves': detalle_claves,
        
        # Filtros y Búsqueda
        'dependencias_disponibles': deps_list,
        'empresas_disponibles': empresas_qs,
        'contratos_disponibles': contratos_list,
        'filtro_dependencia': filtro_dependencia,
        'filtro_empresa': int(filtro_empresa_id) if filtro_empresa_id.isdigit() else '',
        'filtro_contrato': filtro_contrato,
        'busqueda': busqueda, 
    }
    
    return render(request, 'dashboard_contratos.html', context)


# ==========================================
# 2. DASHBOARD DE LICITACIONES
# ==========================================
@staff_member_required
def dashboard_licitaciones(request):
    # 1. Atrapamos lo que el usuario escribió o seleccionó
    q = request.GET.get('q', '').strip()
    filtro_empresa = request.GET.get('empresa', '')

    # 2. Traemos todo de la base de datos
    licitaciones = Licitacion.objects.all()
    partidas = PartidaRequerimiento.objects.all()

    # 3. FILTRO POR EMPRESA (El nuevo menú desplegable)
    if filtro_empresa:
        licitaciones = licitaciones.filter(empresa_id=filtro_empresa)
        partidas = partidas.filter(licitacion__empresa_id=filtro_empresa)

    # 4. SI HAY BÚSQUEDA DE TEXTO, FILTRAMOS
    if q:
        licitaciones = licitaciones.filter(
            Q(num_procedimiento__icontains=q) | 
            Q(dependencia__icontains=q)
        )
        partidas = partidas.filter(
            Q(licitacion__num_procedimiento__icontains=q) |
            Q(medicamento__clave_sector__icontains=q) |
            Q(medicamento__fabricante__icontains=q) |
            Q(resultado__icontains=q)
        )

    # 5. Ahora sí, calculamos con la info filtrada
    total_licitaciones = licitaciones.count()
    en_proceso = licitaciones.filter(estatus__estado='EN_PROCESO').count()
    adjudicadas = licitaciones.filter(estatus__estado='ADJUDICADO').count()
    perdidas = licitaciones.filter(estatus__estado='PERDIDO').count()

    monto_total = 0
    monto_ganado = 0
    monto_perdido = 0
    claves_participadas = set()
    claves_ganadas = set()
    
    for p in partidas:
        importe = (p.cantidad_maxima or 0) * (p.precio or 0)
        monto_total += importe
        if p.medicamento_id: claves_participadas.add(p.medicamento_id)
        
        if p.resultado == 'Asignada':
            monto_ganado += importe
            if p.medicamento_id: claves_ganadas.add(p.medicamento_id)
        elif p.resultado in ['Perdida por precio', 'Perdida técnicamente']:
            monto_perdido += importe

    monto_en_proceso = monto_total - monto_ganado - monto_perdido
    if monto_en_proceso < 0: monto_en_proceso = 0

    top_claves = partidas.filter(resultado='Asignada').annotate(
        importe_total=F('cantidad_maxima') * F('precio')
    ).values('medicamento__clave_sector', 'medicamento__fabricante').annotate(
        total_importe=Sum('importe_total')
    ).order_by('-total_importe')[:5]

    proximos_fallos = licitaciones.filter(
        estatus__estado='EN_PROCESO', fecha_fallo__gte=timezone.now()
    ).order_by('fecha_fallo')[:5]

    detalle_ganadas = partidas.filter(resultado='Asignada').select_related('licitacion', 'medicamento').annotate(
        importe_ganado=F('cantidad_maxima') * F('precio')
    ).order_by('-licitacion__fecha_fallo')[:50]

    detalle_perdidas = partidas.filter(resultado__in=['Perdida por precio', 'Perdida técnicamente']).select_related('licitacion', 'medicamento').order_by('-licitacion__fecha_fallo')[:50]

    nombres_top = [c['medicamento__clave_sector'] for c in top_claves]
    montos_top = [float(c['total_importe']) for c in top_claves]

    # Traemos las empresas para pintar el menú desplegable en el HTML
    empresas_disponibles = Empresa.objects.all()

    context = {
        'title': 'Panel Ejecutivo Compacto',
        'q': q, 
        'filtro_empresa': int(filtro_empresa) if filtro_empresa.isdigit() else '',
        'empresas_disponibles': empresas_disponibles,
        
        'monto_total_str': f"${monto_total:,.2f}",
        'monto_ganado_str': f"${monto_ganado:,.2f}",
        'monto_perdido_str': f"${monto_perdido:,.2f}",
        'monto_en_proceso_str': f"${monto_en_proceso:,.2f}",
        
        'monto_ganado_raw': monto_ganado,
        'monto_perdido_raw': monto_perdido,
        'monto_en_proceso_raw': monto_en_proceso,
        'nombres_top_json': json.dumps(nombres_top),
        'montos_top_json': json.dumps(montos_top),
        
        'total_licitaciones': total_licitaciones,
        'adjudicadas': adjudicadas,
        'perdidas': perdidas,
        'en_proceso': en_proceso,
        'total_claves_participadas': len(claves_participadas),
        'total_claves_ganadas': len(claves_ganadas),
        'top_claves': top_claves,
        'proximos_fallos': proximos_fallos,
        'detalle_ganadas': detalle_ganadas,
        'detalle_perdidas': detalle_perdidas,
    }
        
    return render(request, 'dashboard_licitaciones.html', context)

# ==========================================
# 3. DASHBOARD DE ÓRDENES DE SUMINISTRO (LOGÍSTICA)
# ==========================================
from django.db.models import Sum, Count, Q
from django.utils import timezone
import json

@staff_member_required
def dashboard_ordenes(request):
    busqueda = request.GET.get('q', '').strip()
    
    ordenes = OrdenSuministro.objects.all()
    
    if busqueda:
        ordenes = ordenes.filter(
            Q(numero_orden_suministro__icontains=busqueda) |
            Q(nombre_unidad__icontains=busqueda) |
            Q(clave_contrato__medicamento__clave_sector__icontains=busqueda) |
            Q(razon_social__icontains=busqueda)
        ).distinct()

    # KPIs Básicos (Las Órdenes)
    total_ordenes = ordenes.count()
    entregadas = ordenes.filter(estatus='ENTREGADA').count()
    pendientes = ordenes.filter(estatus__in=['PENDIENTE', 'PARCIAL']).count()
    
    # Atrasadas (Fecha límite menor a hoy y no entregadas)
    hoy = timezone.now().date()
    atrasadas = ordenes.filter(estatus__in=['PENDIENTE', 'PARCIAL'], fecha_limite__lt=hoy).count()

    # Cálculos Financieros (Iteramos porque son @properties)
    monto_total_solicitado = sum((float(o.cantidad_solicitada or 0) * float(o.precio_unitario or 0)) for o in ordenes)
    penalizaciones_totales = sum(float(o.penalizacion_estimada) for o in ordenes)

    # ==========================================
    # 🔥 NUEVOS CÁLCULOS: INTELIGENCIA DE PIEZAS
    # ==========================================
    # 1. Total de Piezas Solicitadas
    total_piezas_solicitadas = ordenes.aggregate(total=Sum('cantidad_solicitada'))['total'] or 0
    
    # 2. Total de Piezas Entregadas
    total_piezas_entregadas = ordenes.aggregate(total=Sum('cantidad_entregada'))['total'] or 0
    
    # 3. Total de Piezas Pendientes (Calculado)
    total_piezas_pendientes = total_piezas_solicitadas - total_piezas_entregadas
    if total_piezas_pendientes < 0:
        total_piezas_pendientes = 0 # Protección por si hay entregas de más
        
    # 4. Total de Piezas Canceladas (Asumiendo que tienes un estatus CANCELADA)
    total_piezas_canceladas = ordenes.filter(estatus='CANCELADA').aggregate(total=Sum('cantidad_solicitada'))['total'] or 0

    # ==========================================
    # 🔥 NUEVO TOP: CLAVES MÁS ENTREGADAS
    # ==========================================
    # Agrupamos por la clave del medicamento y sumamos sus cantidades entregadas
    top_claves_entregadas = ordenes.filter(cantidad_entregada__gt=0).values(
        'clave_contrato__medicamento__clave_sector'
    ).annotate(
        total_entregado=Sum('cantidad_entregada')
    ).order_by('-total_entregado')[:10] # Top 10

    # Top 5 Unidades/Hospitales con más pedidos (Para la gráfica)
    top_unidades = ordenes.values('nombre_unidad').annotate(
        total=Count('id')
    ).order_by('-total')[:5]

    nombres_unidades = [u['nombre_unidad'] or 'Sin Asignar' for u in top_unidades]
    cantidades_unidades = [u['total'] for u in top_unidades]

    # Órdenes Críticas (Atrasadas con multas creciendo)
    ordenes_criticas = [o for o in ordenes if o.dias_atraso > 0 and o.estatus in ['PENDIENTE', 'PARCIAL']]
    ordenes_criticas.sort(key=lambda x: x.penalizacion_estimada, reverse=True)
    ordenes_criticas = ordenes_criticas[:50] 

    context = {
        'busqueda': busqueda,
        # KPIs Originales
        'total_ordenes': total_ordenes,
        'entregadas': entregadas,
        'pendientes': pendientes,
        'atrasadas': atrasadas,
        'monto_total_str': f"${monto_total_solicitado:,.2f}",
        'penalizaciones_str': f"${penalizaciones_totales:,.2f}",
        'nombres_unidades_json': json.dumps(nombres_unidades),
        'cantidades_unidades_json': json.dumps(cantidades_unidades),
        'ordenes_criticas': ordenes_criticas,
        
        # 🔥 Nuevas variables inyectadas al HTML
        'piezas_solicitadas': total_piezas_solicitadas,
        'piezas_entregadas': total_piezas_entregadas,
        'piezas_pendientes': total_piezas_pendientes,
        'piezas_canceladas': total_piezas_canceladas,
        'top_claves_entregadas': top_claves_entregadas,
    }
    
    return render(request, 'dashboard_ordenes.html', context)

from django.shortcuts import render
from django.db.models import Sum, F, FloatField
from django.utils import timezone
from .models import OrdenCompra, PartidaCompra, SocioComercial

def dashboard_compras(request):
    # 1. FILTROS
    fecha_inicio = request.GET.get('fecha_inicio', '')
    fecha_fin = request.GET.get('fecha_fin', '')
    proveedor_id = request.GET.get('proveedor', '')
    
    ordenes = OrdenCompra.objects.all()
    partidas = PartidaCompra.objects.exclude(orden__estatus='CANCELADA')

    if fecha_inicio:
        ordenes = ordenes.filter(fecha_emision__gte=fecha_inicio)
        partidas = partidas.filter(orden__fecha_emision__gte=fecha_inicio)
    if fecha_fin:
        ordenes = ordenes.filter(fecha_emision__lte=fecha_fin)
        partidas = partidas.filter(orden__fecha_emision__lte=fecha_fin)
    if proveedor_id:
        ordenes = ordenes.filter(proveedor_id=proveedor_id)
        partidas = partidas.filter(orden__proveedor_id=proveedor_id)

    # 2. KPIs REALES BÁSICOS
    completas = ordenes.filter(estatus='RECIBIDA').count()
    total_ordenes = ordenes.count()
    
    stats_piezas = partidas.aggregate(
        total_pedidas=Sum('cantidad'),
        total_recibidas=Sum('cantidad_recibida')
    )
    piezas_pedidas = stats_piezas['total_pedidas'] or 0
    piezas_recibidas = stats_piezas['total_recibidas'] or 0
    
    # 3. CÁLCULOS PARA EL SEMÁFORO (Adaptados a tu imagen)
    porcentaje_otd = 0 # On Time Delivery (Cumplimiento)
    if piezas_pedidas > 0:
        porcentaje_otd = round((piezas_recibidas / piezas_pedidas) * 100, 1)
        
    tasa_desabasto = round(100 - porcentaje_otd, 1) if porcentaje_otd > 0 else 0.0

    # 4. ALERTAS DEL SISTEMA (Reales)
    alertas_criticas = []
    hoy = timezone.now().date()
    
    # Buscar OC atrasadas
    oc_atrasadas = ordenes.filter(estatus__in=['BORRADOR', 'AUTORIZADA', 'TRANSITO'], fecha_entrega_esperada__lt=hoy)
    for oc in oc_atrasadas:
        dias_retraso = (hoy - oc.fecha_entrega_esperada).days
        alertas_criticas.append({
            'tipo': 'oc_atrasada',
            'mensaje': f"{oc.folio} presenta un atraso de {dias_retraso} días.",
            'color': '#c0392b' # Rojo
        })

    # 5. TABLA DE ESTADO ACTUAL (Últimas 6 órdenes)
    ultimas_ordenes = ordenes.order_by('-fecha_emision')[:6]
    tabla_ordenes = []
    for oc in ultimas_ordenes:
        primera_partida = oc.partidas_compra.first()
        descripcion = "Varias partidas"
        if primera_partida:
            descripcion = f"{primera_partida.medicamento.denominacion_generica} x {primera_partida.cantidad}"
            if oc.partidas_compra.count() > 1:
                descripcion += " (+)"
                
        dias_transcurridos = (hoy - oc.fecha_emision).days if oc.fecha_emision else 0
        
        tabla_ordenes.append({
            'folio': oc.folio,
            'descripcion': descripcion,
            'proveedor': oc.proveedor.nombre[:20] + '...' if len(oc.proveedor.nombre) > 20 else oc.proveedor.nombre,
            'monto': oc.total_compra,
            'estatus': oc.get_estatus_display(),
            'estatus_raw': oc.estatus,
            'dias': dias_transcurridos
        })

    proveedores_con_ordenes = SocioComercial.objects.filter(ordencompra__isnull=False).distinct()

    # 6. EMPAQUETAMOS SOLO LO QUE EL NUEVO HTML NECESITA
    context = {
        'porcentaje_otd': porcentaje_otd,
        'tasa_desabasto': tasa_desabasto,
        'num_alertas': len(alertas_criticas),
        'alertas': alertas_criticas,
        'tabla_ordenes': tabla_ordenes,
        
        'proveedores': proveedores_con_ordenes,
        'filtros': {
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin,
            'proveedor': int(proveedor_id) if proveedor_id else ''
        }
    }
    
    return render(request, 'dashboard_compras.html', context)

def dashboard_inventario(request):
    import datetime
    from django.db.models import Sum
    from .models import Inventario
    
    hoy = datetime.date.today()
    limite_caducidad = hoy + datetime.timedelta(days=180) # Alerta: 6 meses
    
    # Solo tomamos lo que realmente tiene existencias
    inventario_activo = Inventario.objects.filter(cantidad_disponible__gt=0)
    
    # 1. KPIs Básicos
    total_piezas = inventario_activo.aggregate(total=Sum('cantidad_disponible'))['total'] or 0
    total_lotes = inventario_activo.count()
    
    # 2. Análisis de Caducidades
    lotes_riesgo = inventario_activo.filter(fecha_caducidad__lte=limite_caducidad).order_by('fecha_caducidad')
    alertas_caducidad = lotes_riesgo.count()
    
    # Lotes caducados (ya pasaron la fecha de hoy)
    lotes_caducados = inventario_activo.filter(fecha_caducidad__lt=hoy).count()

    # 3. Alertas para el panel derecho
    alertas_sistema = []
    if lotes_caducados > 0:
        alertas_sistema.append({
            'mensaje': f"CRÍTICO: Tienes {lotes_caducados} lotes CADUCADOS en el almacén.",
            'color': 'red'
        })
    if alertas_caducidad > 0:
        alertas_sistema.append({
            'mensaje': f"ATENCIÓN: {alertas_caducidad} lotes caducan en menos de 6 meses.",
            'color': 'orange'
        })
        
    # 4. Tabla de Estado Actual (Mostramos los más recientes o más críticos)
    tabla_inventario = inventario_activo.order_by('fecha_caducidad')[:10] # Ordenamos por caducidad para sacar lo más viejo primero

    context = {
        'total_piezas': total_piezas,
        'total_lotes': total_lotes,
        'alertas_caducidad': alertas_caducidad,
        'lotes_caducados': lotes_caducados,
        'alertas_sistema': alertas_sistema,
        'tabla_inventario': tabla_inventario,
        'hoy': hoy
    }
    
    return render(request, 'dashboard_inventario.html', context)

@staff_member_required
def dashboard_inicio(request):
    # Instanciamos tu código
    svc = DashboardService("GPHARMA")
    
    # Extraemos los datos (por ejemplo, del Año a la Fecha 'YTD')
    datos_ejecutivos = svc.obtener_datos_dashboard("YTD")
    
    context = {
        'title': 'Inicio',
        'datos': datos_ejecutivos, # Mandamos todo tu paquete al HTML
    }
    return render(request, 'dashboard_inicio.html', context)

from django.shortcuts import render
from django.contrib.admin.views.decorators import staff_member_required
from .models import Inventario, MovimientoKardex

@staff_member_required
def buscar_kardex(request):
    query = request.GET.get('q', '').strip()
    movimientos = []
    inventario_actual = None
    mensaje_error = None

    if query:
        # Usamos Q para buscar en código de barras O lote en una sola consulta
        # icontains ayuda por si escriben el lote en minúsculas o incompleto
        inventario_actual = Inventario.objects.filter(
            Q(codigo_barras=query) | Q(lote__icontains=query)
        ).first()

        if inventario_actual:
            # Si lo encontramos, traemos su historial del Kardex
            movimientos = MovimientoKardex.objects.filter(
                medicamento=inventario_actual.medicamento,
                lote=inventario_actual.lote
            ).order_by('-fecha')
        else:
            mensaje_error = f"No hay existencias ni registros para: {query}"

    context = {
        'query': query,
        'inventario_actual': inventario_actual,
        'movimientos': movimientos,
        'mensaje_error': mensaje_error,
    }
    return render(request, 'buscar_kardex.html', context)