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

# Traemos todos los modelos (¡Incluidos PartidaOrden, Cotizacion y DEPENDENCIAS_MAESTRAS!)
from .models import (
    Contrato, Licitacion, PartidaRequerimiento, Empresa, 
    OrdenSuministro, RemisionEntrega, ClaveContrato, CatalogoMedicamento,
    PartidaOrden, Cotizacion, PartidaCotizacion, PedidoDirecto,
    DEPENDENCIAS_MAESTRAS
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
    # 4. CALCULAMOS EL AVANCE COMERCIAL Y LOGÍSTICO (CARRITO NUEVO)
    # =========================================================
    
    # 4.1 Piezas y Monto Solicitado (Leemos desde PartidaOrden)
    agg_solicitadas = PartidaOrden.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
        tot_pzas=Sum('cantidad_solicitada'),
        tot_dinero=Sum(F('cantidad_solicitada') * F('clave_contrato__precio_neto'))
    )
    piezas_solicitadas = agg_solicitadas.get('tot_pzas') or 0
    monto_solicitado = agg_solicitadas.get('tot_dinero') or 0

    # 4.2 Piezas y Monto Entregado (Leemos desde PartidaOrden)
    agg_entregadas = PartidaOrden.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
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
    ordenes_vinculadas = OrdenSuministro.objects.filter(partidas__clave_contrato__contrato__in=contratos).distinct()
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
        entregas_dict = PartidaOrden.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_entregada'))
        clave.pzas_entregadas = entregas_dict.get('tot') or 0

        solicitadas_dict = PartidaOrden.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_solicitada'))
        clave.pzas_solicitadas = solicitadas_dict.get('tot') or 0
        
        faltantes = clave.pzas_solicitadas - clave.pzas_entregadas
        clave.pzas_faltantes = faltantes if faltantes > 0 else 0

        detalle_claves.append(clave)

    # 7. FILTROS EN CASCADA (MOSTRANDO NOMBRES REALES)
    codigos_presentes = Contrato.objects.values_list('dependencia', flat=True).distinct()
    
    # Aplanamos la lista de dependencias para buscar los nombres bonitos
    deps_flat = {}
    for cat, items in DEPENDENCIAS_MAESTRAS:
        if isinstance(items, (list, tuple)):
            for code, name in items:
                deps_flat[code] = name
        else:
            deps_flat[cat] = items

    # Creamos la lista para el dropdown del HTML
    dependencias_formateadas = []
    for cod in codigos_presentes:
        if cod:  # Ignorar vacíos
            dependencias_formateadas.append({
                'id': cod,
                'nombre': deps_flat.get(cod, cod) # Si no halla el código, deja el código
            })
    
    empresas_qs = Empresa.objects.filter(contrato__isnull=False).distinct()
    if filtro_dependencia:
        empresas_qs = empresas_qs.filter(contrato__dependencia=filtro_dependencia)
        
    contratos_qs = Contrato.objects.all()
    if filtro_dependencia: contratos_qs = contratos_qs.filter(dependencia=filtro_dependencia)
    if filtro_empresa_id.isdigit(): contratos_qs = contratos_qs.filter(empresa_id=filtro_empresa_id)
    contratos_list = contratos_qs.values_list('numero_contrato', flat=True).distinct()

    context = {
        'monto_minimo_str': f"${monto_minimo:,.2f}",
        'monto_maximo_str': f"${monto_maximo:,.2f}",
        'monto_solicitado_str': f"${monto_solicitado:,.2f}",
        'monto_entregado_str': f"${monto_entregado:,.2f}",
        'piezas_minimas_str': f"{piezas_minimas:,}",
        'piezas_maximas_str': f"{piezas_maximas:,}",
        'total_penalizado_str': f"-${total_penalizado:,.2f}",
        'piezas_solicitadas': piezas_solicitadas,
        'piezas_pendientes_solicitar': piezas_pendientes_solicitar,
        'piezas_entregadas': piezas_entregadas,
        'piezas_pendientes_entregar': piezas_pendientes_entregar,
        'avance_min_pct': f"{avance_min_pct:,.1f}%",
        'avance_max_pct': f"{avance_max_pct:,.1f}%",
        'nombres_top_json': json.dumps(nombres_top),
        'montos_top_json': json.dumps(montos_top),
        'detalle_claves': detalle_claves,
        
        # Filtro de dependencias inteligente inyectado aquí 👇
        'dependencias_disponibles': dependencias_formateadas,
        
        'empresas_disponibles': empresas_qs,
        'contratos_disponibles': contratos_list,
        'filtro_dependencia': filtro_dependencia,
        'filtro_empresa': int(filtro_empresa_id) if filtro_empresa_id.isdigit() else '',
        'filtro_contrato': filtro_contrato,
        'busqueda': busqueda, 
    }
    return render(request, 'dashboard_contratos.html', context)


# ==========================================
# 🔥 2. DASHBOARD OMNI-CANAL (LICITACIONES + COTIZACIONES)
# ==========================================
@staff_member_required
def dashboard_licitaciones(request):
    q = request.GET.get('q', '').strip()
    filtro_empresa = request.GET.get('empresa', '')

    # 1. TRAEMOS AMBOS MUNDOS
    licitaciones = Licitacion.objects.all()
    partidas_lic = PartidaRequerimiento.objects.select_related('licitacion', 'medicamento')

    cotizaciones = Cotizacion.objects.all()
    partidas_cot = PartidaCotizacion.objects.select_related('cotizacion', 'medicamento')

    # 2. FILTRO POR EMPRESA
    if filtro_empresa:
        licitaciones = licitaciones.filter(empresa_id=filtro_empresa)
        partidas_lic = partidas_lic.filter(licitacion__empresa_id=filtro_empresa)
        # Cotización no tiene empresa atada por ahora, por lo que es global.

    # 3. BÚSQUEDA GLOBAL SIMULTÁNEA EN AMBAS TABLAS
    if q:
        licitaciones = licitaciones.filter(Q(num_procedimiento__icontains=q) | Q(dependencia__icontains=q))
        partidas_lic = partidas_lic.filter(
            Q(licitacion__num_procedimiento__icontains=q) |
            Q(medicamento__clave_sector__icontains=q) |
            Q(medicamento__fabricante__icontains=q) |
            Q(resultado__icontains=q)
        )

        cotizaciones = cotizaciones.filter(Q(folio__icontains=q) | Q(razon_social__icontains=q) | Q(dependencia__icontains=q))
        partidas_cot = partidas_cot.filter(
            Q(cotizacion__folio__icontains=q) |
            Q(medicamento__clave_sector__icontains=q) |
            Q(medicamento__fabricante__icontains=q)
        )

    # 4. SUMAMOS TOTALES Y ESTATUS
    total_licitaciones = licitaciones.count() + cotizaciones.count()
    en_proceso = licitaciones.filter(estatus__estado='EN_PROCESO').count() + cotizaciones.filter(estatus__in=['BORRADOR', 'ENVIADA']).count()
    adjudicadas = licitaciones.filter(estatus__estado='ADJUDICADO').count() + cotizaciones.filter(estatus='GANADA').count()
    perdidas = licitaciones.filter(estatus__estado='PERDIDO').count() + cotizaciones.filter(estatus__in=['PERDIDA', 'CANCELADA']).count()

    monto_total = 0
    monto_ganado = 0
    monto_perdido = 0
    claves_participadas = set()
    claves_ganadas = set()
    
    # Procesamos Licitaciones
    for p in partidas_lic:
        importe = float(p.cantidad_maxima or 0) * float(p.precio or 0)
        monto_total += importe
        if p.medicamento_id: claves_participadas.add(p.medicamento_id)
        
        if p.resultado == 'Asignada':
            monto_ganado += importe
            if p.medicamento_id: claves_ganadas.add(p.medicamento_id)
        elif p.resultado in ['Perdida por precio', 'Perdida técnicamente']:
            monto_perdido += importe

    # Procesamos Ventas Directas
    for p in partidas_cot:
        importe = float(p.cantidad or 0) * float(p.precio_unitario or 0)
        monto_total += importe
        if p.medicamento_id: claves_participadas.add(p.medicamento_id)
        
        if p.cotizacion.estatus == 'GANADA':
            monto_ganado += importe
            if p.medicamento_id: claves_ganadas.add(p.medicamento_id)
        elif p.cotizacion.estatus in ['PERDIDA', 'CANCELADA']:
            monto_perdido += importe

    monto_en_proceso = monto_total - monto_ganado - monto_perdido
    if monto_en_proceso < 0: monto_en_proceso = 0

    # 5. TOP 5 CLAVES (Fusionadas)
    claves_dict = {}
    for p in partidas_lic.filter(resultado='Asignada'):
        k = (p.medicamento.clave_sector, p.medicamento.fabricante)
        claves_dict[k] = claves_dict.get(k, 0) + (float(p.cantidad_maxima or 0) * float(p.precio or 0))

    for p in partidas_cot.filter(cotizacion__estatus='GANADA'):
        k = (p.medicamento.clave_sector, p.medicamento.fabricante)
        claves_dict[k] = claves_dict.get(k, 0) + (float(p.cantidad or 0) * float(p.precio_unitario or 0))

    top_claves_sorted = sorted(claves_dict.items(), key=lambda x: x[1], reverse=True)[:5]
    top_claves = []
    nombres_top = []
    montos_top = []

    for (clave, fab), importe in top_claves_sorted:
        top_claves.append({
            'medicamento__clave_sector': clave,
            'medicamento__fabricante': fab,
            'total_importe': importe
        })
        nombres_top.append(clave)
        montos_top.append(importe)

    # 6. TABLAS INFERIORES DE DETALLE (Objeto Simulado para no romper el HTML)
    class DummyObj: pass
    detalle_ganadas = []
    detalle_perdidas = []

    for p in partidas_lic.filter(resultado__in=['Asignada', 'Perdida por precio', 'Perdida técnicamente']):
        item = DummyObj()
        item.licitacion = DummyObj()
        item.licitacion.num_procedimiento = p.licitacion.num_procedimiento
        item.licitacion.fecha_fallo = p.licitacion.fecha_fallo
        item.medicamento = DummyObj()
        item.medicamento.clave_sector = p.medicamento.clave_sector
        item.medicamento.fabricante = p.medicamento.fabricante
        item.resultado = p.resultado
        item.importe = float(p.cantidad_maxima or 0) * float(p.precio or 0)
        item.importe_ganado = item.importe

        if p.resultado == 'Asignada':
            detalle_ganadas.append(item)
        else:
            detalle_perdidas.append(item)

    for p in partidas_cot.filter(cotizacion__estatus__in=['GANADA', 'PERDIDA', 'CANCELADA']):
        item = DummyObj()
        item.licitacion = DummyObj()
        item.licitacion.num_procedimiento = p.cotizacion.folio
        item.licitacion.fecha_fallo = p.cotizacion.fecha_emision
        item.medicamento = DummyObj()
        item.medicamento.clave_sector = p.medicamento.clave_sector
        item.medicamento.fabricante = p.medicamento.fabricante
        item.resultado = p.cotizacion.get_estatus_display()
        item.importe = float(p.cantidad or 0) * float(p.precio_unitario or 0)
        item.importe_ganado = item.importe

        if p.cotizacion.estatus == 'GANADA':
            detalle_ganadas.append(item)
        else:
            detalle_perdidas.append(item)

    detalle_ganadas.sort(key=lambda x: x.importe, reverse=True)
    detalle_ganadas = detalle_ganadas[:50]
    detalle_perdidas.sort(key=lambda x: x.importe, reverse=True)
    detalle_perdidas = detalle_perdidas[:50]

    proximos_fallos = licitaciones.filter(
        estatus__estado='EN_PROCESO', fecha_fallo__gte=timezone.now()
    ).order_by('fecha_fallo')[:5]

    empresas_disponibles = Empresa.objects.all()

    context = {
        'title': 'Panel Ejecutivo Omnicanal',
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

@staff_member_required
def dashboard_ordenes(request):
    busqueda = request.GET.get('q', '').strip()
    
    ordenes = OrdenSuministro.objects.all()
    
    if busqueda:
        ordenes = ordenes.filter(
            Q(numero_orden_suministro__icontains=busqueda) |
            Q(nombre_unidad__icontains=busqueda) |
            Q(partidas__clave_contrato__medicamento__clave_sector__icontains=busqueda) |
            Q(partidas__medicamento__clave_sector__icontains=busqueda) |
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
    monto_total_solicitado = sum(float(o.valor_total) for o in ordenes)
    penalizaciones_totales = sum(float(o.penalizacion_estimada) for o in ordenes)

    # ==========================================
    # 🔥 NUEVOS CÁLCULOS: INTELIGENCIA DE PIEZAS (Desde PartidaOrden)
    # ==========================================
    partidas_qs = PartidaOrden.objects.filter(orden__in=ordenes)

    # 1. Total de Piezas Solicitadas
    total_piezas_solicitadas = partidas_qs.aggregate(total=Sum('cantidad_solicitada'))['total'] or 0
    
    # 2. Total de Piezas Entregadas
    total_piezas_entregadas = partidas_qs.aggregate(total=Sum('cantidad_entregada'))['total'] or 0
    
    # 3. Total de Piezas Pendientes (Calculado)
    total_piezas_pendientes = total_piezas_solicitadas - total_piezas_entregadas
    if total_piezas_pendientes < 0:
        total_piezas_pendientes = 0 # Protección por si hay entregas de más
        
    # 4. Total de Piezas Canceladas 
    total_piezas_canceladas = PartidaOrden.objects.filter(orden__estatus='CANCELADA', orden__in=ordenes).aggregate(total=Sum('cantidad_solicitada'))['total'] or 0

    # ==========================================
    # 🔥 NUEVO TOP: CLAVES MÁS ENTREGADAS
    # ==========================================
    top_claves_entregadas = partidas_qs.filter(cantidad_entregada__gt=0).values(
        'clave_contrato__medicamento__clave_sector'
    ).annotate(
        total_entregado=Sum('cantidad_entregada')
    ).order_by('-total_entregado')[:10]

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
        'total_ordenes': total_ordenes,
        'entregadas': entregadas,
        'pendientes': pendientes,
        'atrasadas': atrasadas,
        'monto_total_str': f"${monto_total_solicitado:,.2f}",
        'penalizaciones_str': f"${penalizaciones_totales:,.2f}",
        'nombres_unidades_json': json.dumps(nombres_unidades),
        'cantidades_unidades_json': json.dumps(cantidades_unidades),
        'ordenes_criticas': ordenes_criticas,
        'piezas_solicitadas': total_piezas_solicitadas,
        'piezas_entregadas': total_piezas_entregadas,
        'piezas_pendientes': total_piezas_pendientes,
        'piezas_canceladas': total_piezas_canceladas,
        'top_claves_entregadas': top_claves_entregadas,
    }
    
    return render(request, 'dashboard_ordenes.html', context)

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