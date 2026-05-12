import json
import csv
from datetime import timedelta
import datetime
from django.http import HttpResponse
from django.shortcuts import render
from django.db.models import Sum, F, Q, Count, Avg
from django.db.models.functions import TruncMonth, Coalesce
from django.utils import timezone
from django.contrib.admin.views.decorators import staff_member_required
from .models import Inventario, MovimientoKardex
from .services import DashboardService

from .models import (
    Contrato, Licitacion, PartidaRequerimiento, Empresa, 
    OrdenSuministro, RemisionEntrega, ClaveContrato, CatalogoMedicamento,
    PartidaOrden, Cotizacion, PartidaCotizacion, PedidoDirecto,
    OrdenCompra, PartidaCompra, SocioComercial, EntradaAlmacen,
    TraspasoIntercompany, DEPENDENCIAS_MAESTRAS
)

# ==========================================
# 1. DASHBOARD DE CONTRATOS (COMERCIAL VS LOGÍSTICO)
# ==========================================
@staff_member_required
def dashboard_contratos(request):
    contratos = Contrato.objects.all()
    
    busqueda = request.GET.get('q', '').strip() 
    filtro_dependencia = request.GET.get('dependencia', '')
    filtro_empresa_id = request.GET.get('empresa', '')
    filtro_contrato = request.GET.get('contrato', '')

    if busqueda:
        contratos = contratos.filter(
            Q(numero_contrato__icontains=busqueda) |
            Q(dependencia__icontains=busqueda) |
            Q(licitacion_origen__num_procedimiento__icontains=busqueda)
        ).distinct()

    if filtro_dependencia:
        contratos = contratos.filter(dependencia=filtro_dependencia)
    if filtro_empresa_id.isdigit():
        contratos = contratos.filter(empresa_id=filtro_empresa_id)
    if filtro_contrato:
        contratos = contratos.filter(numero_contrato=filtro_contrato)

    claves_qs = ClaveContrato.objects.filter(contrato__in=contratos)

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

    agg_sol_nuevo = PartidaOrden.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
        pzas=Sum('cantidad_solicitada'), din=Sum(F('cantidad_solicitada') * F('clave_contrato__precio_neto'))
    )
    agg_sol_hist = ClaveContrato.objects.filter(contrato__in=contratos).aggregate(
        pzas=Sum('piezas_historicas_solicitadas'), din=Sum(F('piezas_historicas_solicitadas') * F('precio_neto'))
    )
    piezas_solicitadas = (agg_sol_nuevo['pzas'] or 0) + (agg_sol_hist['pzas'] or 0)
    monto_solicitado = (agg_sol_nuevo['din'] or 0) + (agg_sol_hist['din'] or 0)

    agg_ent_nuevo = PartidaOrden.objects.filter(clave_contrato__contrato__in=contratos).aggregate(
        pzas=Sum('cantidad_entregada'), din=Sum(F('cantidad_entregada') * F('clave_contrato__precio_neto'))
    )
    agg_ent_hist = ClaveContrato.objects.filter(contrato__in=contratos).aggregate(
        pzas=Sum('piezas_historicas_entregadas'), din=Sum(F('piezas_historicas_entregadas') * F('precio_neto'))
    )
    piezas_entregadas = (agg_ent_nuevo['pzas'] or 0) + (agg_ent_hist['pzas'] or 0)
    monto_entregado = (agg_ent_nuevo['din'] or 0) + (agg_ent_hist['din'] or 0)

    piezas_pendientes_solicitar = piezas_maximas - piezas_solicitadas
    if piezas_pendientes_solicitar < 0: piezas_pendientes_solicitar = 0

    piezas_pendientes_entregar = piezas_solicitadas - piezas_entregadas
    if piezas_pendientes_entregar < 0: piezas_pendientes_entregar = 0

    avance_min_pct = (piezas_solicitadas / piezas_minimas * 100) if piezas_minimas > 0 else 0
    avance_max_pct = (piezas_solicitadas / piezas_maximas * 100) if piezas_maximas > 0 else 0

    ordenes_vinculadas = OrdenSuministro.objects.filter(partidas__clave_contrato__contrato__in=contratos).distinct()
    total_penalizado = sum(float(o.penalizacion_estimada) for o in ordenes_vinculadas)

    top_claves = claves_qs.annotate(
        importe_max=F('cantidad_maxima') * F('precio_neto')
    ).values('medicamento__clave_sector').annotate(
        total_asignado=Sum('importe_max')
    ).order_by('-total_asignado')[:5]
    
    nombres_top = [c['medicamento__clave_sector'] for c in top_claves]
    montos_top = [float(c['total_asignado']) for c in top_claves]

    detalle_claves_qs = claves_qs.select_related('medicamento', 'contrato').annotate(
        importe_max=F('cantidad_maxima') * F('precio_neto')
    ).order_by('-importe_max')[:100]
    
    detalle_claves = []
    for clave in detalle_claves_qs:
        ent_nuevo = PartidaOrden.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_entregada'))['tot'] or 0
        clave.pzas_entregadas = ent_nuevo + clave.piezas_historicas_entregadas

        sol_nuevo = PartidaOrden.objects.filter(clave_contrato=clave).aggregate(tot=Sum('cantidad_solicitada'))['tot'] or 0
        clave.pzas_solicitadas = sol_nuevo + clave.piezas_historicas_solicitadas
        
        faltantes = clave.pzas_solicitadas - clave.pzas_entregadas
        clave.pzas_faltantes = faltantes if faltantes > 0 else 0

        detalle_claves.append(clave)

    codigos_presentes = Contrato.objects.values_list('dependencia', flat=True).distinct()
    dependencias_agrupadas = []
    for categoria, items in DEPENDENCIAS_MAESTRAS:
        if isinstance(items, (list, tuple)):
            hospitales_del_grupo = []
            for cod, nombre in items:
                if cod in codigos_presentes:
                    hospitales_del_grupo.append((cod, nombre))
            if hospitales_del_grupo:
                dependencias_agrupadas.append((categoria, hospitales_del_grupo))
        else:
            if categoria in codigos_presentes:
                dependencias_agrupadas.append(('OTROS', [(categoria, items)]))

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
        'dependencias_disponibles': dependencias_agrupadas,
        'empresas_disponibles': empresas_qs,
        'contratos_disponibles': contratos_list,
        'filtro_dependencia': filtro_dependencia,
        'filtro_empresa': int(filtro_empresa_id) if filtro_empresa_id.isdigit() else '',
        'filtro_contrato': filtro_contrato,
        'busqueda': busqueda, 
    }
    return render(request, 'dashboard_contratos.html', context)


# ==========================================
# 2. DASHBOARD OMNI-CANAL (LICITACIONES + COTIZACIONES)
# ==========================================
@staff_member_required
def dashboard_licitaciones(request):
    q = request.GET.get('q', '').strip()
    filtro_empresa = request.GET.get('empresa', '')

    licitaciones = Licitacion.objects.all()
    partidas_lic = PartidaRequerimiento.objects.select_related('licitacion', 'medicamento')

    cotizaciones = Cotizacion.objects.all()
    partidas_cot = PartidaCotizacion.objects.select_related('cotizacion', 'medicamento')

    if filtro_empresa:
        licitaciones = licitaciones.filter(empresa_id=filtro_empresa)
        partidas_lic = partidas_lic.filter(licitacion__empresa_id=filtro_empresa)

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

    total_licitaciones = licitaciones.count() + cotizaciones.count()
    en_proceso = licitaciones.filter(estatus__estado='EN_PROCESO').count() + cotizaciones.filter(estatus__in=['BORRADOR', 'ENVIADA']).count()
    adjudicadas = licitaciones.filter(estatus__estado='ADJUDICADO').count() + cotizaciones.filter(estatus='GANADA').count()
    perdidas = licitaciones.filter(estatus__estado='PERDIDO').count() + cotizaciones.filter(estatus__in=['PERDIDA', 'CANCELADA']).count()

    monto_total = 0
    monto_ganado = 0
    monto_perdido = 0
    claves_participadas = set()
    claves_ganadas = set()
    
    for p in partidas_lic:
        importe = float(p.cantidad_maxima or 0) * float(p.precio or 0)
        monto_total += importe
        if p.medicamento_id: claves_participadas.add(p.medicamento_id)
        
        if p.resultado == 'Asignada':
            monto_ganado += importe
            if p.medicamento_id: claves_ganadas.add(p.medicamento_id)
        elif p.resultado in ['Perdida por precio', 'Perdida técnicamente']:
            monto_perdido += importe

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

    total_ordenes = ordenes.count()
    entregadas = ordenes.filter(estatus='ENTREGADA').count()
    pendientes = ordenes.filter(estatus__in=['PENDIENTE', 'PARCIAL']).count()
    
    hoy = timezone.now().date()
    atrasadas = ordenes.filter(estatus__in=['PENDIENTE', 'PARCIAL'], fecha_limite__lt=hoy).count()

    monto_total_solicitado = sum(float(o.valor_total) for o in ordenes)
    penalizaciones_totales = sum(float(o.penalizacion_estimada) for o in ordenes)

    partidas_qs = PartidaOrden.objects.filter(orden__in=ordenes)

    total_piezas_solicitadas = partidas_qs.aggregate(total=Sum('cantidad_solicitada'))['total'] or 0
    total_piezas_entregadas = partidas_qs.aggregate(total=Sum('cantidad_entregada'))['total'] or 0
    
    total_piezas_pendientes = total_piezas_solicitadas - total_piezas_entregadas
    if total_piezas_pendientes < 0: total_piezas_pendientes = 0 
        
    total_piezas_canceladas = PartidaOrden.objects.filter(orden__estatus='CANCELADA', orden__in=ordenes).aggregate(total=Sum('cantidad_solicitada'))['total'] or 0

    top_claves_entregadas = partidas_qs.filter(cantidad_entregada__gt=0).values(
        'clave_contrato__medicamento__clave_sector'
    ).annotate(
        total_entregado=Sum('cantidad_entregada')
    ).order_by('-total_entregado')[:10]

    top_unidades = ordenes.values('nombre_unidad').annotate(
        total=Count('id')
    ).order_by('-total')[:5]

    nombres_unidades = [u['nombre_unidad'] or 'Sin Asignar' for u in top_unidades]
    cantidades_unidades = [u['total'] for u in top_unidades]

    ordenes_criticas = [o for o in ordenes if o.dias_atraso > 0 and o.estatus in ['PENDIENTE', 'PARCIAL']]
    ordenes_criticas.sort(key=lambda x: x.penalizacion_estimada, reverse=True)
    ordenes_criticas = ordenes_criticas[:50] 

    # ==========================================
    # 🚀 NUEVO: RADAR LOGÍSTICO (ÓRDENES SURTIBLES)
    # ==========================================
    pedidos_pendientes = ordenes.filter(estatus__in=['PENDIENTE', 'PARCIAL'])
    ordenes_surtibles = []

    for ped in pedidos_pendientes:
        es_surtible = False
        detalles_surtibles = []
        
        for p in ped.partidas.all():
            pendientes_item = p.cantidad_solicitada - p.cantidad_entregada
            if pendientes_item > 0 and p.medicamento:
                # Buscamos si en almacén hay stock de esta clave
                stock_real = Inventario.objects.filter(medicamento=p.medicamento).aggregate(tot=Sum('cantidad_disponible'))['tot'] or 0
                if stock_real > 0:
                    es_surtible = True
                    detalles_surtibles.append({
                        'clave': p.medicamento.clave_sector,
                        'faltan': pendientes_item,
                        'hay': stock_real
                    })
        
        if es_surtible:
            # Calculamos si con el stock actual se puede matar la orden completa o solo un cacho
            surtido_completo = all((det['hay'] >= det['faltan']) for det in detalles_surtibles)
            
            dias_restantes = (ped.fecha_limite - hoy).days if ped.fecha_limite else 99
            
            ordenes_surtibles.append({
                'id': ped.id,
                'folio': ped.numero_orden_suministro,
                'cliente': ped.nombre_unidad or ped.dependencia or ped.razon_social,
                'tipo': ped.get_tipo_documento_display(),
                'detalles': detalles_surtibles,
                'surtido_completo': surtido_completo,
                'dias_restantes': dias_restantes
            })
    
    # Ordenamos: Primero las más urgentes (menos días) que se puedan surtir completas
    ordenes_surtibles.sort(key=lambda x: (-x['surtido_completo'], x['dias_restantes']))

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
        'ordenes_surtibles': ordenes_surtibles, # 👈 SE LO MANDAMOS AL HTML
        'num_surtibles': len(ordenes_surtibles)
    }
    
    return render(request, 'dashboard_ordenes.html', context)


# ==========================================
# 🔥 4. DASHBOARD DE COMPRAS E INTELIGENCIA 
# ==========================================
@staff_member_required
def dashboard_compras(request):
    fecha_inicio = request.GET.get('fecha_inicio', '')
    fecha_fin = request.GET.get('fecha_fin', '')
    proveedor_id = request.GET.get('proveedor', '')
    
    ordenes = OrdenCompra.objects.all().prefetch_related('partidas_compra')

    if fecha_inicio:
        ordenes = ordenes.filter(fecha_emision__gte=fecha_inicio)
    if fecha_fin:
        ordenes = ordenes.filter(fecha_emision__lte=fecha_fin)
    if proveedor_id:
        ordenes = ordenes.filter(proveedor_id=proveedor_id)

    total_gasto = 0
    ahorro_ppv = 0
    maverick_spend = 0
    tiempo_ciclo_dias = 0
    ordenes_recibidas = 0
    otif_count = 0
    piezas_totales_recibidas = 0
    piezas_totales_rechazadas = 0
    costo_admin_total = 0

    for oc in ordenes:
        total_oc = float(oc.total_compra)
        total_gasto += total_oc
        
        if hasattr(oc, 'costo_administrativo'):
            costo_admin_total += float(oc.costo_administrativo)
        
        if hasattr(oc, 'es_compra_no_planeada') and oc.es_compra_no_planeada:
            maverick_spend += total_oc

        if oc.estatus == 'RECIBIDA' and hasattr(oc, 'fecha_recepcion_real') and oc.fecha_recepcion_real:
            ordenes_recibidas += 1
            
            if hasattr(oc, 'fecha_necesidad') and oc.fecha_necesidad:
                dias_ciclo = (oc.fecha_recepcion_real - oc.fecha_necesidad).days
                tiempo_ciclo_dias += max(dias_ciclo, 1)

            a_tiempo = oc.fecha_recepcion_real <= oc.fecha_entrega_esperada
            completa = all(p.cantidad_recibida >= p.cantidad for p in oc.partidas_compra.all())
            
            sin_rechazos = True
            for p in oc.partidas_compra.all():
                if hasattr(p, 'piezas_rechazadas') and p.piezas_rechazadas > 0:
                    sin_rechazos = False
                    break
            
            if a_tiempo and completa and sin_rechazos:
                otif_count += 1

        for p in oc.partidas_compra.all():
            piezas_totales_recibidas += (p.cantidad_recibida or 0)
            
            if hasattr(p, 'piezas_rechazadas'):
                piezas_totales_rechazadas += p.piezas_rechazadas
            
            if hasattr(p, 'precio_referencia') and p.precio_referencia > 0:
                ahorro = float(p.precio_referencia - p.precio_unitario) * p.cantidad
                ahorro_ppv += ahorro

    pct_maverick = (maverick_spend / total_gasto * 100) if total_gasto > 0 else 0
    avg_tiempo_ciclo = (tiempo_ciclo_dias / ordenes_recibidas) if ordenes_recibidas > 0 else 0
    
    calidad_proveedores = 100
    if piezas_totales_recibidas > 0:
        calidad_proveedores = ((piezas_totales_recibidas - piezas_totales_rechazadas) / piezas_totales_recibidas) * 100

    pct_otif = (otif_count / ordenes_recibidas * 100) if ordenes_recibidas > 0 else 0

    # 🚨 RADAR DE ALERTAS INTER-MÓDULOS
    alertas_criticas = []
    hoy = timezone.now().date()

    oc_atrasadas = ordenes.filter(estatus__in=['BORRADOR', 'AUTORIZADA', 'TRANSITO'], fecha_entrega_esperada__lt=hoy)
    for oc in oc_atrasadas:
        dias_retraso = (hoy - oc.fecha_entrega_esperada).days
        alertas_criticas.append({
            'tipo': 'compras', 'color': '#e74c3c',
            'mensaje': f"PROVEEDOR ATRASADO: La OC-{oc.folio} presenta un atraso de {dias_retraso} días."
        })

    pedidos_pendientes = OrdenSuministro.objects.filter(estatus__in=['PENDIENTE', 'PARCIAL'])
    alertas_stockout = 0
    for ped in pedidos_pendientes:
        for p in ped.partidas.all():
            pendientes = p.cantidad_solicitada - p.cantidad_entregada
            if pendientes > 0 and p.medicamento:
                stock_real = Inventario.objects.filter(medicamento=p.medicamento).aggregate(tot=Sum('cantidad_disponible'))['tot'] or 0
                if stock_real < pendientes:
                    alertas_stockout += 1
                    break
    if alertas_stockout > 0:
        alertas_criticas.append({
            'tipo': 'stockout', 'color': '#e67e22',
            'mensaje': f"FALTA DE STOCK: Hay {alertas_stockout} órdenes/pedidos detenidos porque no hay suficiente inventario para surtirlos."
        })

    ocs_discrepancia = 0
    for oc in ordenes.filter(estatus='RECIBIDA'):
        reclamado = sum(p.cantidad_recibida for p in oc.partidas_compra.all() if p.cantidad_recibida)
        recibido = sum(e.cantidad_recibida for e in oc.entradaalmacen_set.all() if e.cantidad_recibida)
        if reclamado != recibido:
            ocs_discrepancia += 1
    if ocs_discrepancia > 0:
        alertas_criticas.append({
            'tipo': 'discrepancia', 'color': '#d35400',
            'mensaje': f"AUDITORÍA CIEGA: Existen {ocs_discrepancia} Órdenes de Compra con discrepancia entre lo pagado y lo recibido físicamente."
        })

    limite_caducidad = hoy + datetime.timedelta(days=180)
    lotes_riesgo = Inventario.objects.filter(cantidad_disponible__gt=0, fecha_caducidad__lte=limite_caducidad).count()
    if lotes_riesgo > 0:
        alertas_criticas.append({
            'tipo': 'caducidad', 'color': '#f1c40f',
            'mensaje': f"RIESGO DE MERMA: Tienes {lotes_riesgo} lotes de medicamentos próximos a caducar en almacén."
        })

    traspasos_pendientes = TraspasoIntercompany.objects.filter(estatus='BORRADOR').count()
    if traspasos_pendientes > 0:
        alertas_criticas.append({
            'tipo': 'traspaso', 'color': '#3498db',
            'mensaje': f"LOGÍSTICA INTERNA: Tienes {traspasos_pendientes} traspasos entre almacenes pendientes de procesar."
        })

    fecha_reciente = hoy - datetime.timedelta(days=7)
    devoluciones = MovimientoKardex.objects.filter(tipo='ENTRADA_DEVOLUCION', fecha__gte=fecha_reciente).count()
    if devoluciones > 0:
        alertas_criticas.append({
            'tipo': 'devolucion', 'color': '#9b59b6',
            'mensaje': f"REINTEGROS: Se han devuelto {devoluciones} lotes de mercancía al almacén por rechazo de hospitales (últimos 7 días)."
        })

    # 👇 NUEVA ALERTA 6: REGISTROS SANITARIOS POR VENCER (90 DÍAS) 👇
    limite_registro = hoy + datetime.timedelta(days=90)
    registros_por_vencer = CatalogoMedicamento.objects.filter(
        fecha_vigencia__lte=limite_registro, 
        fecha_vigencia__gte=hoy
    ).count()

    if registros_por_vencer > 0:
        alertas_criticas.append({
            'tipo': 'sanitario', 'color': '#2c3e50', # Gris oscuro / Elegante pero serio
            'mensaje': f"REGISTROS SANITARIOS: Tienes {registros_por_vencer} claves cuyo registro sanitario vence en menos de 3 meses. Revisar en Módulo Claves."
        })
    # 👆 FIN DE LA NUEVA ALERTA 👆

    # 👇 NUEVA ALERTA 7: Mermas / Piezas dañadas reportadas por Almacén 👇
    # Solo mostramos los rechazos reportados en los últimos 30 días para no saturar
    fecha_merma = timezone.now() - datetime.timedelta(days=30)
    rechazos_almacen = EntradaAlmacen.objects.filter(piezas_rechazadas__gt=0, fecha_ingreso__gte=fecha_merma)

    for rechazo in rechazos_almacen:
        alertas_criticas.append({
            'tipo': 'merma', 'color': '#8e44ad', # Morado oscuro / Calidad
            'mensaje': f"DEFECTO DE CALIDAD: Almacén rechazó {rechazo.piezas_rechazadas} pzas de la clave {rechazo.medicamento.clave_sector} (Lote: {rechazo.lote}) por llegar en mal estado."
        })
    # 👆 FIN DE LA NUEVA ALERTA 7 👆

    ultimas_ordenes = ordenes.order_by('-fecha_emision')[:10]
    tabla_ordenes = []
    for oc in ultimas_ordenes:
        primera_partida = oc.partidas_compra.first()
        descripcion = "Varias partidas"
        if primera_partida:
            descripcion = f"{primera_partida.medicamento.denominacion_generica} x {primera_partida.cantidad}"
            if oc.partidas_compra.count() > 1: descripcion += " (+)"
        
        tabla_ordenes.append({
            'folio': oc.folio,
            'proveedor': oc.proveedor.nombre[:25],
            'descripcion': descripcion,
            'monto': oc.total_compra,
            'estatus': oc.get_estatus_display(),
            'dias': (hoy - oc.fecha_emision).days
        })


    # ===============================================
    # 🛒 EL PUENTE: VENTAS -> COMPRAS (Planificador)
    # ===============================================
    requerimientos = {}

    partidas_lic = PartidaRequerimiento.objects.filter(resultado='Asignada').select_related('medicamento', 'licitacion')
    for p in partidas_lic:
        if not p.medicamento: continue
        med_id = p.medicamento.id
        if med_id not in requerimientos:
            requerimientos[med_id] = {'medicamento': p.medicamento, 'requerido': 0, 'origen': []}
        requerimientos[med_id]['requerido'] += (p.cantidad_maxima or 0)
        requerimientos[med_id]['origen'].append(f"{p.licitacion.num_procedimiento}")

    partidas_cot = PartidaCotizacion.objects.filter(cotizacion__estatus='GANADA').select_related('medicamento', 'cotizacion')
    for p in partidas_cot:
        if not p.medicamento: continue
        med_id = p.medicamento.id
        if med_id not in requerimientos:
            requerimientos[med_id] = {'medicamento': p.medicamento, 'requerido': 0, 'origen': []}
        requerimientos[med_id]['requerido'] += (p.cantidad or 0)
        requerimientos[med_id]['origen'].append(f"{p.cotizacion.folio}")

    planificador_abasto = []
    for med_id, data in requerimientos.items():
        med = data['medicamento']
        stock_real = Inventario.objects.filter(medicamento=med).aggregate(tot=Sum('cantidad_disponible'))['tot'] or 0
        faltante = data['requerido'] - stock_real
        
        if faltante > 0:
            planificador_abasto.append({
                'clave': med.clave_sector,
                'descripcion': med.denominacion_generica[:45] + "..." if len(med.denominacion_generica) > 45 else med.denominacion_generica,
                'socio': med.socio_contacto.nombre if med.socio_contacto else 'S/A',
                'requerido': data['requerido'],
                'stock': stock_real,
                'faltante': faltante,
                'origenes': list(set(data['origen']))[:2] # Máximo 2 para no saturar la vista
            })
    
    # Ordenar por el que más falte
    planificador_abasto.sort(key=lambda x: x['faltante'], reverse=True)

    proveedores_con_ordenes = SocioComercial.objects.filter(ordencompra__isnull=False).distinct()

    context = {
        'total_gasto': f"${total_gasto:,.2f}",
        'ahorro_ppv': f"${ahorro_ppv:,.2f}",
        'avg_tiempo_ciclo': f"{avg_tiempo_ciclo:.1f} días",
        'calidad_proveedores': f"{calidad_proveedores:.1f}%",
        'pct_otif': f"{pct_otif:.1f}%",
        'pct_maverick': f"{pct_maverick:.1f}%",
        'costo_admin_total': f"${costo_admin_total:,.2f}",
        
        'num_alertas': len(alertas_criticas),
        'alertas': alertas_criticas,
        'tabla_ordenes': tabla_ordenes,
        'planificador_abasto': planificador_abasto, 
        'proveedores': proveedores_con_ordenes,
        'filtros': {
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin,
            'proveedor': int(proveedor_id) if proveedor_id else ''
        }
    }
    
    return render(request, 'dashboard_compras.html', context)


# ==========================================
# 5. DASHBOARD DE INVENTARIO
# ==========================================
@staff_member_required
def dashboard_inventario(request):
    import datetime
    from django.db.models import Sum
    from .models import Inventario
    
    hoy = datetime.date.today()
    limite_caducidad = hoy + datetime.timedelta(days=180) # Alerta: 6 meses
    
    inventario_activo = Inventario.objects.filter(cantidad_disponible__gt=0)
    
    total_piezas = inventario_activo.aggregate(total=Sum('cantidad_disponible'))['total'] or 0
    total_lotes = inventario_activo.count()
    
    lotes_riesgo = inventario_activo.filter(fecha_caducidad__lte=limite_caducidad).order_by('fecha_caducidad')
    alertas_caducidad = lotes_riesgo.count()
    
    lotes_caducados = inventario_activo.filter(fecha_caducidad__lt=hoy).count()

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
        
    tabla_inventario = inventario_activo.order_by('fecha_caducidad')[:10] 

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
    svc = DashboardService("GPHARMA")
    datos_ejecutivos = svc.obtener_datos_dashboard("YTD")
    
    context = {
        'title': 'Inicio',
        'datos': datos_ejecutivos, 
    }
    return render(request, 'dashboard_inicio.html', context)


@staff_member_required
def buscar_kardex(request):
    query = request.GET.get('q', '').strip()
    movimientos = []
    inventario_actual = None
    mensaje_error = None

    if query:
        inventario_actual = Inventario.objects.filter(
            Q(codigo_barras=query) | Q(lote__icontains=query)
        ).first()

        if inventario_actual:
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