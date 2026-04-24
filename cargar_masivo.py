import os
import django
import pandas as pd

# 1. Conectar al cerebro del ERP
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'erp_core.settings') 
django.setup()

from licitaciones.models import OrdenSuministro

def importar_ordenes():
    # 🔥 OJO: Pon aquí el nombre exacto de tu archivo
    archivo = 'ordenes erp PRUEBA 2.xlsx' 
    
    print(f"🚀 Iniciando lectura hiper-inteligente de: {archivo}...")
    
    try:
        # Leemos el Excel
        df = pd.read_excel(archivo)
        df = df.fillna('') 
        
        creadas = 0
        actualizadas = 0
        
        for index, row in df.iterrows():
            folio = str(row.get('numero_orden_suministro', '')).strip()
            if not folio:
                continue 
                
            # 🧠 MAGIA 1: Deducir la Institución
            nombre_uni = str(row.get('nombre_unidad', '')).upper()
            if 'SEGURO SOCIAL' in nombre_uni or 'IMSS' in nombre_uni:
                dep_calculada = 'IMSS_ORDINARIO'
            elif 'ISSSTE' in nombre_uni:
                dep_calculada = 'ISSSTE'
            elif 'BIENESTAR' in nombre_uni:
                dep_calculada = 'IMSS_BIENESTAR'
            else:
                dep_calculada = 'OTRA'

            # 🧠 MAGIA 2: Calcular el Estatus en base al historial
            # Convertimos a números seguros por si hay celdas vacías
            solicitadas = int(pd.to_numeric(row.get('cantidad_solicitada', 0), errors='coerce') or 0)
            entregadas = int(pd.to_numeric(row.get('cantidad_entregada', 0), errors='coerce') or 0)
            
            if entregadas >= solicitadas and solicitadas > 0:
                estatus_calc = 'ENTREGADA'
            elif entregadas > 0:
                estatus_calc = 'PARCIAL'
            else:
                estatus_calc = 'PENDIENTE'

            # Inyectar a la base de datos mapeando tus columnas viejas a los campos nuevos
            orden, creada = OrdenSuministro.objects.update_or_create(
                numero_orden_suministro=folio,
                defaults={
                    'razon_social': str(row.get('razon_social', '')),
                    'numero_contrato_historico': str(row.get('numero_contrato', '')),
                    'clave_medicamento_historico': str(row.get('clave_medicamento', '')),
                    'dependencia': dep_calculada,
                    'nombre_unidad': str(row.get('nombre_unidad', '')),
                    'cantidad_solicitada': solicitadas,
                    'cantidad_entregada': entregadas,
                    'precio_unitario': float(pd.to_numeric(row.get('precio_unitario', 0), errors='coerce') or 0),
                    'fecha_limite': row.get('fecha_limite'),
                    'estatus': estatus_calc
                }
            )
            
            if creada:
                creadas += 1
            else:
                actualizadas += 1
                
        print("\n" + "="*50)
        print("✅ ¡CARGA MASIVA COMPLETADA CON ÉXITO!")
        print(f"📦 Órdenes Nuevas Creadas: {creadas}")
        print(f"🔄 Órdenes Actualizadas: {actualizadas}")
        print("="*50 + "\n")
        
    except Exception as e:
        print(f"🚨 ERROR FATAL: {str(e)}")

if __name__ == '__main__':
    importar_ordenes()