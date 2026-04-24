import os
import django
import pandas as pd

# 1. Conectar este script al cerebro de GPHARMA
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'erp_core.settings') # Cambia erp_core si tu carpeta principal se llama diferente
django.setup()

from licitaciones.models import OrdenSuministro

def importar_ordenes():
    # Nombre del archivo Excel que vamos a leer
    archivo_excel = 'ordenes_historicas.xlsx'
    
    print(f"🚀 Iniciando lectura del archivo: {archivo_excel}...")
    
    try:
        # Leemos el Excel y rellenamos los espacios vacíos
        df = pd.read_excel(archivo_excel)
        df = df.fillna('') 
        
        creadas = 0
        actualizadas = 0
        
        for index, row in df.iterrows():
            folio = str(row['numero_orden_suministro']).strip()
            if not folio:
                continue # Saltamos filas vacías
                
            # Esto busca si la orden ya existe para actualizarla, si no, la crea nueva. ¡Así no hay duplicados!
            orden, creada = OrdenSuministro.objects.update_or_create(
                numero_orden_suministro=folio,
                defaults={
                    'razon_social': str(row['razon_social']),
                    'numero_contrato_historico': str(row['numero_contrato_historico']),
                    'clave_medicamento_historico': str(row['clave_medicamento_historico']),
                    'dependencia': str(row['dependencia']),
                    'nombre_unidad': str(row['nombre_unidad']),
                    'cantidad_solicitada': int(row['cantidad_solicitada'] or 0),
                    'cantidad_entregada': int(row['cantidad_entregada'] or 0),
                    'precio_unitario': float(row['precio_unitario'] or 0.0),
                    'fecha_limite': row['fecha_limite'],
                    'estatus': str(row['estatus']).strip().upper() or 'PENDIENTE'
                }
            )
            
            if creada:
                creadas += 1
            else:
                actualizadas += 1
                
        print("\n" + "="*40)
        print("✅ ¡CARGA MASIVA COMPLETADA CON ÉXITO!")
        print(f"📦 Órdenes Nuevas Creadas: {creadas}")
        print(f"🔄 Órdenes Actualizadas: {actualizadas}")
        print("="*40 + "\n")
        
    except Exception as e:
        print(f"🚨 ERROR: {str(e)}")

if __name__ == '__main__':
    importar_ordenes()