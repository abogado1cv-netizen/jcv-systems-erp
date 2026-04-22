from django.contrib import admin
from django.urls import path
from django.views.generic import RedirectView 
from django.conf import settings
from django.conf.urls.static import static

# Importamos tus vistas para el tablero
from licitaciones import views

# ==============================================================
# 🚀 EL GOLPE DEFINITIVO A JAZZMIN (Monkey Patching)
# Reemplazamos el "cerebro" del inicio oficial del admin por el tuyo
# ==============================================================
def inicio_personalizado(request, extra_context=None):
    return views.dashboard_inicio(request)

admin.site.index = inicio_personalizado
# ==============================================================

urlpatterns = [
    # Si alguien entra a la raíz, lo mandamos al admin
    path('', RedirectView.as_view(url='/admin/', permanent=False)),
    
    # La ruta oficial del admin
    path('admin/', admin.site.urls),
    
    # --- LA LÍNEA DEL ERROR LA BORRAMOS ---
    
    # Rutas para tus dashboards independientes
    path('dashboard/inicio/', views.dashboard_inicio, name='dashboard_inicio'),
    path('dashboard/contratos/', views.dashboard_contratos, name='dashboard_contratos'),
    path('dashboard/licitaciones/', views.dashboard_licitaciones, name='dashboard_licitaciones'),
    path('dashboard/ordenes/', views.dashboard_ordenes, name='dashboard_ordenes'),
    path('dashboard/compras/', views.dashboard_compras, name='dashboard_compras'),
    path('dashboard/inventario/', views.dashboard_inventario, name='dashboard_inventario'),
]

# ==============================================================
# 📂 SERVIDOR DE ARCHIVOS (ESTÁTICOS Y MEDIA)
# ==============================================================
if settings.DEBUG:
    # Para CSS, JS e Imágenes de diseño
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
    # 🚀 LA MAGIA: Esto permite que el Admin abra tus PDFs de evidencias
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)