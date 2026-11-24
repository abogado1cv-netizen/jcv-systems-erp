from django.http import HttpResponseForbidden
from django.shortcuts import redirect

MAPA_DASHBOARDS = {
    '/dashboard/licitaciones/': 'dash_comercial',
    '/dashboard/historial-precios/': 'dash_precios',
    '/dashboard/contratos/': 'dash_contratos',
    '/dashboard/compras/': 'dash_compras',
    '/dashboard/ordenes/': 'dash_logistica',
    '/dashboard/inventario/': 'dash_almacen',
}

class ControlAccesoMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user = getattr(request, 'user', None)

        if user and user.is_authenticated and not user.is_superuser:
            acceso = getattr(user, 'control_acceso', None)
            if acceso:
                ruta = request.path

                # 1. Si es "Solo Dashboards" y entra al /admin/ (excepto login/logout), lo mandamos al inicio
                if acceso.solo_dashboards and ruta.startswith('/admin/') and not ruta.startswith(('/admin/login/', '/admin/logout/')):
                    return redirect('/dashboard/inicio/')

                # 2. Validar acceso a cada dashboard específico
                for prefijo, campo in MAPA_DASHBOARDS.items():
                    if ruta.startswith(prefijo) and not getattr(acceso, campo, False):
                        return HttpResponseForbidden(
                            "<h2 style='font-family:sans-serif; text-align:center; margin-top:60px;'>"
                            "🔒 No tienes permiso para ver este Dashboard.</h2>"
                        )

        return self.get_response(request)