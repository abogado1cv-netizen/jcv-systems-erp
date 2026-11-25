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

HTML_SIN_ACCESO = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Acceso Restringido</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f3f4f6; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
        .card { background: #ffffff; padding: 40px 36px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); text-align: center; max-width: 420px; width: 90%; border: 1px solid #e5e7eb; }
        .icon { font-size: 48px; margin-bottom: 12px; }
        h2 { color: #111827; margin: 0 0 10px; font-size: 22px; }
        p { color: #4b5563; margin: 0 0 24px; font-size: 15px; line-height: 1.5; }
        a { display: inline-block; background: #2563eb; color: #ffffff; text-decoration: none; padding: 10px 22px; border-radius: 8px; font-weight: 500; font-size: 14px; }
        a:hover { background: #1d4ed8; }
    </style>
</head>
<body>
    <div class="card">
        <div class="icon">🔒</div>
        <h2>Acceso Restringido</h2>
        <p>No tienes acceso a esta sección.<br>Comunícate con el administrador.</p>
        <a href="javascript:history.back()">← Regresar</a>
    </div>
</body>
</html>
"""

class ControlAccesoMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user = getattr(request, 'user', None)

        if user and user.is_authenticated and not user.is_superuser:
            acceso = getattr(user, 'control_acceso', None)
            if acceso:
                ruta = request.path

                if acceso.solo_dashboards and ruta.startswith('/admin/') and not ruta.startswith(('/admin/login/', '/admin/logout/')):
                    return redirect('/dashboard/inicio/')

                for prefijo, campo in MAPA_DASHBOARDS.items():
                    if ruta.startswith(prefijo) and not getattr(acceso, campo, False):
                        return HttpResponseForbidden(HTML_SIN_ACCESO)

        response = self.get_response(request)

        # Reemplaza cualquier pantalla genérica de 403 Forbidden
        if response.status_code == 403:
            return HttpResponseForbidden(HTML_SIN_ACCESO)

        return response