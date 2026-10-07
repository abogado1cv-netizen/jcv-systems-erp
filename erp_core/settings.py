import os
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = 'django-insecure-aakcsii+s+*f$%&7t5lm2w50m@9q(1)%4&lw3ffvm@!2vm%)@y'
DEBUG = True
ALLOWED_HOSTS = ['*']

INSTALLED_APPS = [
    'jazzmin',
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django.contrib.humanize', # <--- Formato de dinero
    'licitaciones',
    'import_export',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'erp_core.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'erp_core.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',},
]

LANGUAGE_CODE = 'es-mx'
TIME_ZONE = 'America/Mexico_City'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_DIRS = [os.path.join(BASE_DIR, 'static')]

# ==========================================
# GESTIÓN DE ARCHIVOS SUBIDOS
# ==========================================
MEDIA_URL = 'http://67.205.175.17/media/'
MEDIA_ROOT = '/var/www/media/'

LOGIN_REDIRECT_URL = '/dashboard/inicio/'
LOGOUT_REDIRECT_URL = '/admin/login/'
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ==========================================
# SEGURIDAD Y CIERRE DE SESIÓN AUTOMÁTICO
# ==========================================
SESSION_COOKIE_AGE = 900 
SESSION_SAVE_EVERY_REQUEST = True 
SESSION_EXPIRE_AT_BROWSER_CLOSE = True

# ==========================================
# CONFIGURACIÓN DEL MENÚ GPHARMA (JAZZMIN PREMIUM AURA)
# ==========================================
JAZZMIN_SETTINGS = {
    "site_title": "GPHARMA ERP",
    "site_header": "Módulos",
    "site_brand": "GPHARMA",
    
    "site_logo": "logo.png",
    "login_logo": "logo.png",
    "custom_css": "css/aura_pro.css", 
    "site_logo_classes": "img-fluid",
    "login_logo_classes": "img-fluid", 
    
    "welcome_sign": "Bienvenido al Centro de Control GPHARMA",
    "copyright": "JCV Systems",
    
    # ==================================================
    # 🎯 1. MENÚ SUPERIOR
    # ==================================================
    "topmenu_links": [
        {"name": "Inicio",  "url": "admin:index", "permissions": ["auth.view_user"]},
        {"name": "📊 Comercial", "url": "/dashboard/licitaciones/"},
        {'name': '🔍 Historial de Precios', 'url': '/dashboard/historial-precios/'},
        {"name": "📁 Contratos", "url": "/dashboard/contratos/"},
        {"name": "🛒 Compras", "url": "/dashboard/compras/"},
        {"name": "🚚 Logística", "url": "/dashboard/ordenes/"},
        {"name": "📦 Almacén", "url": "/dashboard/inventario/"},
    ],

    "custom_links": {
        "licitaciones": [{
            "name": "Gestión de Equipo",
            "url": "/dashboard/equipo/",
            "icon": "fas fa-users",
        }]
    },

    "hide_models": [
        "licitaciones.estatusprocedimiento",
        "licitaciones.configuracionemail",
        "licitaciones.partidarequerimiento",
    ],
    "hide_apps": ["auth"], 

    # ==================================================
    # 🎨 3. ICONOS MINIMALISTAS (Look Apple / SF Symbols)
    # ==================================================
    "icons": {
        # Maestros
        "licitaciones.catalogomedicamento": "fas fa-capsules", 
        "licitaciones.empresa": "fas fa-building",           
        "licitaciones.sociocomercial": "fas fa-handshake", 
        "licitaciones.almacen": "fas fa-warehouse",
        
        # Comercial
        "licitaciones.licitacion": "fas fa-folder-open", # Carpeta limpia
        "licitaciones.cotizacion": "fas fa-file-invoice", # Documento ligero
        "licitaciones.contrato": "fas fa-file-signature",          
        
        # Compras
        "licitaciones.ordencompra": "fas fa-shopping-cart",
        "licitaciones.entradaalmacen": "fas fa-box-open", # Caja abierta limpia
        
        # Inventario
        "licitaciones.inventario": "fas fa-boxes",               
        "licitaciones.traspasointercompany": "fas fa-exchange-alt",
        "licitaciones.incidenciainventario": "fas fa-radiation-alt", 
        "licitaciones.escanerkardex": "fas fa-barcode",
        
        # Logística
        "licitaciones.ordensuministro": "fas fa-truck-loading",
        "licitaciones.pedidodirecto": "fas fa-paper-plane",
        "licitaciones.remisionentrega": "fas fa-receipt",
        
        # Extras y Auth
        "licitaciones.registroubicacion": "fas fa-map-marker-alt",
        "auth.user": "fas fa-user-circle",
        "auth.group": "fas fa-users-cog",
    },
    
    # ... tu "order_with_respect_to" se queda igual ...
    "order_with_respect_to": [
        "licitaciones.catalogomedicamento", "licitaciones.empresa", "licitaciones.sociocomercial", "licitaciones.almacen",
        "licitaciones.licitacion", "licitaciones.cotizacion", "licitaciones.contrato",
        "licitaciones.ordencompra", "licitaciones.entradaalmacen",
        "licitaciones.inventario", "licitaciones.traspasointercompany", "licitaciones.incidenciainventario", "licitaciones.escanerkardex", 
        "licitaciones.ordensuministro", "licitaciones.pedidodirecto", "licitaciones.remisionentrega",
        "licitaciones.registroubicacion",
    ],
    
    "default_icon_parents": "fas fa-folder",
    "default_icon_children": "fas fa-circle",
    "show_sidebar": True,
    "navigation_expanded": True,
    
    # 👇 ESTO ES CLAVE PARA ACTIVAR EL INTERRUPTOR DE MODO OSCURO EN LA PANTALLA 👇
    "show_ui_builder": False, 
}

# ==================================================
# 🌗 TWEAKS ADAPTATIVOS (SOPORTE LIGHT/DARK AUTOMÁTICO)
# ==================================================
JAZZMIN_UI_TWEAKS = {
    "navbar_small_text": False,
    "footer_small_text": False,
    "body_small_text": False,
    "brand_small_text": False,
    
    # Dejamos que los colores los dicte nuestro CSS
    "brand_colour": False, 
    "navbar": False,
    "no_navbar_border": True,
    "navbar_fixed": True,
    "layout_boxed": False,
    "footer_fixed": False,
    "sidebar_fixed": True,
    "sidebar": "sidebar-dark-primary", # El sidebar siempre se ve más pro oscuro
    "sidebar_nav_small_text": False,
    "sidebar_disable_expand": False,
    "sidebar_nav_child_indent": True,
    "sidebar_nav_compact_style": False,
    "sidebar_nav_legacy_style": False,
    "sidebar_nav_flat_style": True, 
    
    # Habilitamos temas neutros para que nuestra magia CSS brille
    "theme": "default", 
    "dark_mode_theme": "darkly", # 🌙 Activa el soporte nativo oscuro
    
    "button_classes": {
        "primary": "btn-primary",
        "secondary": "btn-secondary",
        "info": "btn-info",
        "warning": "btn-warning",
        "danger": "btn-danger",
        "success": "btn-success"
    }
}

# ==========================================
# CONFIGURACIÓN DE CORREO
# ==========================================
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.resend.com'
EMAIL_PORT = 587
EMAIL_HOST_USER = 'resend' 
EMAIL_HOST_PASSWORD = os.getenv('RESEND_API_KEY')
EMAIL_USE_SSL = False 
EMAIL_USE_TLS = True 
DEFAULT_FROM_EMAIL = 'notificaciones@jcv-sistemas.lat'

# ==========================================
# FORMATOS DE FECHA Y HORA
# ==========================================
DATETIME_FORMAT = 'd/m/Y H:i'
DATE_FORMAT = 'd/m/Y'
TIME_FORMAT = 'H:i'
SHORT_DATETIME_FORMAT = 'd/m/Y H:i'
SHORT_DATE_FORMAT = 'd/m/Y'

DATETIME_INPUT_FORMATS = [
    '%d/%m/%Y %H:%M',
    '%Y-%m-%dT%H:%M',
]

# --- CONFIGURACIÓN PARA LÍMITES MASIVOS ---
DATA_UPLOAD_MAX_NUMBER_FIELDS = 100000
DATA_UPLOAD_MAX_MEMORY_SIZE = 52428800

# ==========================================
# CONFIGURACIÓN DE TWILIO (WHATSAPP Y SMS)
# ==========================================

TWILIO_ACCOUNT_SID = 'SID_FALSO_AQUI'
TWILIO_AUTH_TOKEN = 'TOKEN_FALSO_AQUI'
TWILIO_PHONE_NUMBER = 'whatsapp:NUM_FALSO_AQUI'