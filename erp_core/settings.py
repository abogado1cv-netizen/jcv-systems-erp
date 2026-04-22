import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = 'django-insecure-aakcsii+s+*f$%&7t5lm2w50m@9q(1)%4&lw3ffvm@!2vm%)@y'
DEBUG = False
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
# GESTIÓN DE ARCHIVOS SUBIDOS (NUEVO)
# ==========================================
# Aquí es donde se guardarán los PDFs de las OC y evidencias
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

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
# CONFIGURACIÓN DEL MENÚ GPHARMA (JAZZMIN PREMIUM)
# ==========================================
JAZZMIN_SETTINGS = {
    "site_title": "GPHARMA ERP",
    "site_header": "Gestión Integral",
    "site_brand": "GPHARMA",
    
    "site_logo": "logo.png",
    "login_logo": "logo.png",
    "custom_css": "custom_login.css", 
    "site_logo_classes": "img-fluid",
    "login_logo_classes": "img-fluid", 
    
    "welcome_sign": "Gestión Integral GPHARMA - Bienvenido",
    
    "topmenu_links": [
        {"name": "📊 KPIs Licitaciones", "url": "/dashboard/licitaciones/"},
        {"name": "📁 KPIs Contratos", "url": "/dashboard/contratos/"},
        {"name": "🛒 KPIs Compras", "url": "/dashboard/compras/"},
        {"name": "📦 Almacén / Inventario", "url": "/dashboard/inventario/"},
        {'name': '🚚 KPIs Logística (OPMs)', 'url': '/dashboard/ordenes/'},
        {"name": " Atención (Pronto)", "url": "#"},
    ],

    "icons": {
        "licitaciones.Licitacion": "fas fa-chart-line",           
        "licitaciones.CatalogoMedicamento": "fas fa-pills", 
        "licitaciones.ContratoMaestro": "fas fa-file-signature", 
        "licitaciones.Contrato": "fas fa-database",          
        "licitaciones.Empresa": "fas fa-building",           
        "licitaciones.RegistroUbicacion": "fas fa-map-marker-alt",
        
        # --- LOS NUEVOS MÓDULOS ---
        "licitaciones.SocioComercial": "fas fa-handshake-angle", 
        "licitaciones.Proveedor": "fas fa-truck-field",          
        "licitaciones.OrdenCompra": "fas fa-file-invoice-dollar",
        "licitaciones.Inventario": "fas fa-boxes",               
        
        "auth.Group": "fas fa-users",
        "auth.User": "fas fa-user-circle",
    },
    
    "order_with_respect_to": ["licitaciones.Licitacion", "licitaciones.ContratoMaestro", "licitaciones.Empresa", "licitaciones.CatalogoMedicamento"],
    "show_sidebar": True,
    "navigation_expanded": True,
    "hide_apps": ["auth"], 
}

JAZZMIN_UI_TWEAKS = {
    "navbar_small_text": False,
    "footer_small_text": False,
    "body_small_text": False,
    "brand_small_text": False,
    "brand_colour": "navbar-white", 
    "accent": "accent-primary",
    "navbar": "navbar-white navbar-light",
    "no_navbar_border": False,
    "navbar_fixed": True,
    "layout_boxed": False,
    "footer_fixed": False,
    "sidebar_fixed": True,
    "sidebar": "sidebar-light-primary",
    "sidebar_nav_small_text": False,
    "sidebar_disable_expand": False,
    "sidebar_nav_child_indent": False,
    "sidebar_nav_compact_style": False,
    "sidebar_nav_legacy_style": False,
    "sidebar_nav_flat_style": True, 
    "theme": "lumen", 
    "dark_mode_theme": None,
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
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = 'sagomedical.licitaciones@gmail.com'
EMAIL_HOST_PASSWORD = 'vwubooksybnnbcpa'

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