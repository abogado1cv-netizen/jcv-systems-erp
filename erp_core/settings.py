import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent

# 👇 CARGA EL .ENV DE LA EMPRESA CORRESPONDIENTE (Por defecto usa '.env') 👇
env_file = os.getenv('ENV_FILE', BASE_DIR / '.env')
load_dotenv(env_file)

# ==========================================
# 🔐 IDENTIDAD Y SEGURIDAD POR INSTANCIA
# ==========================================
EMPRESA_SLUG = os.getenv('EMPRESA_SLUG', 'gpharma') # Ej: sago, gsm, gams, empresa4...
SECRET_KEY = os.getenv('SECRET_KEY', 'django-insecure-aakcsii+s+*f$%&7t5lm2w50m@9q(1)%4&lw3ffvm@!2vm%)@y')
DEBUG = os.getenv('DEBUG', 'True').lower() == 'true'
ALLOWED_HOSTS = [h.strip() for h in os.getenv('ALLOWED_HOSTS', '*').split(',') if h.strip()]
CSRF_TRUSTED_ORIGINS = [
    'https://*.jcv-sistemas.lat',
]
SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')

# Evita que las sesiones choquen si un admin abre 2 empresas en el mismo navegador
SESSION_COOKIE_NAME = f"sessionid_{EMPRESA_SLUG}"
CSRF_COOKIE_NAME = f"csrftoken_{EMPRESA_SLUG}"

INSTALLED_APPS = [
    'jazzmin',
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django.contrib.humanize',
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

# ==========================================
# 🗄️ BASE DE DATOS INDEPENDIENTE POR EMPRESA
# ==========================================
DB_NAME = os.getenv('DB_NAME', f'db_{EMPRESA_SLUG}.sqlite3')

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / DB_NAME,
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
# 📂 ARCHIVOS SUBIDOS AISLADOS POR EMPRESA
# ==========================================
MEDIA_URL = os.getenv('MEDIA_URL', f'/media/{EMPRESA_SLUG}/')
MEDIA_ROOT = os.getenv('MEDIA_ROOT', f'/var/www/media/{EMPRESA_SLUG}/')

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
# 🎛️ CONTROL DINÁMICO DE MÓDULOS (DESDE .ENV)
# ==========================================
# 1. Modelos que SIEMPRE van ocultos + los que quieras ocultar por empresa
BASE_HIDE_MODELS = [
    "licitaciones.estatusprocedimiento",
    "licitaciones.configuracionemail",
    "licitaciones.partidarequerimiento",
]
EXTRA_HIDE_MODELS = [
    m.strip() for m in os.getenv('HIDE_MODELS_EXTRA', '').split(',') if m.strip()
]

# 2. Menú superior dinámico según los módulos activos en el .env
CATALOGO_TOPMENU = {
    "comercial": {"name": "📊 Comercial", "url": "/dashboard/licitaciones/"},
    "precios":   {"name": "🔍 Historial de Precios", "url": "/dashboard/historial-precios/"},
    "contratos": {"name": "📁 Contratos", "url": "/dashboard/contratos/"},
    "compras":   {"name": "🛒 Compras", "url": "/dashboard/compras/"},
    "logistica": {"name": "🚚 Logística", "url": "/dashboard/ordenes/"},
    "almacen":   {"name": "📦 Almacén", "url": "/dashboard/inventario/"},
}

TOPMENU_ACTIVOS = [
    k.strip().lower() for k in os.getenv('TOPMENU_MODULOS', 'comercial,precios,contratos,compras,logistica,almacen').split(',') if k.strip()
]

TOPMENU_LINKS = [{"name": "Inicio", "url": "admin:index", "permissions": ["auth.view_user"]}]
for modulo in TOPMENU_ACTIVOS:
    if modulo in CATALOGO_TOPMENU:
        TOPMENU_LINKS.append(CATALOGO_TOPMENU[modulo])

# ==========================================
# 🎨 CONFIGURACIÓN VISUAL (JAZZMIN WHITE-LABEL)
# ==========================================
SITE_BRAND = os.getenv('SITE_BRAND', 'GPHARMA')
SITE_LOGO = os.getenv('SITE_LOGO', 'logo.png')

JAZZMIN_SETTINGS = {
    "site_title": os.getenv('SITE_TITLE', f'{SITE_BRAND} ERP'),
    "site_header": "Módulos",
    "site_brand": SITE_BRAND,
    
    "site_logo": SITE_LOGO,
    "login_logo": SITE_LOGO,
    "custom_css": "css/aura_pro.css", 
    "site_logo_classes": "img-fluid",
    "login_logo_classes": "img-fluid", 
    
    "welcome_sign": os.getenv('WELCOME_SIGN', f'Bienvenido al Centro de Control {SITE_BRAND}'),
    "copyright": os.getenv('COPYRIGHT_TEXT', 'JCV Systems'),
    
    "topmenu_links": TOPMENU_LINKS,

    "custom_links": {
        "licitaciones": [{
            "name": "Gestión de Equipo",
            "url": "/dashboard/equipo/",
            "icon": "fas fa-users",
        }]
    },

    "hide_models": BASE_HIDE_MODELS + EXTRA_HIDE_MODELS,
    "hide_apps": ["auth"], 

    "icons": {
        "licitaciones.catalogomedicamento": "fas fa-capsules", 
        "licitaciones.empresa": "fas fa-building",           
        "licitaciones.sociocomercial": "fas fa-handshake", 
        "licitaciones.almacen": "fas fa-warehouse",
        "licitaciones.licitacion": "fas fa-folder-open",
        "licitaciones.cotizacion": "fas fa-file-invoice",
        "licitaciones.contrato": "fas fa-file-signature",          
        "licitaciones.ordencompra": "fas fa-shopping-cart",
        "licitaciones.entradaalmacen": "fas fa-box-open",
        "licitaciones.inventario": "fas fa-boxes",               
        "licitaciones.traspasointercompany": "fas fa-exchange-alt",
        "licitaciones.incidenciainventario": "fas fa-radiation-alt", 
        "licitaciones.escanerkardex": "fas fa-barcode",
        "licitaciones.ordensuministro": "fas fa-truck-loading",
        "licitaciones.pedidodirecto": "fas fa-paper-plane",
        "licitaciones.remisionentrega": "fas fa-receipt",
        "licitaciones.registroubicacion": "fas fa-map-marker-alt",
        "auth.user": "fas fa-user-circle",
        "auth.group": "fas fa-users-cog",
    },
    
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
    "show_ui_builder": False, 
}

JAZZMIN_UI_TWEAKS = {
    "navbar_small_text": False,
    "footer_small_text": False,
    "body_small_text": False,
    "brand_small_text": False,
    "brand_colour": False, 
    "navbar": False,
    "no_navbar_border": True,
    "navbar_fixed": True,
    "layout_boxed": False,
    "footer_fixed": False,
    "sidebar_fixed": True,
    "sidebar": "sidebar-dark-primary",
    "sidebar_nav_small_text": False,
    "sidebar_disable_expand": False,
    "sidebar_nav_child_indent": True,
    "sidebar_nav_compact_style": False,
    "sidebar_nav_legacy_style": False,
    "sidebar_nav_flat_style": True, 
    "theme": "default", 
    "dark_mode_theme": "darkly",
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
DEFAULT_FROM_EMAIL = os.getenv('DEFAULT_FROM_EMAIL', 'notificaciones@jcv-sistemas.lat')

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

DATA_UPLOAD_MAX_NUMBER_FIELDS = 100000
DATA_UPLOAD_MAX_MEMORY_SIZE = 52428800

# ==========================================
# CONFIGURACIÓN DE TWILIO (WHATSAPP Y SMS)
# ==========================================
TWILIO_ACCOUNT_SID = os.getenv('TWILIO_ACCOUNT_SID', 'SID_FALSO_AQUI')
TWILIO_AUTH_TOKEN = os.getenv('TWILIO_AUTH_TOKEN', 'TOKEN_FALSO_AQUI')
TWILIO_PHONE_NUMBER = os.getenv('TWILIO_PHONE_NUMBER', 'whatsapp:NUM_FALSO_AQUI')