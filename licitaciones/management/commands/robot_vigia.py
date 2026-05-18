from django.core.management.base import BaseCommand
from django.utils import timezone
from datetime import timedelta
from licitaciones.models import Licitacion, PerfilEquipo
from twilio.rest import Client
from django.conf import settings

class Command(BaseCommand):
    help = 'Ejecuta el Robot Vigía para notificar aperturas y fallos diarios.'

    def handle(self, *args, **kwargs):
        hoy = timezone.localdate()
        ayer = hoy - timedelta(days=1)

        # 1. Buscamos al equipo que debe recibir la alerta (Analistas y Coordinadores)
        equipo = PerfilEquipo.objects.filter(rol__in=['ANALISTA_LICITA', 'COORD_COMERCIAL'], activo=True)

        if not equipo.exists():
            self.stdout.write(self.style.WARNING('No hay equipo comercial para notificar.'))
            return

        client = Client(settings.TWILIO_ACCOUNT_SID, settings.TWILIO_AUTH_TOKEN)
        mensajes_enviados = 0

        # ==========================================
        # 🤖 MISIÓN 1: APERTURAS DE AYER (Sin marcas)
        # ==========================================
        # Buscamos eventos cuya fecha de apertura fue ayer
        licitaciones_ayer = Licitacion.objects.filter(fecha_apertura__date=ayer)
        
        for lic in licitaciones_ayer:
            # Aquí el robot pregunta: "¿Están vacías las marcas/partidas?"
            # Si no hay partidas registradas aún, disparamos la alerta de regaño:
            if not lic.partidas.exists(): 
                mensaje = f"🚨 *GPHARMA - Acción Requerida*\n\n" \
                          f"Ayer fue la apertura del evento *{lic.num_procedimiento}* de {lic.get_dependencia_display()}.\n\n" \
                          f"⚠️ No has capturado tus marcas en el evento y es importante para dar continuidad al proceso."
                
                for miembro in equipo:
                    if miembro.whatsapp:
                        try:
                            client.messages.create(from_=settings.TWILIO_PHONE_NUMBER, body=mensaje, to=f"whatsapp:{miembro.whatsapp}")
                            mensajes_enviados += 1
                        except Exception: pass

        # ==========================================
        # 🤖 MISIÓN 2: RECORDATORIO DE FALLOS (Hoy)
        # ==========================================
        # Buscamos eventos cuyo fallo está programado para el día de hoy
        licitaciones_fallo = Licitacion.objects.filter(fecha_fallo__date=hoy)
        
        for lic in licitaciones_fallo:
            mensaje = f"🔔 *GPHARMA - Día de Fallo*\n\n" \
                      f"Hoy sale el fallo del evento *{lic.num_procedimiento}* de {lic.get_dependencia_display()}.\n\n" \
                      f"Recuerda tres puntos importantes:\n" \
                      f"1️⃣ Captura los resultados del evento.\n" \
                      f"2️⃣ Notifica los resultados desde el sistema.\n" \
                      f"3️⃣ Si se difiere, modifica la fecha de fallo en el sistema.\n\n¡Éxito!"
            
            for miembro in equipo:
                if miembro.whatsapp:
                    try:
                        client.messages.create(from_=settings.TWILIO_PHONE_NUMBER, body=mensaje, to=f"whatsapp:{miembro.whatsapp}")
                        mensajes_enviados += 1
                    except Exception: pass

        self.stdout.write(self.style.SUCCESS(f'✅ Robot Vigía terminó exitosamente. Mensajes enviados: {mensajes_enviados}'))