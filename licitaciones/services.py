# Archivo: licitaciones/services.py
from dataclasses import dataclass, field
from datetime import datetime
import random

@dataclass
class KPI:
    nombre: str
    valor: float
    unidad: str
    variacion: float
    descripcion: str = ""

    def tendencia(self) -> str:
        return "▲" if self.variacion >= 0 else "▼"

@dataclass
class Alerta:
    tipo: str
    mensaje: str
    hora: datetime = field(default_factory=datetime.now)
    ICONOS = {"info": "ℹ", "warning": "⚠", "error": "✖", "success": "✔"}

@dataclass
class Producto:
    nombre: str
    ventas: float
    crecimiento: float

class DashboardService:
    PERIODOS = {
        "Q1": {"label": "1er Trimestre", "factor": 0.25},
        "Q2": {"label": "2do Trimestre", "factor": 0.50},
        "Q3": {"label": "3er Trimestre", "factor": 0.75},
        "Q4": {"label": "4to Trimestre", "factor": 1.00},
        "YTD": {"label": "Año a la fecha", "factor": 0.85},
    }

    def __init__(self, empresa: str):
        self.empresa = empresa
        self._base_ingresos = 125000000.0

    def obtener_kpis(self, periodo: str = "YTD"):
        factor = self.PERIODOS[periodo]["factor"]
        return [
            KPI("Ingresos Totales", self._base_ingresos * factor, "$", 12.5),
            KPI("Margen Operativo", 24.5, "%", 2.1),
            KPI("Nuevos Contratos", int(45 * factor), "u", -5.0),
            KPI("Satisfacción Cliente", 94.2, "pts", 1.2),
        ]

    def ingresos_por_area(self):
        return {"Sector Salud Público": 0.65, "Hospitales Privados": 0.25, "Exportación": 0.10}

    def top_productos(self):
        return [
            Producto("Lenalidomida 10mg", 15000000, 18.5),
            Producto("Paracetamol 500mg IV", 8500000, 5.2),
            Producto("Topiramato 25mg", 6200000, -2.1),
        ]

    def objetivos_trimestrales(self):
        return [{"obj": "Expansión a Zona Norte", "progreso": 75}, {"obj": "Certificación ISO 9001", "progreso": 100}, {"obj": "Reducción de costos logísticos", "progreso": 40}]

    def obtener_alertas(self):
        return [
            Alerta("error", "Stock crítico: Topiramato (Quedan 500 uds)"),
            Alerta("warning", "Contrato DAyF/0183 vence en 30 días"),
            Alerta("success", "Pago recibido: IMSS Delegación Sur"),
            Alerta("info", "Actualización del catálogo completada"),
        ]

    # --- LA MAGIA NUEVA: DEVOLVER DATOS PARA DJANGO ---
    def obtener_datos_dashboard(self, periodo="YTD"):
        return {
            "periodo": self.PERIODOS[periodo]["label"],
            "kpis": self.obtener_kpis(periodo),
            "ingresos_area": [{"area": k, "pct": v * 100, "monto": (self._base_ingresos * self.PERIODOS[periodo]["factor"]) * v} for k, v in self.ingresos_por_area().items()],
            "top_productos": self.top_productos(),
            "objetivos": self.objetivos_trimestrales(),
            "alertas": self.obtener_alertas()
        }