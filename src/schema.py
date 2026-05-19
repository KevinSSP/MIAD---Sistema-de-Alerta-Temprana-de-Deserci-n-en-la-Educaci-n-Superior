"""Esquema de variables del modelo de deserción.

Definiciones centralizadas para mantener un único origen de la verdad entre
el formulario, la inferencia y los insights.
"""
from __future__ import annotations

# Columnas numéricas (StandardScaler en el preprocesador)
NUM_COLS = [
    "edad_ingreso",
    "promedio_academico",
    "materias_reprobadas",
    "distancia_hogar_ies_km",
    "puntaje_icfes_percentil",
]

# Columnas categóricas (OneHotEncoder drop='first')
CAT_COLS = [
    "sexo",
    "estado_civil",
    "zona_residencia",
    "departamento",
    "nivel_educativo_padre",
    "nivel_educativo_madre",
    "sector_ies",
    "nivel_formacion",
    "metodologia",
    "area_conocimiento",
]

# Columnas ordinales / discretas escaladas
ORD_COLS = ["semestre_cursado", "estrato_socioeconomico"]

# Columnas remainder (pasan tal cual al modelo)
REM_COLS = [
    "anio_registro",
    "trabaja_mientras_estudia",
    "beneficiario_icetex",
    "beneficiario_beca",
]

# Orden exacto requerido por el preprocesador / modelo (raw DataFrame).
RAW_COLS = NUM_COLS + CAT_COLS + ORD_COLS + REM_COLS

# Variables binarias 0/1
BIN_COLS = ["trabaja_mientras_estudia", "beneficiario_icetex", "beneficiario_beca"]

# Categorías conocidas (extraídas del preprocesador entrenado)
CATEGORIES: dict[str, list[str]] = {
    "sexo": ["Hombre", "Mujer"],
    "estado_civil": ["Casado", "Otro", "Soltero", "Union libre"],
    "zona_residencia": ["Rural", "Urbana"],
    "departamento": [
        "Antioquia", "Arauca", "Atlántico", "Bogotá D.C.", "Bolívar", "Boyacá",
        "Caldas", "Caquetá", "Casanare", "Cauca", "Cesar", "Chocó",
        "Cundinamarca", "Córdoba", "Huila", "La Guajira", "Magdalena", "Meta",
        "Nariño", "Norte de Santander", "Putumayo", "Quindío", "Risaralda",
        "Santander", "Sucre", "Tolima", "Valle del Cauca",
    ],
    "nivel_educativo_padre": [
        "Ninguno", "Primaria", "Bachillerato", "Tecnico", "Universitario", "Posgrado",
    ],
    "nivel_educativo_madre": [
        "Ninguno", "Primaria", "Bachillerato", "Tecnico", "Universitario", "Posgrado",
    ],
    "sector_ies": ["Privado", "Publico"],
    "nivel_formacion": ["Tecnico profesional", "Tecnologico", "Universitario"],
    "metodologia": ["Presencial", "Distancia tradicional", "Virtual"],
    "area_conocimiento": [
        "Agronomia, veterinaria y afines",
        "Bellas artes",
        "Ciencias de la educacion",
        "Ciencias de la salud",
        "Ciencias sociales y humanas",
        "Economia, administracion, contaduria y afines",
        "Ingenieria, arquitectura, urbanismo y afines",
        "Matematicas y ciencias naturales",
    ],
}

# Etiquetas amigables en español para mostrar en la UI
LABELS: dict[str, str] = {
    "edad_ingreso": "Edad de ingreso",
    "promedio_academico": "Promedio académico (0-5)",
    "materias_reprobadas": "Materias reprobadas (acumuladas)",
    "distancia_hogar_ies_km": "Distancia hogar–IES (km)",
    "puntaje_icfes_percentil": "Puntaje ICFES (percentil 0-100)",
    "sexo": "Sexo",
    "estado_civil": "Estado civil",
    "zona_residencia": "Zona de residencia",
    "departamento": "Departamento",
    "nivel_educativo_padre": "Nivel educativo del padre",
    "nivel_educativo_madre": "Nivel educativo de la madre",
    "sector_ies": "Sector de la IES",
    "nivel_formacion": "Nivel de formación",
    "metodologia": "Metodología",
    "area_conocimiento": "Área de conocimiento",
    "semestre_cursado": "Semestre cursado",
    "estrato_socioeconomico": "Estrato socioeconómico (1-6)",
    "anio_registro": "Año de registro",
    "trabaja_mientras_estudia": "¿Trabaja mientras estudia?",
    "beneficiario_icetex": "¿Beneficiario ICETEX?",
    "beneficiario_beca": "¿Beneficiario de beca?",
}

# Agrupación visual para el formulario individual
FORM_GROUPS = {
    "Demográfico": [
        "edad_ingreso", "sexo", "estado_civil", "zona_residencia", "departamento",
    ],
    "Socioeconómico": [
        "estrato_socioeconomico", "nivel_educativo_padre", "nivel_educativo_madre",
        "trabaja_mientras_estudia", "beneficiario_icetex", "beneficiario_beca",
        "distancia_hogar_ies_km",
    ],
    "Académico": [
        "puntaje_icfes_percentil", "promedio_academico", "materias_reprobadas",
        "semestre_cursado", "sector_ies", "nivel_formacion", "metodologia",
        "area_conocimiento", "anio_registro",
    ],
}

# Umbrales de riesgo (semáforo)
RISK_THRESHOLDS = {"medio": 0.30, "alto": 0.60}
RISK_COLORS = {"Bajo": "#22c55e", "Medio": "#eab308", "Alto": "#ef4444"}


def risk_tier(prob: float) -> str:
    if prob >= RISK_THRESHOLDS["alto"]:
        return "Alto"
    if prob >= RISK_THRESHOLDS["medio"]:
        return "Medio"
    return "Bajo"
