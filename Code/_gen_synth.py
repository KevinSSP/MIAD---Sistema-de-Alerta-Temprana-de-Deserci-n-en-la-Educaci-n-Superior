"""Genera un dataset sintético de estudiantes (CSV) compatible con el modelo.

Uso:
    python Code/_gen_synth.py
Produce: Code/estudiantes_sinteticos_11500.csv
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

# Permitir importar src.schema
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.schema import CATEGORIES, RAW_COLS  # noqa: E402

N = 11500
RNG = np.random.default_rng(seed=20260518)

# Pesos por categoría (aproximan distribuciones colombianas plausibles)
DEPT_POP = {
    "Bogotá D.C.": 18, "Antioquia": 13, "Valle del Cauca": 9, "Cundinamarca": 7,
    "Atlántico": 5, "Santander": 4.5, "Bolívar": 4, "Nariño": 3.2, "Córdoba": 3.2,
    "Tolima": 2.8, "Cauca": 2.8, "Magdalena": 2.7, "Boyacá": 2.5, "Norte de Santander": 2.5,
    "Cesar": 2.2, "Huila": 2.2, "Risaralda": 1.9, "Caldas": 1.8, "Sucre": 1.7,
    "Meta": 1.7, "Quindío": 1.1, "La Guajira": 1.6, "Chocó": 1.0, "Caquetá": 0.8,
    "Casanare": 0.8, "Arauca": 0.5, "Putumayo": 0.4,
}


def _pick(values, weights=None):
    if weights is None:
        return RNG.choice(values, size=N)
    w = np.asarray([weights[v] for v in values], dtype=float)
    w = w / w.sum()
    return RNG.choice(values, size=N, p=w)


def generate() -> pd.DataFrame:
    df = pd.DataFrame(index=range(N))

    # --- Demográficas ---
    df["sexo"] = _pick(CATEGORIES["sexo"], {"Hombre": 49, "Mujer": 51})
    df["estado_civil"] = _pick(
        CATEGORIES["estado_civil"],
        {"Soltero": 78, "Union libre": 12, "Casado": 8, "Otro": 2},
    )
    df["zona_residencia"] = _pick(CATEGORIES["zona_residencia"], {"Urbana": 82, "Rural": 18})
    df["departamento"] = _pick(CATEGORIES["departamento"], DEPT_POP)

    df["edad_ingreso"] = np.clip(
        np.round(RNG.normal(19.5, 3.2, N)).astype(int), 15, 45
    )

    # --- Socioeconómico ---
    df["estrato_socioeconomico"] = RNG.choice(
        [1, 2, 3, 4, 5, 6], size=N, p=[0.18, 0.34, 0.28, 0.12, 0.06, 0.02]
    )

    niveles = CATEGORIES["nivel_educativo_padre"]
    base_p = np.array([0.06, 0.30, 0.34, 0.14, 0.13, 0.03])
    df["nivel_educativo_padre"] = RNG.choice(niveles, size=N, p=base_p)
    base_m = np.array([0.05, 0.28, 0.36, 0.15, 0.13, 0.03])
    df["nivel_educativo_madre"] = RNG.choice(niveles, size=N, p=base_m)

    df["trabaja_mientras_estudia"] = RNG.binomial(1, 0.38, N)
    df["beneficiario_icetex"] = RNG.binomial(1, 0.22, N)
    df["beneficiario_beca"] = RNG.binomial(1, 0.18, N)

    dist = RNG.lognormal(mean=2.0, sigma=0.9, size=N)
    df["distancia_hogar_ies_km"] = np.round(np.clip(dist, 0.2, 250.0), 2)

    # --- Académico ---
    df["puntaje_icfes_percentil"] = np.clip(
        np.round(RNG.normal(58, 18, N)).astype(int), 1, 100
    )

    # Promedio académico (0-5) correlacionado positivamente con ICFES
    z = (df["puntaje_icfes_percentil"] - 58) / 18
    prom = 3.4 + 0.45 * z + RNG.normal(0, 0.45, N)
    df["promedio_academico"] = np.round(np.clip(prom, 0.0, 5.0), 2)

    # Materias reprobadas: más cuando promedio es bajo
    lam = np.clip(4.5 - df["promedio_academico"], 0.05, 5.0)
    df["materias_reprobadas"] = RNG.poisson(lam=lam).clip(0, 20)

    df["semestre_cursado"] = RNG.integers(1, 13, N)

    df["sector_ies"] = _pick(CATEGORIES["sector_ies"], {"Privado": 55, "Publico": 45})
    df["nivel_formacion"] = _pick(
        CATEGORIES["nivel_formacion"],
        {"Universitario": 64, "Tecnologico": 26, "Tecnico profesional": 10},
    )
    df["metodologia"] = _pick(
        CATEGORIES["metodologia"],
        {"Presencial": 72, "Distancia tradicional": 16, "Virtual": 12},
    )
    df["area_conocimiento"] = _pick(
        CATEGORIES["area_conocimiento"],
        {
            "Economia, administracion, contaduria y afines": 28,
            "Ingenieria, arquitectura, urbanismo y afines": 24,
            "Ciencias sociales y humanas": 14,
            "Ciencias de la salud": 12,
            "Ciencias de la educacion": 9,
            "Matematicas y ciencias naturales": 5,
            "Bellas artes": 4,
            "Agronomia, veterinaria y afines": 4,
        },
    )

    df["anio_registro"] = RNG.choice([2021, 2022, 2023, 2024, 2025], size=N,
                                     p=[0.10, 0.18, 0.24, 0.26, 0.22])

    # Reordenar al esquema exacto del modelo
    df = df[RAW_COLS]
    return df


def main() -> None:
    df = generate()
    out = Path(__file__).resolve().parent / "estudiantes_sinteticos_11500.csv"
    df.to_csv(out, index=False, encoding="utf-8")
    print(f"OK -> {out}  shape={df.shape}")
    print(df.head(3).to_string())


if __name__ == "__main__":
    main()
