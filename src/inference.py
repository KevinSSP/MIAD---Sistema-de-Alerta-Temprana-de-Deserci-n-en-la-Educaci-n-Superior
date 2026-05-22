"""Carga del modelo serializado e inferencia por lote."""
from __future__ import annotations

import pickle
import warnings
from functools import lru_cache
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

# --- Shim de compatibilidad sklearn ------------------------------------------------
# El pickle fue generado con scikit-learn 1.6.1 y referencia una clase privada
# (`_RemainderColsList`) que ya no existe en sklearn >= 1.7. Le creamos un alias
# antes de cargar para evitar AttributeError sin necesidad de re-entrenar.
import sklearn.compose._column_transformer as _ct  # noqa: E402
if not hasattr(_ct, "_RemainderColsList"):
    class _RemainderColsList(list):  # pragma: no cover - shim
        pass
    _ct._RemainderColsList = _RemainderColsList

from .schema import (  # noqa: E402
    BIN_COLS,
    CATEGORIES,
    NUM_COLS,
    ORD_COLS,
    RAW_COLS,
    REM_COLS,
    RISK_THRESHOLDS,
    risk_tier,
)

MODEL_PATH = Path(__file__).resolve().parent.parent / "Code" / "modelo_xgboost_desercion_V1.1.1.pkl"


@lru_cache(maxsize=1)
def load_model():
    """Carga el dict {'xgboost', 'preprocessor'} desde disco."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with open(MODEL_PATH, "rb") as fh:
            bundle = pickle.load(fh)
    if not isinstance(bundle, dict) or "xgboost" not in bundle or "preprocessor" not in bundle:
        raise RuntimeError("El archivo de modelo no contiene las llaves esperadas.")
    return bundle


def expected_columns() -> list[str]:
    return list(RAW_COLS)


def _coerce_row(row: dict) -> dict:
    """Normaliza tipos y rellena defaults razonables para una fila individual."""
    out = dict(row)
    # Booleanos a 0/1
    for c in BIN_COLS:
        v = out.get(c, 0)
        if isinstance(v, bool):
            out[c] = int(v)
        elif isinstance(v, str):
            out[c] = 1 if v.strip().lower() in {"si", "sí", "true", "1", "yes"} else 0
        else:
            out[c] = int(v) if pd.notna(v) else 0
    # Numéricos
    for c in NUM_COLS + ORD_COLS + ["anio_registro"]:
        v = out.get(c)
        try:
            out[c] = float(v) if v is not None and v != "" else np.nan
        except (TypeError, ValueError):
            out[c] = np.nan
    return out


def build_dataframe(records: Iterable[dict]) -> pd.DataFrame:
    """Construye un DataFrame con las columnas en el orden esperado."""
    rows = [_coerce_row(r) for r in records]
    df = pd.DataFrame(rows)
    for c in RAW_COLS:
        if c not in df.columns:
            df[c] = np.nan
    return df[RAW_COLS]


class SchemaError(ValueError):
    pass


def validate_dataframe(df: pd.DataFrame) -> list[str]:
    """Valida columnas y valores categóricos. Devuelve lista de advertencias."""
    warns: list[str] = []
    missing = [c for c in RAW_COLS if c not in df.columns]
    if missing:
        raise SchemaError(f"Faltan columnas requeridas: {missing}")
    for c, cats in CATEGORIES.items():
        bad = sorted(set(df[c].dropna().astype(str)) - set(cats))
        if bad:
            warns.append(f"Valores no vistos en '{c}': {bad} (serán ignorados por el modelo)")
    return warns


def _prepare(df: pd.DataFrame) -> pd.DataFrame:
    df_in = df.copy()
    for c in NUM_COLS + ORD_COLS + ["anio_registro"]:
        df_in[c] = pd.to_numeric(df_in[c], errors="coerce")
    for c in BIN_COLS:
        df_in[c] = pd.to_numeric(df_in[c], errors="coerce").fillna(0).astype(int)
    for c in CATEGORIES:
        df_in[c] = df_in[c].astype(str)
    return df_in


def _transform(df_in: pd.DataFrame) -> np.ndarray:
    """Aplica preprocesador y, si hace falta, concatena columnas remainder.

    Algunos bundles del modelo perdieron la lista de columnas remainder en
    el ``ColumnTransformer`` y requieren que se agreguen manualmente
    (``anio_registro`` + binarios). Otros bundles ya las incluyen y agregarlas
    duplicaría features. Detectamos cuál es el caso comparando con el número
    de features esperado por el booster.
    """
    bundle = load_model()
    pre = bundle["preprocessor"]
    mdl = bundle["xgboost"]
    X_pre = pre.transform(df_in[RAW_COLS])
    if hasattr(X_pre, "toarray"):
        X_pre = X_pre.toarray()
    X_pre = np.asarray(X_pre, dtype=float)
    if X_pre.ndim == 1:
        X_pre = X_pre.reshape(1, -1)

    expected = getattr(mdl, "n_features_in_", None)
    if expected is None:
        try:
            expected = mdl.get_booster().num_features()
        except Exception:
            expected = X_pre.shape[1] + len(REM_COLS)

    n_pre = X_pre.shape[1]
    if n_pre == expected:
        return X_pre
    if n_pre + len(REM_COLS) == expected:
        rem = df_in[REM_COLS].to_numpy(dtype=float)
        if rem.ndim == 1:
            rem = rem.reshape(1, -1)
        return np.hstack([X_pre, rem])
    raise RuntimeError(
        f"Inconsistencia de features: el preprocesador produjo {n_pre} columnas "
        f"y el modelo espera {expected}. Verifique la versión del bundle."
    )


def predict(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve DataFrame con columnas: prob_desercion, riesgo, prediccion."""
    bundle = load_model()
    mdl = bundle["xgboost"]
    df_in = _prepare(df)
    X = _transform(df_in)
    proba = mdl.predict_proba(X)[:, 1]
    out = pd.DataFrame({
        "prob_desercion": proba,
        "riesgo": [risk_tier(float(p)) for p in proba],
        "prediccion": (proba >= RISK_THRESHOLDS["alto"]).astype(int),
    })
    return out


def transform_features(df: pd.DataFrame) -> tuple[np.ndarray, list[str]]:
    """Aplica el pipeline completo y devuelve (X, nombres de features del modelo)."""
    bundle = load_model()
    mdl = bundle["xgboost"]
    df_in = _prepare(df)
    X = _transform(df_in)
    try:
        names = list(mdl.get_booster().feature_names)
    except Exception:
        names = [f"f{i}" for i in range(X.shape[1])]
    if X.shape[1] != len(names):
        names = names[: X.shape[1]] + [f"f{i}" for i in range(len(names), X.shape[1])]
    return np.asarray(X), names
