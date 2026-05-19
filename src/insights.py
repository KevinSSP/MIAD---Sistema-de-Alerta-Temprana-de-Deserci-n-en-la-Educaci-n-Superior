"""Insights por estudiante y a nivel de cohorte."""
from __future__ import annotations

import numpy as np
import pandas as pd

from .inference import load_model, transform_features
from .schema import LABELS, RISK_COLORS


# Mapa de nombres internos del modelo → variable raíz humana
def _root_var(feature_name: str) -> str:
    name = feature_name
    for prefix in ("num__", "cat__", "ord__", "remainder__"):
        if name.startswith(prefix):
            name = name[len(prefix):]
            break
    # Para one-hot: la categoría va después de un '_' tras el nombre de variable
    for raw in LABELS:
        if name == raw:
            return raw
        if name.startswith(raw + "_"):
            return raw
    return name


def top_drivers(record: dict, k: int = 5) -> pd.DataFrame:
    """Devuelve los k factores que más empujan la predicción para un estudiante.

    Usa los valores SHAP nativos de XGBoost (`pred_contribs=True`) para
    asignar una contribución exacta de cada feature a la predicción.
    """
    bundle = load_model()
    mdl = bundle["xgboost"]
    df = pd.DataFrame([record])
    X, names = transform_features(df)
    try:
        import xgboost as xgb
        dmat = xgb.DMatrix(X, feature_names=names)
        contribs = mdl.get_booster().predict(dmat, pred_contribs=True)
        # contribs shape: (1, n_features + 1) — última columna es bias
        shap = contribs[0, :-1]
    except Exception:
        shap = np.zeros(len(names))

    # Agregar por variable raíz (para no repetir mismas variables en one-hot)
    agg: dict[str, float] = {}
    for n, v in zip(names, shap):
        root = _root_var(n)
        agg[root] = agg.get(root, 0.0) + float(v)

    rows = [
        {
            "factor": LABELS.get(root, root),
            "variable": root,
            "contribucion": val,
            "direccion": "Riesgo" if val > 0 else "Protección",
            "abs": abs(val),
        }
        for root, val in agg.items()
    ]
    out = pd.DataFrame(rows).sort_values("abs", ascending=False).head(k).reset_index(drop=True)
    total_abs = out["abs"].sum()
    out["contribucion_rel"] = out["abs"] / total_abs if total_abs else 0.0
    return out


def recommendation_text(prob: float, drivers: pd.DataFrame) -> str:
    """Genera una recomendación breve para consejería académica."""
    if prob >= 0.6:
        base = (
            "Riesgo alto de deserción. Se sugiere intervención prioritaria: "
            "tutoría académica personalizada, revisión de carga, acompañamiento "
            "psicosocial y verificación de apoyos financieros."
        )
    elif prob >= 0.3:
        base = (
            "Riesgo medio. Se recomienda seguimiento mensual, oferta de "
            "tutorías focalizadas en materias críticas y orientación vocacional."
        )
    else:
        base = (
            "Riesgo bajo. Mantener seguimiento estándar y reforzar programas "
            "de bienestar para sostener la permanencia."
        )
    if not drivers.empty:
        top = drivers.head(3)["factor"].tolist()
        base += f" Factores más influyentes: {', '.join(top)}."
    return base


def cohort_summary(df_raw: pd.DataFrame, preds: pd.DataFrame) -> dict:
    """KPIs y datasets agregados para el tablero global."""
    df = df_raw.reset_index(drop=True).copy()
    df["prob_desercion"] = preds["prob_desercion"].values
    df["riesgo"] = preds["riesgo"].values

    total = len(df)
    by_risk = df["riesgo"].value_counts().reindex(["Alto", "Medio", "Bajo"]).fillna(0).astype(int)

    kpis = {
        "total": total,
        "alto": int(by_risk.get("Alto", 0)),
        "medio": int(by_risk.get("Medio", 0)),
        "bajo": int(by_risk.get("Bajo", 0)),
        "pct_alto": float(by_risk.get("Alto", 0)) / total if total else 0.0,
        "prom_riesgo": float(df["prob_desercion"].mean()) if total else 0.0,
    }

    by_estrato = (
        df.groupby("estrato_socioeconomico", dropna=False)["prob_desercion"]
        .mean().reset_index().sort_values("estrato_socioeconomico")
    )
    by_beca = df.groupby("beneficiario_beca", dropna=False)["prob_desercion"].mean().reset_index()
    by_sector = df.groupby("sector_ies", dropna=False)["prob_desercion"].mean().reset_index()
    by_area = (
        df.groupby("area_conocimiento", dropna=False)["prob_desercion"]
        .agg(["mean", "count"]).reset_index().sort_values("mean", ascending=False)
    )

    return {
        "df": df,
        "kpis": kpis,
        "by_risk": by_risk.reset_index().rename(columns={"index": "riesgo", "count": "n"}),
        "by_estrato": by_estrato,
        "by_beca": by_beca,
        "by_sector": by_sector,
        "by_area": by_area,
        "risk_colors": RISK_COLORS,
    }
