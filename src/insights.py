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
    if prob >= 0.85:
        base = (
            "Riesgo crítico de deserción (≥ 85 %). Intervención inmediata: "
            "asignar tutor académico, plan psicosocial, revisión de carga y "
            "validación urgente de apoyos financieros (ICETEX/beca)."
        )
    elif prob >= 0.69:
        base = (
            "Riesgo alto. Intervención prioritaria: tutoría académica "
            "personalizada, monitoreo quincenal y acompañamiento psicosocial."
        )
    elif prob >= 0.51:
        base = (
            "Riesgo medio. Seguimiento mensual con tutorías focalizadas en "
            "materias críticas y orientación vocacional."
        )
    elif prob >= 0.31:
        base = (
            "Riesgo normal. Mantener seguimiento estándar e incluir al "
            "estudiante en programas de bienestar y orientación académica."
        )
    else:
        base = (
            "Riesgo bajo. Permanencia esperada; reforzar programas de "
            "bienestar y mentoría para sostener el desempeño."
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
    risk_order = ["Crítico", "Alto", "Medio", "Normal", "Bajo"]
    by_risk = df["riesgo"].value_counts().reindex(risk_order).fillna(0).astype(int)

    prioritarios = int(by_risk.get("Crítico", 0) + by_risk.get("Alto", 0))
    kpis = {
        "total": total,
        "critico": int(by_risk.get("Crítico", 0)),
        "alto": int(by_risk.get("Alto", 0)),
        "medio": int(by_risk.get("Medio", 0)),
        "normal": int(by_risk.get("Normal", 0)),
        "bajo": int(by_risk.get("Bajo", 0)),
        "prioritarios": prioritarios,
        "pct_alto": prioritarios / total if total else 0.0,
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
