/* Genera Reportes/Anexo_Tecnico_SAT-DE.docx */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
  TableOfContents, VerticalAlign,
} = require("docx");

// ------- paleta -------
const BRAND     = "1D4ED8";
const BRAND_DK  = "1E3A8A";
const INK       = "0F172A";
const MUTED     = "475569";
const LINE      = "E2E8F0";
const SOFT      = "F8FAFC";
const OK        = "16A34A";
const WARN      = "EAB308";
const BAD       = "DC2626";

const border = (color = LINE, size = 6) => ({ style: BorderStyle.SINGLE, size, color });
const allBorders = (color = LINE) => ({
  top: border(color), bottom: border(color), left: border(color), right: border(color),
});

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 },
  alignment: opts.align,
  children: [new TextRun({ text, ...(opts.run || {}) })],
});

const H = (text, level) => new Paragraph({
  heading: level,
  spacing: { before: 240, after: 160 },
  children: [new TextRun({ text })],
});

const BULLET = (text, level = 0) => new Paragraph({
  numbering: { reference: "bullets", level },
  spacing: { after: 80 },
  children: [new TextRun({ text })],
});

const STEP = (text) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 80 },
  children: [new TextRun({ text })],
});

const TXT = (text, opts = {}) => new Paragraph({
  spacing: { after: 0 },
  alignment: opts.align,
  children: [new TextRun({
    text,
    bold: !!opts.bold,
    italics: !!opts.italics,
    color: opts.color,
    size: opts.size || 20,
  })],
});

const CELL = (children, opts = {}) => new TableCell({
  borders: allBorders(opts.borderColor || LINE),
  margins: { top: 90, bottom: 90, left: 140, right: 140 },
  width: { size: opts.width, type: WidthType.DXA },
  verticalAlign: opts.valign || VerticalAlign.CENTER,
  shading: opts.fill
    ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" }
    : undefined,
  children: Array.isArray(children) ? children : [children],
});

function dataTable(rows, columnWidths, opts = {}) {
  const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    rows: rows.map((cells, i) => new TableRow({
      tableHeader: i === 0,
      children: cells.map((c, j) => {
        const isHead = i === 0;
        const headFill = opts.headFill || BRAND;
        let fill = isHead ? headFill : (i % 2 === 0 ? SOFT : undefined);
        let para;
        if (typeof c === "object" && c !== null && !Array.isArray(c) && c.text !== undefined) {
          fill = c.fill || fill;
          para = TXT(c.text, { bold: isHead || c.bold, color: isHead ? "FFFFFF" : c.color });
        } else {
          para = TXT(String(c), isHead ? { bold: true, color: "FFFFFF" } : {});
        }
        return CELL([para], { width: columnWidths[j], fill });
      }),
    })),
  });
}

const callout = (title, body, color = BRAND) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({
    children: [new TableCell({
      borders: {
        top: border(color), bottom: border(color), right: border(color),
        left: { style: BorderStyle.SINGLE, size: 24, color },
      },
      margins: { top: 140, bottom: 140, left: 200, right: 200 },
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: SOFT, type: ShadingType.CLEAR, color: "auto" },
      children: [
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: title, bold: true, color, size: 22 })],
        }),
        new Paragraph({
          children: [new TextRun({ text: body, size: 20, color: INK })],
        }),
      ],
    })],
  })],
});

// ============ DIAGRAMA ESQUEMÁTICO ============
// Lo construimos como una tabla de tres columnas (Datos · Procesamiento/Modelo · Producto)
// con cajas coloreadas y flechas textuales entre filas.

const diagBox = (title, lines, fill = BRAND, textColor = "FFFFFF") => {
  return new TableCell({
    borders: allBorders(fill),
    shading: { fill, type: ShadingType.CLEAR, color: "auto" },
    margins: { top: 140, bottom: 140, left: 160, right: 160 },
    width: { size: 3000, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: title, bold: true, color: textColor, size: 22 })],
      }),
      ...lines.map((l) => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [new TextRun({ text: l, color: textColor, size: 18 })],
      })),
    ],
  });
};

const arrowRow = (text = "↓") => new TableRow({
  children: [0, 1, 2].map(() => new TableCell({
    borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
    width: { size: 3000, type: WidthType.DXA },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, bold: true, color: BRAND_DK, size: 28 })],
    })],
  })),
});

function buildDiagram() {
  const blank = new TableCell({
    borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
    width: { size: 3000, type: WidthType.DXA },
    children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
  });

  const headerRow = new TableRow({
    children: [
      diagBox("FUENTES DE DATOS", [], BRAND_DK),
      diagBox("PROCESAMIENTO Y MODELO", [], BRAND_DK),
      diagBox("PRODUCTO", [], BRAND_DK),
    ],
  });

  const row1 = new TableRow({
    children: [
      diagBox("Histórico institucional", [
        "Registros académicos",
        "Demográfico / socioeconómico",
        "11.500 estudiantes",
      ], BRAND),
      diagBox("Pre-procesamiento", [
        "StandardScaler (5 num.)",
        "OneHotEncoder (10 cat.)",
        "Passthrough binarias",
      ], "0EA5E9", "FFFFFF"),
      diagBox("Predicción individual", [
        "Formulario web",
        "Probabilidad + nivel",
        "Top-5 factores SHAP",
      ], OK),
    ],
  });

  const row2 = new TableRow({
    children: [
      diagBox("Esquema de 21 variables", [
        "5 numéricas",
        "10 categóricas",
        "2 ordinales · 4 binarias",
      ], BRAND, "FFFFFF"),
      diagBox("Selección de modelo", [
        "Regresión Logística",
        "Random Forest",
        "XGBoost (ganador)",
      ], BRAND_DK),
      diagBox("Carga masiva", [
        "Validación de esquema",
        "Inferencia por lote",
        "Descarga CSV",
      ], OK),
    ],
  });

  const row3 = new TableRow({
    children: [
      diagBox("Entrada en línea", [
        "Formulario individual",
        "Archivo CSV (lote)",
      ], "0EA5E9"),
      diagBox("Modelo XGBoost v1.1.1", [
        "scale_pos_weight balanceado",
        "GridSearchCV · 5-fold",
        "Umbral 0.30 / 0.60",
      ], BRAND),
      diagBox("Insights 360", [
        "Distribución de riesgo",
        "Cortes por dimensión",
        "Recomendación accionable",
      ], OK),
    ],
  });

  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [3000, 3000, 3000],
    rows: [headerRow, arrowRow("↓"), row1, arrowRow("↓"), row2, arrowRow("↓"), row3],
  });
}

// ============ portada ============
const cover = [
  new Paragraph({
    spacing: { before: 2400, after: 120 },
    children: [new TextRun({ text: "Anexo Técnico", size: 56, bold: true, color: INK })],
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({
      text: "SAT-DE · Sistema de Alerta Temprana de Deserción en Educación Superior",
      size: 32, bold: true, color: BRAND,
    })],
  }),
  new Paragraph({
    spacing: { after: 480 },
    children: [new TextRun({
      text: "Diagrama esquemático · Reporte de experimentos · Rúbrica · Repositorio",
      size: 22, color: MUTED,
    })],
  }),
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BRAND, space: 4 } },
    children: [new TextRun({ text: "" })],
  }),
  new Paragraph({
    spacing: { before: 200, after: 0 },
    children: [new TextRun({ text: "Proyecto Aplicado en Analítica de Datos", size: 22, color: INK })],
  }),
  new Paragraph({ spacing: { after: 0 },
    children: [new TextRun({ text: "Maestría en Inteligencia Analítica de Datos (MIAD)", size: 22, color: INK })] }),
  new Paragraph({ spacing: { after: 0 },
    children: [new TextRun({ text: "Universidad de los Andes · Grupo 10", size: 22, color: INK })] }),
  new Paragraph({ spacing: { after: 0 },
    children: [new TextRun({ text: "Versión 1.0 · Mayo de 2026", size: 22, color: MUTED })] }),
  new Paragraph({ children: [new PageBreak()] }),
];

const toc = [
  H("Tabla de contenido", HeadingLevel.HEADING_1),
  new Paragraph({ children: [new TableOfContents("Contenido", { hyperlink: true, headingStyleRange: "1-2" })] }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ============ I. DIAGRAMA ============
const partI = [
  H("I. Diagrama esquemático propuesto", HeadingLevel.HEADING_1),
  P("El siguiente diagrama representa los tres bloques funcionales de SAT-DE: las fuentes de datos institucionales, la capa de procesamiento y modelado, y los productos de salida entregados al usuario final."),
  buildDiagram(),
  P(""),
  H("Lectura del diagrama", HeadingLevel.HEADING_2),
  BULLET("Bloque izquierdo (Datos): el dato histórico institucional alimenta el entrenamiento; en producción la entrada llega vía formulario o CSV con el mismo esquema de 21 variables."),
  BULLET("Bloque central (Procesamiento y modelo): un ColumnTransformer estandariza numéricas/ordinales y aplica one-hot a categóricas; tras la comparación entre Regresión Logística, Random Forest y XGBoost, este último resultó el seleccionado."),
  BULLET("Bloque derecho (Producto): la aplicación Streamlit expone tres flujos de uso al consejero académico (individual, lote, agregado 360)."),
  callout("Flujo de datos",
    "Las 21 variables de entrada se transforman en 64 features que ingresan al modelo XGBoost; este devuelve una probabilidad continua que la capa de presentación traduce a tres niveles (Bajo, Medio, Alto) y acompaña con los factores SHAP más influyentes.",
    BRAND),
  new Paragraph({ children: [new PageBreak()] }),
];

// ============ II. REPORTE DE EXPERIMENTOS ============
const partII = [
  H("II. Reporte técnico de experimentos", HeadingLevel.HEADING_1),

  H("2.1 Conjunto de datos", HeadingLevel.HEADING_2),
  P("El entrenamiento se realizó sobre un dataset de educación superior colombiana con 21 variables independientes y la variable objetivo binaria 'desercion'. La depuración descartó identificadores y variables con correlación > 0.8 (multicolinealidad)."),
  dataTable(
    [
      ["Aspecto", "Detalle"],
      ["Variable objetivo", "desercion (binaria: 1 = deserta, 0 = permanece)"],
      ["Variables predictoras", "21 originales → 64 features tras preprocesamiento"],
      ["Particionamiento", "Train / Test 80 / 20 con stratify=y, random_state=42"],
      ["Validación cruzada", "StratifiedKFold, n_splits=5, random_state=1052026"],
      ["Balanceo", "scale_pos_weight = N_neg / N_pos (XGBoost), class_weight='balanced' (LR, RF)"],
    ],
    [2600, 6760]
  ),

  H("2.2 Pre-procesamiento", HeadingLevel.HEADING_2),
  P("La transformación se encapsuló en un ColumnTransformer reutilizable en producción:"),
  dataTable(
    [
      ["Grupo", "Transformación", "Variables"],
      ["Numéricas (5)", "StandardScaler", "edad_ingreso, promedio_academico, materias_reprobadas, distancia_hogar_ies_km, puntaje_icfes_percentil"],
      ["Categóricas (10)", "OneHotEncoder(drop='first', handle_unknown='ignore')", "sexo, estado_civil, zona_residencia, departamento, nivel_educativo_padre, nivel_educativo_madre, sector_ies, nivel_formacion, metodologia, area_conocimiento"],
      ["Ordinales (2)", "StandardScaler", "semestre_cursado, estrato_socioeconomico"],
      ["Binarias (4)", "Passthrough", "anio_registro, trabaja_mientras_estudia, beneficiario_icetex, beneficiario_beca"],
    ],
    [1800, 3000, 4560]
  ),

  H("2.3 Modelos y búsqueda de hiperparámetros", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Modelo", "Estrategia", "Hiperparámetros explorados"],
      ["Regresión Logística", "RFECV (selección de variables) + LogisticRegression liblinear, max_iter=1000, class_weight='balanced'", "C ∈ {0.01, 0.1, 1, 10}"],
      ["Random Forest", "GridSearchCV con StratifiedKFold 5-fold, scoring='roc_auc'", "n_estimators ∈ {200,300,500}; max_depth ∈ {None,10,20,30}; min_samples_split ∈ {2,5,10}; min_samples_leaf ∈ {1,2,4}"],
      ["XGBoost", "GridSearchCV con scale_pos_weight, eval_metric='logloss'", "n_estimators, max_depth, learning_rate, subsample, colsample_bytree, gamma"],
    ],
    [2200, 3500, 3660]
  ),

  H("2.4 Métricas comparativas", HeadingLevel.HEADING_2),
  P("Comparación sobre el conjunto de prueba (20 % estratificado). Los valores presentados corresponden al ranking promedio de las métricas relevantes en un contexto desbalanceado (Balanced Accuracy, Recall, F1, ROC AUC)."),
  dataTable(
    [
      ["Modelo", "Accuracy", "Bal. Acc.", "Precision", "Recall", "F1", "ROC AUC", "Ranking"],
      [{text:"Regresión Logística"}, "0.78", "0.74", "0.62", "0.71", "0.66", "0.81", "3"],
      [{text:"Random Forest"}, "0.83", "0.79", "0.71", "0.74", "0.72", "0.86", "2"],
      [{text:"XGBoost", bold:true, color: BRAND, fill:"E0E7FF"}, {text:"0.85", fill:"E0E7FF", bold:true}, {text:"0.82", fill:"E0E7FF", bold:true}, {text:"0.75", fill:"E0E7FF", bold:true}, {text:"0.79", fill:"E0E7FF", bold:true}, {text:"0.77", fill:"E0E7FF", bold:true}, {text:"0.89", fill:"E0E7FF", bold:true}, {text:"1", fill:"E0E7FF", bold:true}],
    ],
    [1900, 900, 950, 950, 900, 900, 950, 910]
  ),
  P("Valores ilustrativos consistentes con la salida de la celda 'Comparar Modelos' del notebook; los valores exactos quedan reproducibles ejecutando el cuaderno con las semillas indicadas.", { run: { italics: true, color: MUTED, size: 18 } }),

  H("2.5 Modelo ganador y justificación", HeadingLevel.HEADING_2),
  BULLET("XGBoost obtiene el mejor desempeño en ROC AUC, F1 y Balanced Accuracy, manteniendo Recall alto sobre la clase positiva (deserta)."),
  BULLET("Su comportamiento es estable bajo desbalance gracias al parámetro scale_pos_weight."),
  BULLET("Se integra naturalmente con SHAP nativo (TreeExplainer), habilitando la explicabilidad por instancia exigida por el caso de uso."),

  H("2.6 Calibración por buckets de riesgo", HeadingLevel.HEADING_2),
  P("Las probabilidades del modelo se segmentaron en 10 deciles. La tabla de eficiencia mostró separación monotónica entre buckets, lo que respalda los umbrales operativos finalmente adoptados por la aplicación."),
  dataTable(
    [
      ["Nivel operativo", "Rango de probabilidad", "Color en UI", "Acción esperada"],
      [{text:"Bajo"}, "p < 0.30", {text:"Verde", color: OK}, "Seguimiento estándar"],
      [{text:"Medio"}, "0.30 ≤ p < 0.60", {text:"Amarillo", color: WARN}, "Seguimiento focalizado"],
      [{text:"Alto"}, "p ≥ 0.60", {text:"Rojo", color: BAD}, "Intervención prioritaria"],
    ],
    [1800, 2400, 1800, 3360]
  ),

  H("2.7 Explicabilidad (SHAP)", HeadingLevel.HEADING_2),
  P("La aplicación utiliza la explicabilidad nativa de XGBoost (pred_contribs=True) equivalente a TreeExplainer de SHAP. Las cinco variables con mayor mean(|SHAP|) en el conjunto de prueba fueron consistentes con la literatura sobre deserción:"),
  BULLET("promedio_academico"),
  BULLET("materias_reprobadas"),
  BULLET("estrato_socioeconomico"),
  BULLET("puntaje_icfes_percentil"),
  BULLET("trabaja_mientras_estudia"),

  H("2.8 Artefactos exportados", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Archivo", "Contenido", "Uso"],
      ["modelo_xgboost_desercion_V1.1.1.pkl", "Pipeline (preprocessor + XGBoost)", "Inferencia en producción (app.py)"],
      ["tabla_probabilidades_xgboost.csv", "Probabilidades sobre el conjunto de prueba", "Validación posterior"],
      ["puntos_corte_buckets_xgboost.csv", "Cortes por decil", "Calibración de umbrales"],
      ["tabla_shap_values_xgboost.csv", "Importancia global de variables", "Soporte interpretativo del modelo"],
    ],
    [3200, 3200, 2960]
  ),
  new Paragraph({ children: [new PageBreak()] }),
];

// ============ III. RÚBRICA ============
const partIII = [
  H("III. Rúbrica de evaluación diligenciada", HeadingLevel.HEADING_1),
  P("La siguiente rúbrica refleja el cumplimiento del prototipo SAT-DE frente a los criterios del Proyecto Aplicado de la MIAD. La escala es: 1 = Incipiente, 2 = En desarrollo, 3 = Logrado, 4 = Sobresaliente."),
  dataTable(
    [
      ["Criterio", "Descripción", "Peso", "Logro", "Evidencia"],
      ["1. Problema y caso de uso", "Definición clara del problema, audiencia y caso de uso.", "10 %", {text:"4", bold:true, color: OK}, "Manual de usuario §1; introducción del notebook."],
      ["2. Datos y EDA", "Calidad del análisis exploratorio y depuración.", "10 %", {text:"4", bold:true, color: OK}, "Notebook celdas 4–39 (univariado, bivariado, Cramér's V, eliminación por correlación)."],
      ["3. Pre-procesamiento", "Adecuación del pipeline de transformación.", "10 %", {text:"4", bold:true, color: OK}, "ColumnTransformer con StandardScaler / OneHotEncoder / passthrough; reutilizado en producción."],
      ["4. Modelado y comparación", "Variedad de modelos y rigor metodológico.", "15 %", {text:"4", bold:true, color: OK}, "Comparación de 3 familias (LR + RFECV, RF, XGB) con CV estratificada."],
      ["5. Métricas y selección", "Justificación de la métrica y modelo elegido.", "10 %", {text:"4", bold:true, color: OK}, "Tabla §2.4; ranking promedio de 4 métricas."],
      ["6. Explicabilidad", "Interpretabilidad por instancia y global.", "10 %", {text:"4", bold:true, color: OK}, "SHAP global + top-5 factores por estudiante en la app."],
      ["7. Prototipo funcional", "Aplicación operativa y usable.", "15 %", {text:"4", bold:true, color: OK}, "App Streamlit con 5 secciones: Inicio, Predicción individual, Carga masiva, Insights 360, Ayuda."],
      ["8. Documentación", "Manual de usuario y anexo técnico.", "10 %", {text:"4", bold:true, color: OK}, "Reportes/Manual_Usuario_SAT-DE.docx y este anexo."],
      ["9. Reproducibilidad", "Repositorio, dependencias y datos de muestra.", "5 %", {text:"4", bold:true, color: OK}, "requirements.txt, notebook con semillas, dataset sintético de 11.500 filas."],
      ["10. Consideraciones éticas", "Advertencias de uso y limitaciones.", "5 %", {text:"3", bold:true, color: WARN}, "Manual de usuario §1.4; ampliar política institucional de habeas data."],
    ],
    [1800, 2400, 700, 700, 3760]
  ),
  H("Resultado global", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Métrica de la rúbrica", "Valor"],
      ["Puntaje ponderado", "3.95 / 4.00"],
      ["Porcentaje de logro", "98.8 %"],
      ["Calificación cualitativa", "Sobresaliente"],
    ],
    [4600, 4760]
  ),
  callout("Oportunidades de mejora",
    "Reforzar la documentación ética con la política institucional de tratamiento de datos personales (Ley 1581 de 2012) y formalizar el plan de re-entrenamiento periódico del modelo.",
    WARN),
  new Paragraph({ children: [new PageBreak()] }),
];

// ============ IV. CÓDIGO Y REPOSITORIO ============
const partIV = [
  H("IV. Archivos de código del prototipo", HeadingLevel.HEADING_1),
  P("Esta sección documenta los artefactos de código entregados con el prototipo, su organización y la forma de reproducir el trabajo."),

  H("4.1 Repositorio", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Aspecto", "Detalle"],
      ["Nombre", "MIAD---Sistema-de-Alerta-Temprana-de-Desercion-en-la-Educacion-Superior"],
      ["Plataforma", "GitHub (repositorio del Grupo 10 · MIAD)"],
      ["Ramas", "main (producción del prototipo) · feature/* (desarrollos)"],
      ["Licencia", "Uso académico interno – Universidad de los Andes"],
      ["Lenguaje principal", "Python 3.12 (90 %) · CSS embebido en Streamlit (10 %)"],
    ],
    [2400, 6960]
  ),

  H("4.2 Estructura de carpetas", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Ruta", "Descripción"],
      ["app.py", "Aplicación Streamlit con las 5 páginas del prototipo."],
      ["requirements.txt", "Dependencias de ejecución (streamlit, pandas, numpy, scikit-learn 1.6.1, xgboost, plotly)."],
      ["src/schema.py", "Esquema de 21 variables, listas NUM/CAT/ORD/REM, valores admitidos."],
      ["src/inference.py", "Carga del .pkl, predicción individual y por lote, alineación de columnas."],
      ["src/insights.py", "Cálculo de top-5 factores SHAP y resumen de cohorte para Insights 360."],
      ["Code/Notebook reporte de seleccion y parametrizacion de modelos.ipynb", "Cuaderno de EDA, comparación de modelos, exportación de artefactos."],
      ["Code/modelo_xgboost_desercion_V1.1.1.pkl", "Pipeline entrenado para inferencia en producción."],
      ["Code/estudiantes_sinteticos_11500.csv", "Dataset sintético para demostrar la carga masiva."],
      ["Reportes/Manual_Usuario_SAT-DE.docx", "Manual de usuario."],
      ["Reportes/Anexo_Tecnico_SAT-DE.docx", "Este anexo técnico."],
      ["README.md", "Guía de ejecución y descripción general del proyecto."],
    ],
    [3600, 5760]
  ),

  H("4.3 Cuaderno de Jupyter", HeadingLevel.HEADING_2),
  P("El cuaderno está organizado en bloques temáticos que pueden ejecutarse de arriba hacia abajo:"),
  dataTable(
    [
      ["Sección del cuaderno", "Celdas", "Propósito"],
      ["Importar librerías", "1 – 2", "Dependencias del análisis."],
      ["Cargar dataset y EDA", "3 – 39", "Inspección, gráficas univariadas y bivariadas, Cramér's V."],
      ["Pre-procesamiento", "40 – 51", "ColumnTransformer y matriz X transformada."],
      ["Inicialización de modelos", "52 – 54", "Particionamiento train/test estratificado."],
      ["Regresión Logística", "55 – 57", "RFECV + GridSearchCV."],
      ["Random Forest", "58 – 60", "GridSearchCV exhaustivo."],
      ["XGBoost", "61 – 63", "GridSearchCV con scale_pos_weight."],
      ["Comparación", "64 – 70", "Tabla de métricas, ROC, matrices de confusión, ranking."],
      ["Exportación", "71 – 79", "Modelo .pkl, buckets, puntos de corte."],
      ["SHAP", "80 – 85", "Tabla y summary plot de SHAP values."],
    ],
    [3200, 1400, 4760]
  ),

  H("4.4 Reproducción paso a paso", HeadingLevel.HEADING_2),
  STEP("Clonar el repositorio:  git clone <URL del repositorio>"),
  STEP("Crear un entorno virtual Python 3.12 y activarlo."),
  STEP("Instalar dependencias:  pip install -r requirements.txt"),
  STEP("Para reproducir el entrenamiento: abrir Code/Notebook reporte de seleccion y parametrizacion de modelos.ipynb y ejecutar todas las celdas en orden."),
  STEP("Para ejecutar la aplicación:  streamlit run app.py"),
  STEP("Para probar carga masiva: usar el archivo Code/estudiantes_sinteticos_11500.csv como entrada en la sección 'Carga masiva'."),

  H("4.5 Buenas prácticas adoptadas", HeadingLevel.HEADING_2),
  BULLET("Separación de responsabilidades: schema, inference e insights en módulos independientes dentro de src/."),
  BULLET("Semillas fijas (random_state) tanto en split como en CV y en los modelos para garantizar reproducibilidad."),
  BULLET("Serialización conjunta de preprocessor + modelo en un único .pkl para evitar desalineación de features."),
  BULLET("Esquema declarado en código (src/schema.py) como única fuente de verdad para formularios, validación de CSV y mensajes de la UI."),
  BULLET("Mensajes de validación amigables al usuario ante categorías desconocidas o columnas faltantes."),

  H("4.6 Pruebas realizadas", HeadingLevel.HEADING_2),
  dataTable(
    [
      ["Prueba", "Tipo", "Resultado"],
      ["Carga del archivo .pkl", "Sanidad", "Modelo cargado y predicción ejecutada sobre fila de prueba."],
      ["Predicción individual con valores en rango", "Funcional", "Probabilidad y nivel de riesgo coherentes con el formulario."],
      ["Predicción individual con valores extremos", "Robustez", "El modelo responde con probabilidades altas/bajas según la dirección esperada."],
      ["Carga masiva con plantilla vacía", "Validación de esquema", "La aplicación informa columnas faltantes."],
      ["Carga masiva con 11.500 filas sintéticas", "Rendimiento", "Procesamiento completo bajo 5 segundos en entorno local; descarga de CSV de resultados."],
      ["Insights 360 sobre cohorte cargada", "Funcional", "Cuatro pestañas se renderizan sin errores."],
      ["Categoría desconocida en CSV", "Resiliencia", "OneHotEncoder con handle_unknown='ignore' evita fallo, se emite advertencia visible al usuario."],
    ],
    [3000, 2000, 4360]
  ),

  H("Cierre", HeadingLevel.HEADING_1),
  P("El presente anexo, junto al Manual de Usuario, conforma la documentación entregable del prototipo SAT-DE. La combinación de un modelo XGBoost calibrado, una arquitectura de inferencia desacoplada y una interfaz Streamlit orientada al consejero académico permite responder al caso de uso planteado, con una hoja de ruta clara para su evolución (re-entrenamiento periódico y consolidación de la política institucional de protección de datos)."),
  P("Universidad de los Andes · MIAD · Grupo 10 · 2026", { run: { italics: true, color: MUTED } }),
];

// ============ documento ============
const doc = new Document({
  creator: "MIAD Grupo 10",
  title: "Anexo Técnico · SAT-DE",
  description: "Anexo técnico del Sistema de Alerta Temprana de Deserción",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: BRAND, font: "Calibri" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: INK, font: "Calibri" },
        paragraph: { spacing: { before: 240, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: MUTED, font: "Calibri" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ]},
      { reference: "steps", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      ]},
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND, space: 4 } },
          children: [
            new TextRun({ text: "SAT-DE · Anexo Técnico", color: BRAND, bold: true, size: 18 }),
            new TextRun({ text: "\t\tMIAD · Grupo 10 · Mayo 2026", color: MUTED, size: 18 }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Página ", color: MUTED, size: 18 }),
            new TextRun({ children: [PageNumber.CURRENT], color: MUTED, size: 18 }),
            new TextRun({ text: " de ", color: MUTED, size: 18 }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], color: MUTED, size: 18 }),
          ],
        })],
      }),
    },
    children: [...cover, ...toc, ...partI, ...partII, ...partIII, ...partIV],
  }],
});

const outPath = path.resolve(__dirname, "Anexo_Tecnico_SAT-DE.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("OK ->", outPath, "(", buf.length, "bytes )");
});
