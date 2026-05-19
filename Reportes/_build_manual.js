/* Genera Reportes/Manual_Usuario_SAT-DE.docx */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
  TableOfContents, Bookmark, InternalHyperlink, ExternalHyperlink,
} = require("docx");

// --------- helpers ----------
const BRAND = "1D4ED8";
const INK   = "0F172A";
const MUTED = "475569";
const LINE  = "E2E8F0";

const border = (color = LINE) => ({ style: BorderStyle.SINGLE, size: 6, color });
const cellBorders = {
  top: border(LINE), bottom: border(LINE), left: border(LINE), right: border(LINE),
};

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 },
  ...opts,
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

const CELL = (children, opts = {}) => new TableCell({
  borders: cellBorders,
  margins: { top: 90, bottom: 90, left: 140, right: 140 },
  width: { size: opts.width, type: WidthType.DXA },
  shading: opts.head
    ? { fill: BRAND, type: ShadingType.CLEAR, color: "auto" }
    : (opts.zebra ? { fill: "F8FAFC", type: ShadingType.CLEAR, color: "auto" } : undefined),
  children: Array.isArray(children) ? children : [children],
});

const TXT = (text, opts = {}) => new Paragraph({
  spacing: { after: 0 },
  children: [new TextRun({
    text,
    bold: !!opts.bold,
    color: opts.color,
    size: opts.size || 20, // 10pt
  })],
});

function buildTable(rows, columnWidths) {
  const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    rows: rows.map((cells, i) => new TableRow({
      tableHeader: i === 0,
      children: cells.map((c, j) => CELL(
        [TXT(c, i === 0 ? { bold: true, color: "FFFFFF" } : {})],
        { width: columnWidths[j], head: i === 0, zebra: i > 0 && i % 2 === 0 }
      )),
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
      shading: { fill: "F8FAFC", type: ShadingType.CLEAR, color: "auto" },
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

// --------- portada ----------
const cover = [
  new Paragraph({
    spacing: { before: 2400, after: 120 },
    children: [new TextRun({ text: "Manual de Usuario", size: 56, bold: true, color: INK })],
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text: "SAT-DE  ·  Sistema de Alerta Temprana de Deserción en Educación Superior",
      size: 32, bold: true, color: BRAND })],
  }),
  new Paragraph({
    spacing: { after: 480 },
    children: [new TextRun({ text: "Versión 1.0   ·   Mayo de 2026", size: 22, color: MUTED })],
  }),
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BRAND, space: 4 } },
    children: [new TextRun({ text: "" })],
  }),
  new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({ text: "Proyecto Aplicado en Analítica de Datos\n", size: 22, color: INK }),
    ],
  }),
  new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: "Maestría en Inteligencia Analítica de Datos (MIAD)", size: 22, color: INK })],
  }),
  new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: "Universidad de los Andes – Grupo 10", size: 22, color: INK })],
  }),
  new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: "Modelo subyacente: XGBoost v1.1.1", size: 22, color: MUTED })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// --------- TOC ----------
const toc = [
  H("Tabla de contenido", HeadingLevel.HEADING_1),
  new Paragraph({
    children: [new TableOfContents("Contenido", { hyperlink: true, headingStyleRange: "1-3" })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 1. Qué es y qué hace ----------
const sec1 = [
  H("1. ¿Qué es y qué hace SAT-DE?", HeadingLevel.HEADING_1),
  P("SAT-DE es una aplicación web de alerta temprana que estima la probabilidad de que un estudiante de educación superior deserte en el siguiente período académico. Su propósito es apoyar la consejería académica entregando, además de la probabilidad, los factores que más influyen en la decisión y una recomendación textual para el caso."),
  P("La aplicación expone tres capacidades principales:"),
  BULLET("Predicción individual a través de un formulario guiado con los datos del estudiante."),
  BULLET("Procesamiento por lote (CSV) para clasificar cohortes completas y descargar los resultados."),
  BULLET("Insights 360 a nivel de cohorte: distribución del riesgo y cortes por dimensiones demográficas, socioeconómicas y académicas."),

  H("1.1 Componentes y arquitectura", HeadingLevel.HEADING_2),
  P("SAT-DE se ejecuta como una aplicación Streamlit (Python) que carga en memoria un modelo XGBoost previamente entrenado (archivo .pkl). Los componentes son:"),
  buildTable(
    [
      ["Componente", "Tecnología", "Función"],
      ["Interfaz de usuario", "Streamlit + Plotly", "Formularios, visualizaciones y navegación."],
      ["Capa de inferencia", "scikit-learn + XGBoost", "Pre-procesamiento de variables y predicción."],
      ["Explicabilidad", "SHAP nativo de XGBoost", "Factores que influyen en cada predicción."],
      ["Modelo serializado", "modelo_xgboost_desercion_V1.1.1.pkl", "Pipeline pre-entrenado, 64 features."],
    ],
    [2200, 2800, 4360]
  ),

  H("1.2 Ventajas", HeadingLevel.HEADING_2),
  BULLET("Modelo XGBoost entrenado y validado sobre datos históricos colombianos de educación superior."),
  BULLET("Explicabilidad por estudiante: además del riesgo, se muestran los factores que aportan a la decisión."),
  BULLET("Tres modos de uso (individual, masivo, agregado) con una única interfaz unificada."),
  BULLET("Visualizaciones interactivas (Plotly) y exportación de resultados en CSV."),
  BULLET("Pensada para uso interno: no requiere conexión a sistemas externos ni envío de datos a terceros."),

  H("1.3 Limitaciones", HeadingLevel.HEADING_2),
  BULLET("El modelo es estadístico: produce una probabilidad estimada, no una decisión definitiva sobre el estudiante."),
  BULLET("Funciona sobre el esquema de 21 variables descrito en la sección 4; categorías o valores nuevos fuera del esquema son reemplazados o ignorados."),
  BULLET("La calidad de la predicción depende de la calidad y vigencia de los datos cargados."),
  BULLET("La aplicación se ejecuta como prototipo local; no incluye autenticación de usuarios ni control de auditoría."),
  BULLET("El modelo se entrenó con datos hasta cierto período; un re-entrenamiento periódico es necesario para mantener su desempeño."),

  H("1.4 Advertencias de uso responsable", HeadingLevel.HEADING_2),
  callout("Importante",
    "Las predicciones de SAT-DE son una ayuda diagnóstica para la consejería académica. NO deben usarse como única fuente para sanciones, expulsiones o cualquier decisión que afecte derechos del estudiante. La interpretación final corresponde al consejero o equipo académico responsable.",
    "DC2626"),
  BULLET("La información cargada puede contener datos personales: trate el archivo CSV conforme a la política de protección de datos de la institución."),
  BULLET("No comparta capturas que expongan identificadores junto con la probabilidad de deserción."),
  BULLET("Los umbrales de riesgo (Bajo / Medio / Alto) son institucionales y pueden requerir ajuste para su contexto."),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 2. Puesta en funcionamiento ----------
const sec2 = [
  H("2. Puesta en funcionamiento", HeadingLevel.HEADING_1),

  H("2.1 Conocimientos y habilidades del usuario", HeadingLevel.HEADING_2),
  P("SAT-DE está pensada para dos perfiles: el usuario final (consejero académico) y el responsable técnico de instalación."),
  buildTable(
    [
      ["Perfil", "Conocimientos requeridos"],
      ["Usuario final", "Manejo básico de navegador web; comprensión de los conceptos de probabilidad, riesgo y promedio académico; manejo de archivos CSV en Excel o similar."],
      ["Responsable técnico", "Línea de comandos (PowerShell o Bash); Python 3.12 o superior; instalación de paquetes con pip; redes locales y puertos."],
    ],
    [2500, 6860]
  ),

  H("2.2 Requisitos previos", HeadingLevel.HEADING_2),
  BULLET("Sistema operativo Windows 10/11, macOS o Linux."),
  BULLET("Python 3.12 o superior (recomendado) con pip habilitado."),
  BULLET("Acceso a internet la primera vez (para descargar dependencias)."),
  BULLET("Navegador moderno (Chrome, Edge, Firefox)."),
  BULLET("Aproximadamente 1 GB de espacio libre en disco para dependencias."),

  H("2.3 Descarga e instalación", HeadingLevel.HEADING_2),
  P("Los pasos siguientes se ejecutan UNA sola vez en el equipo donde se hospedará la aplicación."),
  STEP("Descargar el repositorio del proyecto desde el portal institucional o clonarlo con Git en una carpeta local."),
  STEP("Abrir una terminal (PowerShell en Windows) y ubicarse en la carpeta del proyecto."),
  STEP("Crear un entorno virtual de Python:  python -m venv .venv"),
  STEP("Activarlo:  .\\.venv\\Scripts\\Activate.ps1  (Windows) o  source .venv/bin/activate  (macOS/Linux)."),
  STEP("Instalar las dependencias:  pip install -r requirements.txt"),
  STEP("Verificar que el archivo Code/modelo_xgboost_desercion_V1.1.1.pkl está presente."),
  STEP("Lanzar la aplicación:  streamlit run app.py"),
  P("Streamlit abrirá automáticamente el navegador en la dirección http://localhost:8501."),

  H("2.4 Actualización del aplicativo", HeadingLevel.HEADING_2),
  STEP("Detener la aplicación con Ctrl + C en la terminal."),
  STEP("Descargar o sincronizar la nueva versión del repositorio (git pull o reemplazo manual)."),
  STEP("Reinstalar dependencias si requirements.txt cambió:  pip install -r requirements.txt"),
  STEP("Si el archivo .pkl del modelo fue reemplazado, no se requiere ninguna acción adicional: la aplicación detecta el archivo al iniciar."),
  STEP("Volver a ejecutar  streamlit run app.py"),

  H("2.5 Acceso a la versión publicada", HeadingLevel.HEADING_2),
  P("Cuando la aplicación se publica en un servidor (por ejemplo Azure App Service), el usuario final no necesita instalar nada: basta con abrir la URL que proporcione la institución desde un navegador moderno."),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 3. Casos de uso ----------
const sec3 = [
  H("3. Casos de uso y guías paso a paso", HeadingLevel.HEADING_1),
  P("SAT-DE soporta tres casos de uso principales. Todos se realizan desde la barra lateral de navegación."),
  buildTable(
    [
      ["Caso de uso", "Descripción", "Sección de la app"],
      ["Evaluar a un estudiante", "Cuando se desea analizar un caso individual con su perfil completo.", "Predicción individual"],
      ["Evaluar una cohorte", "Cuando hay un grupo de estudiantes en un archivo CSV.", "Carga masiva"],
      ["Explorar la cohorte", "Cuando ya se procesó un CSV y se desea analizarlo por dimensiones.", "Insights 360"],
    ],
    [2400, 4960, 2000]
  ),

  H("3.1 Caso 1 · Predicción individual", HeadingLevel.HEADING_2),
  P("Indicado para consejería caso a caso. El consejero ingresa los datos del estudiante y obtiene la probabilidad estimada, el nivel de riesgo y los factores explicativos."),
  STEP("En la barra lateral, seleccionar 👤 Predicción individual."),
  STEP("Completar las tres secciones del formulario: Demográfico, Socioeconómico y Académico. Todos los campos son requeridos."),
  STEP("Hacer clic en el botón Calcular riesgo (botón principal al final del formulario)."),
  STEP("La aplicación mostrará:  (i) la insignia de riesgo (Bajo, Medio o Alto);  (ii) la probabilidad de deserción como porcentaje;  (iii) la barra Deserción / Permanencia;  (iv) los cinco factores más influyentes con su contribución relativa y dirección (Riesgo o Protección);  (v) una recomendación textual."),
  STEP("Modificar valores del formulario y volver a calcular para realizar análisis de sensibilidad."),
  callout("Interpretación",
    "Una variable marcada como 'Riesgo' aporta a aumentar la probabilidad de deserción; 'Protección' aporta a disminuirla. El porcentaje refleja la importancia relativa entre los cinco factores mostrados.",
    BRAND),

  H("3.2 Caso 2 · Carga masiva por CSV", HeadingLevel.HEADING_2),
  P("Indicado para clasificar una cohorte completa. La aplicación valida el archivo, calcula la predicción de cada fila y permite descargar los resultados."),
  STEP("En la barra lateral, seleccionar 📥 Carga masiva."),
  STEP("Abrir el expander 'Ver columnas esperadas / descargar plantilla' para conocer el esquema exacto."),
  STEP("Descargar la plantilla CSV vacía si se desea trabajar a partir de ella, o usar el archivo de ejemplo Code/estudiantes_sinteticos_11500.csv."),
  STEP("Diligenciar el CSV respetando los nombres de columnas, los valores admitidos para variables categóricas y los rangos numéricos (ver sección 4)."),
  STEP("Cargar el archivo arrastrándolo al cuadro 'Archivo CSV' o usando el botón Browse files."),
  STEP("Revisar las advertencias amarillas (categorías desconocidas) y los errores rojos (esquema inválido). En caso de error, corregir el CSV y volver a cargar."),
  STEP("La aplicación procesa el archivo y muestra: KPIs de la cohorte, tabla de resultados filtrable por nivel de riesgo y botón de descarga del CSV con predicciones."),
  STEP("Filtrar por Alto / Medio / Bajo desde el selector multiselect para enfocar el análisis."),
  STEP("Hacer clic en ⬇️ Descargar resultados (CSV) para obtener el archivo con las columnas originales más prob_desercion, riesgo y prediccion."),

  H("3.3 Caso 3 · Insights 360 de la cohorte", HeadingLevel.HEADING_2),
  P("Indicado para explorar la cohorte ya cargada. Si aún no se ha cargado un CSV, la sección lo solicitará explícitamente."),
  STEP("Cargar primero la cohorte usando 📥 Carga masiva (Caso 2)."),
  STEP("Seleccionar 📊 Insights 360 en la barra lateral."),
  STEP("Revisar los KPIs superiores: total de estudiantes, número y porcentaje en cada nivel de riesgo."),
  STEP("Navegar entre las cuatro pestañas:"),
  BULLET("Distribución: histograma de la probabilidad y donut con la composición por nivel de riesgo.", 1),
  BULLET("Demográfico / Socioeconómico: riesgo medio por estrato, por beca y distribución por zona de residencia.", 1),
  BULLET("Académico: riesgo medio por área de conocimiento y diagrama de dispersión entre promedio académico y materias reprobadas.", 1),
  BULLET("Detalle: tabla ordenada de mayor a menor probabilidad, lista para análisis caso por caso.", 1),
  STEP("Cada gráfico es interactivo: hover para detalles, doble clic en la leyenda para aislar series, y selección rectangular en los scatter para hacer zoom."),

  H("3.4 Visualización y descarga", HeadingLevel.HEADING_2),
  P("Todas las visualizaciones tienen un menú flotante (esquina superior derecha del gráfico) que permite descargarlas como imagen PNG. La tabla detallada también puede copiarse al portapapeles."),
  buildTable(
    [
      ["Acción", "Cómo realizarla"],
      ["Descargar un gráfico", "Pasar el cursor sobre el gráfico → ícono de cámara en la barra superior del Plotly."],
      ["Descargar la tabla", "En Carga masiva, usar el botón Descargar resultados (CSV)."],
      ["Compartir un análisis", "Capturar pantalla del navegador o exportar el CSV con las predicciones."],
    ],
    [3000, 6360]
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 4. Anexos ----------
const sec4 = [
  H("4. Anexo · Esquema de variables", HeadingLevel.HEADING_1),
  P("Las 21 variables esperadas por SAT-DE, agrupadas por dimensión. Para el detalle exhaustivo de categorías, consulte la pestaña Ayuda dentro de la aplicación."),

  H("4.1 Variables académicas", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Variable", "Tipo", "Rango / Valores"],
      ["edad_ingreso", "Numérica", "Años, típicamente 15 – 45"],
      ["promedio_academico", "Numérica", "Escala 0.0 – 5.0"],
      ["materias_reprobadas", "Numérica", "Conteo acumulado, ≥ 0"],
      ["puntaje_icfes_percentil", "Numérica", "Percentil 1 – 100"],
      ["semestre_cursado", "Ordinal", "1 – 12+"],
      ["sector_ies", "Categórica", "Privado, Publico"],
      ["nivel_formacion", "Categórica", "Tecnico profesional, Tecnologico, Universitario"],
      ["metodologia", "Categórica", "Presencial, Distancia tradicional, Virtual"],
      ["area_conocimiento", "Categórica", "8 áreas (Ingeniería, Salud, Educación, etc.)"],
      ["anio_registro", "Numérica", "Año del registro académico"],
    ],
    [3000, 1800, 4560]
  ),

  H("4.2 Variables demográficas y socioeconómicas", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Variable", "Tipo", "Rango / Valores"],
      ["sexo", "Categórica", "Hombre, Mujer"],
      ["estado_civil", "Categórica", "Soltero, Union libre, Casado, Otro"],
      ["zona_residencia", "Categórica", "Urbana, Rural"],
      ["departamento", "Categórica", "27 departamentos de Colombia"],
      ["distancia_hogar_ies_km", "Numérica", "Kilómetros, ≥ 0"],
      ["estrato_socioeconomico", "Ordinal", "1 – 6"],
      ["nivel_educativo_padre", "Categórica", "Ninguno → Posgrado"],
      ["nivel_educativo_madre", "Categórica", "Ninguno → Posgrado"],
      ["trabaja_mientras_estudia", "Binaria", "0 (No) / 1 (Sí)"],
      ["beneficiario_icetex", "Binaria", "0 / 1"],
      ["beneficiario_beca", "Binaria", "0 / 1"],
    ],
    [3000, 1800, 4560]
  ),

  H("4.3 Umbrales de riesgo", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Nivel", "Probabilidad", "Decisión sugerida"],
      ["Bajo", "< 30 %", "Seguimiento estándar"],
      ["Medio", "30 % – 60 %", "Seguimiento focalizado"],
      ["Alto", "≥ 60 %", "Intervención prioritaria"],
    ],
    [2200, 2800, 4360]
  ),

  H("5. Soporte y solución de problemas", HeadingLevel.HEADING_1),
  buildTable(
    [
      ["Síntoma", "Causa probable", "Acción sugerida"],
      ["La página no abre en localhost:8501", "Streamlit no quedó en ejecución.", "Revisar la terminal y volver a ejecutar streamlit run app.py."],
      ["Error 'Faltan columnas'", "El CSV no respeta el esquema.", "Descargar la plantilla y volver a cargar."],
      ["Advertencia 'Valor desconocido en categoría'", "Una fila usa un valor fuera del catálogo.", "Corregir el valor según la sección 4."],
      ["La predicción cambia mucho al editar un campo", "Variable de alto impacto (p.ej. materias reprobadas).", "Es comportamiento esperado del modelo; validar el valor real del estudiante."],
      ["El archivo .pkl no carga", "Pickle ausente o ruta incorrecta.", "Verificar que Code/modelo_xgboost_desercion_V1.1.1.pkl existe."],
    ],
    [2400, 3300, 3660]
  ),

  H("Contacto", HeadingLevel.HEADING_2),
  P("Equipo Grupo 10 · Maestría en Inteligencia Analítica de Datos (MIAD) · Universidad de los Andes. Para reportar incidencias, contactar al líder técnico del proyecto en el canal institucional asignado."),
];

// --------- documento ----------
const doc = new Document({
  creator: "MIAD Grupo 10",
  title: "Manual de Usuario · SAT-DE",
  description: "Manual de usuario del Sistema de Alerta Temprana de Deserción en Educación Superior",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } }, // 11pt
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
        size: { width: 12240, height: 15840 }, // US Letter
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND, space: 4 } },
          children: [
            new TextRun({ text: "SAT-DE  ·  Manual de Usuario", color: BRAND, bold: true, size: 18 }),
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
    children: [...cover, ...toc, ...sec1, ...sec2, ...sec3, ...sec4],
  }],
});

const outPath = path.resolve(__dirname, "..", "Reportes", "Manual_Usuario_SAT-DE.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("OK ->", outPath, "(", buf.length, "bytes )");
});
