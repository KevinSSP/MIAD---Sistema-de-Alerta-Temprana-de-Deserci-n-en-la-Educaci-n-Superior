/* Genera Reportes/Manual_Usuario_SAT-DE.docx (v2 - mayo 2026) */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
  TableOfContents, ExternalHyperlink,
} = require("docx");

// --------- helpers ----------
const BRAND   = "1D4ED8";
const NAVY    = "1E3A8A";
const SKY     = "0EA5E9";
const INK     = "0F172A";
const MUTED   = "475569";
const LINE    = "E2E8F0";
const OK      = "16A34A";
const WARN    = "EAB308";
const BAD     = "DC2626";
const CRIT    = "B91C1C";
const NORMAL  = "0EA5E9";

const border = (color = LINE, size = 6) => ({ style: BorderStyle.SINGLE, size, color });
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

const CODE = (text) => new Paragraph({
  spacing: { after: 100 },
  shading: { fill: "0F172A", type: ShadingType.CLEAR, color: "auto" },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: SKY } },
  children: [new TextRun({ text, font: "Consolas", size: 18, color: "E2E8F0" })],
});

const LINK = (text, url) => new Paragraph({
  spacing: { after: 120 },
  children: [
    new ExternalHyperlink({
      link: url,
      children: [new TextRun({ text, style: "Hyperlink", color: BRAND, underline: { type: "single", color: BRAND } })],
    }),
  ],
});

const CELL = (children, opts = {}) => new TableCell({
  borders: cellBorders,
  margins: { top: 90, bottom: 90, left: 140, right: 140 },
  width: { size: opts.width, type: WidthType.DXA },
  shading: opts.head
    ? { fill: BRAND, type: ShadingType.CLEAR, color: "auto" }
    : (opts.fill
        ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" }
        : (opts.zebra ? { fill: "F8FAFC", type: ShadingType.CLEAR, color: "auto" } : undefined)),
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

function buildTable(rows, columnWidths, opts = {}) {
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
const APP_URL = "https://miadsatde-h9fndjhfg7esaeet.centralus-01.azurewebsites.net/";

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
    children: [new TextRun({ text: "Versión 2.0   ·   Mayo de 2026", size: 22, color: MUTED })],
  }),
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BRAND, space: 4 } },
    children: [new TextRun({ text: "" })],
  }),
  new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text: "Proyecto Aplicado en Analítica de Datos", size: 22, color: INK })],
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
    spacing: { after: 240 },
    children: [new TextRun({ text: "Modelo subyacente: XGBoost v1.1.1 · Clasificación de riesgo en 5 categorías", size: 22, color: MUTED })],
  }),
  callout(
    "Aplicación desplegada (acceso público)",
    "La plataforma está publicada en Azure App Service. Cualquier usuario con la URL puede acceder directamente desde un navegador moderno, sin instalación local: " + APP_URL,
    OK
  ),
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
  P("SAT-DE es una aplicación web de alerta temprana que estima la probabilidad de que un estudiante de educación superior deserte en el siguiente período académico. Apoya la consejería académica entregando, además de la probabilidad, los factores que más influyen en la decisión y una recomendación textual."),
  P("La aplicación expone cinco funcionalidades organizadas en páginas independientes:"),
  BULLET("Inicio: explicación rápida del sistema y leyenda de los cinco niveles de riesgo."),
  BULLET("Predicción individual: formulario guiado para evaluar un estudiante específico."),
  BULLET("Carga masiva: clasificación por lote desde un archivo CSV."),
  BULLET("Insights 360: visión agregada de una cohorte cargada."),
  BULLET("Ayuda: diccionario de variables y tabla de umbrales de riesgo."),

  H("1.1 Componentes y arquitectura", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Componente", "Tecnología", "Función"],
      ["Interfaz de usuario", "Streamlit + Plotly", "Formularios, visualizaciones interactivas y navegación."],
      ["Capa de inferencia", "scikit-learn + XGBoost", "Pre-procesamiento de variables y predicción."],
      ["Explicabilidad", "Importancia agregada de XGBoost", "Factores que influyen en cada predicción individual."],
      ["Modelo serializado", "modelo_xgboost_desercion_V1.1.1.pkl", "Pipeline pre-entrenado (64 features esperadas)."],
      ["Hosting", "Azure App Service (Linux, Python 3.12)", "Publicación pública en " + APP_URL],
    ],
    [2200, 2800, 4360]
  ),

  H("1.2 Ventajas", HeadingLevel.HEADING_2),
  BULLET("Modelo XGBoost entrenado y validado sobre datos colombianos de educación superior (SPADIES + DANE)."),
  BULLET("Clasificación de riesgo en 5 categorías (Crítico, Alto, Medio, Normal y Bajo) calibradas con la Tabla de Eficiencia Operativa Real."),
  BULLET("Explicabilidad por estudiante: además del riesgo se muestran los factores que aportan a la decisión."),
  BULLET("Cinco modos de uso unificados en una sola aplicación web."),
  BULLET("Visualizaciones interactivas y exportación de resultados en CSV."),
  BULLET("Acceso público a través de Azure App Service: cualquier usuario con la URL puede usar la herramienta sin instalación local."),

  H("1.3 Limitaciones", HeadingLevel.HEADING_2),
  BULLET("El modelo es estadístico: produce una probabilidad estimada, no una decisión definitiva sobre el estudiante."),
  BULLET("Funciona sobre el esquema de 21 variables descrito en el anexo; categorías o valores nuevos son ignorados por el preprocesador."),
  BULLET("La calidad de la predicción depende de la calidad y vigencia de los datos cargados."),
  BULLET("El despliegue actual no incluye autenticación de usuarios: cualquiera con la URL accede; se recomienda restringir vía red corporativa o agregar autenticación si se va a usar productivamente."),
  BULLET("El modelo se entrenó con datos hasta cierto período; un re-entrenamiento periódico es necesario para mantener su desempeño."),

  H("1.4 Advertencias de uso responsable", HeadingLevel.HEADING_2),
  callout("Importante",
    "Las predicciones de SAT-DE son una ayuda diagnóstica para la consejería académica. NO deben usarse como única fuente para sanciones, expulsiones o cualquier decisión que afecte derechos del estudiante. La interpretación final corresponde al consejero o equipo académico responsable.",
    BAD),
  BULLET("La información cargada puede contener datos personales: trate el archivo CSV conforme a la política de protección de datos de la institución."),
  BULLET("No comparta capturas que expongan identificadores junto con la probabilidad de deserción."),
  BULLET("Los umbrales de las 5 bandas de riesgo provienen de un análisis empírico (Tabla de Eficiencia Operativa Real) y pueden requerir ajuste para contextos diferentes."),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 2. Acceso ----------
const sec2 = [
  H("2. Acceso a la aplicación", HeadingLevel.HEADING_1),

  H("2.1 Aplicación publicada (forma recomendada)", HeadingLevel.HEADING_2),
  P("SAT-DE está desplegada en Azure App Service y es accesible desde cualquier navegador moderno (Chrome, Edge, Firefox, Safari) sin necesidad de instalar Python ni dependencias."),
  callout(
    "URL pública",
    APP_URL,
    OK
  ),
  LINK(APP_URL, APP_URL),
  STEP("Abrir el navegador y pegar la URL anterior en la barra de direcciones."),
  STEP("Esperar a que cargue la página (la primera carga puede tomar 20–40 segundos si el servidor estaba en reposo)."),
  STEP("Usar la barra lateral izquierda para navegar entre las cinco páginas (Inicio, Predicción individual, Carga masiva, Insights 360, Ayuda)."),
  P("No se requiere usuario ni contraseña en esta versión. Si la institución decide restringir el acceso, deberá aplicarse a nivel de red o agregar autenticación en Azure (ver sección 4.5)."),

  H("2.2 Ejecución local (desarrollo y pruebas)", HeadingLevel.HEADING_2),
  P("Esta opción aplica solo al equipo técnico o al usuario que desee ejecutar el aplicativo sin conexión a Azure."),
  P("Requisitos previos:"),
  BULLET("Sistema operativo Windows 10/11, macOS o Linux."),
  BULLET("Python 3.12 o superior con pip habilitado."),
  BULLET("Aproximadamente 1 GB de espacio libre para dependencias."),
  BULLET("Navegador moderno."),
  STEP("Clonar o descargar el repositorio del proyecto."),
  STEP("Abrir una terminal y ubicarse en la carpeta raíz."),
  CODE("python -m venv .venv"),
  CODE(".\\.venv\\Scripts\\Activate.ps1     # Windows PowerShell"),
  CODE("source .venv/bin/activate          # macOS / Linux"),
  CODE("pip install -r requirements.txt"),
  STEP("Verificar que existe el archivo Code/modelo_xgboost_desercion_V1.1.1.pkl."),
  CODE("streamlit run app.py"),
  P("Streamlit abrirá automáticamente el navegador en http://localhost:8501."),

  H("2.3 Perfiles de usuario", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Perfil", "Necesita instalar algo?", "Forma de acceso"],
      ["Consejero / decano / coordinador", "No", "Abre la URL pública (sección 2.1)."],
      ["Responsable técnico", "Sí (Python + dependencias)", "Ejecuta localmente (sección 2.2) y publica en Azure (sección 4)."],
      ["Equipo TI institucional", "Solo si va a re-desplegar", "Sigue la guía de despliegue (sección 4)."],
    ],
    [3000, 2400, 3960]
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 3. Casos de uso ----------
const sec3 = [
  H("3. Guía de uso por página", HeadingLevel.HEADING_1),
  P("La barra lateral izquierda muestra las cinco páginas disponibles. A continuación se explica cada una con su propósito, los pasos y los resultados esperados."),

  H("3.1 Página Inicio", HeadingLevel.HEADING_2),
  P("Propósito: presentar la aplicación y la leyenda con los cinco niveles de riesgo manejados por la plataforma. Es el punto de entrada por defecto."),
  P("Qué se muestra:"),
  BULLET("Encabezado corporativo con la fecha del día y la versión del modelo."),
  BULLET("Tres tarjetas resumen de las funcionalidades (Predicción individual, Carga masiva, Insights 360)."),
  BULLET("Leyenda con las cinco categorías de riesgo y sus rangos de probabilidad."),
  P("Cuándo usarla: como pantalla de bienvenida, para recordar los rangos exactos antes de interpretar resultados."),

  H("3.2 Página Predicción individual", HeadingLevel.HEADING_2),
  P("Indicada para consejería caso a caso. El consejero ingresa los datos del estudiante y obtiene la probabilidad estimada, el nivel de riesgo, la decisión sugerida y los factores explicativos."),
  STEP("En la barra lateral, seleccionar 👤 Predicción individual."),
  STEP("Completar las tres secciones del formulario: Demográfico, Socioeconómico y Académico. Todos los campos son requeridos."),
  STEP("Hacer clic en el botón Calcular riesgo (botón principal al final del formulario)."),
  STEP("La aplicación mostrará:"),
  BULLET("La insignia de riesgo en uno de los 5 niveles (Crítico, Alto, Medio, Normal o Bajo).", 1),
  BULLET("La probabilidad de deserción como porcentaje.", 1),
  BULLET("La Decisión sugerida correspondiente a la banda (ver sección 3.6).", 1),
  BULLET("Barra horizontal Deserción / Permanencia.", 1),
  BULLET("Top-5 de factores influyentes con contribución relativa y dirección (Riesgo o Protección).", 1),
  BULLET("Recomendación textual coherente con la banda de riesgo.", 1),
  STEP("Modificar valores y volver a calcular para realizar análisis de sensibilidad sobre un mismo estudiante."),
  callout(
    "Interpretación de factores",
    "Una variable marcada como 'Riesgo' aporta a aumentar la probabilidad de deserción; 'Protección' aporta a disminuirla. El porcentaje refleja la importancia relativa entre los cinco factores mostrados, no la magnitud absoluta del efecto.",
    BRAND
  ),

  H("3.3 Página Carga masiva", HeadingLevel.HEADING_2),
  P("Indicada para clasificar una cohorte completa. La aplicación valida el archivo, calcula la predicción de cada fila y permite descargar los resultados."),
  STEP("En la barra lateral, seleccionar 📥 Carga masiva."),
  STEP("Abrir el expander 'Ver columnas esperadas / descargar plantilla' para conocer el esquema exacto."),
  STEP("Descargar la plantilla CSV vacía si se desea trabajar a partir de ella, o usar un dataset propio que respete las 21 variables."),
  STEP("Diligenciar el CSV respetando los nombres de columnas, los valores admitidos para variables categóricas y los rangos numéricos (ver anexo)."),
  STEP("Cargar el archivo arrastrándolo al cuadro 'Archivo CSV' o usando el botón Browse files."),
  STEP("Revisar advertencias amarillas (categorías desconocidas) o errores rojos (esquema inválido). Si aparecen errores, corregir el CSV y volver a cargar."),
  STEP("La aplicación muestra cuatro KPIs: total de estudiantes, riesgo crítico, % prioritarios (Crítico + Alto) y probabilidad media."),
  STEP("Filtrar la tabla por una o varias bandas de riesgo desde el selector multiselect."),
  STEP("Hacer clic en ⬇️ Descargar resultados (CSV) para obtener el archivo con las columnas originales más prob_desercion, riesgo y prediccion."),
  callout(
    "Persistencia entre páginas",
    "La cohorte cargada queda en memoria de la sesión: al pasar a Insights 360 no es necesario volver a cargar el CSV. Si se cierra la pestaña del navegador o se reinicia el servidor, se pierde y debe cargarse nuevamente.",
    SKY
  ),

  H("3.4 Página Insights 360", HeadingLevel.HEADING_2),
  P("Indicada para explorar la cohorte ya cargada por dimensiones agregadas. Si aún no se ha cargado un CSV, la sección lo solicitará explícitamente."),
  STEP("Cargar primero la cohorte usando 📥 Carga masiva (paso 3.3)."),
  STEP("Seleccionar 📊 Insights 360 en la barra lateral."),
  STEP("Revisar los cinco KPIs superiores: total, Crítico, Alto, Medio y % prioritarios."),
  STEP("Navegar entre las cuatro pestañas:"),
  BULLET("Distribución: histograma de la probabilidad coloreado por banda de riesgo y donut con la composición porcentual de las 5 categorías.", 1),
  BULLET("Demográfico / Socioeconómico: riesgo medio por estrato, por beca y distribución por zona de residencia.", 1),
  BULLET("Académico: riesgo medio por área de conocimiento y diagrama de dispersión entre promedio académico y materias reprobadas (coloreado por banda).", 1),
  BULLET("Detalle: tabla ordenada de mayor a menor probabilidad, lista para análisis caso por caso.", 1),
  STEP("Cada gráfico es interactivo: hover para detalles, doble clic en la leyenda para aislar series, y selección rectangular en los scatter para hacer zoom."),
  STEP("Descargar cualquier gráfico como PNG desde el ícono de cámara en la barra superior del gráfico (Plotly toolbar)."),

  H("3.5 Página Ayuda", HeadingLevel.HEADING_2),
  P("Indicada para consultar el diccionario de variables y los umbrales aplicados por la plataforma."),
  BULLET("Diccionario de variables: tabla con las 21 variables, su etiqueta legible, su tipo (categórica, binaria, ordinal o numérica) y los valores admitidos."),
  BULLET("Umbrales de riesgo: cinco tarjetas con los rangos exactos de cada banda y la acción sugerida, seguidas de una tabla con probabilidad mínima/máxima, % de población esperado, tasa de deserción interna y acción."),
  BULLET("Pie de página con la versión del modelo (XGBoost v1.1.1) y el equipo responsable."),

  H("3.6 Tabla de decisiones y umbrales", HeadingLevel.HEADING_2),
  P("La aplicación clasifica cada estudiante en una de las cinco bandas siguientes según la probabilidad estimada de deserción. La 'Decisión sugerida' es la acción recomendada por defecto y aparece tanto en la página individual como en la recomendación textual."),
  buildTable(
    [
      ["Banda", "Rango de probabilidad", "Decisión sugerida", "% población esperado", "Tasa deserción interna"],
      ["Bajo",    "< 30.98 %",              "Permanencia",              "53.07 %", "6.42 %"],
      ["Normal",  "30.98 % – 51.33 %",      "Seguimiento",              "13.00 %", "23.14 %"],
      ["Medio",   "51.33 % – 69.27 %",      "Plan de acompañamiento",   "8.54 %",  "38.59 %"],
      ["Alto",    "69.27 % – 84.87 %",      "Intervención focalizada",  "8.89 %",  "57.93 %"],
      ["Crítico", "≥ 84.87 %",              "Intervención inmediata",   "16.50 %", "83.97 %"],
    ],
    [1400, 2400, 2400, 1580, 1580]
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 4. Despliegue ----------
const sec4 = [
  H("4. Despliegue de la aplicación", HeadingLevel.HEADING_1),
  P("Esta sección describe cómo publicar SAT-DE en Azure App Service, que es la modalidad utilizada actualmente en producción. La aplicación corre como un proceso Streamlit dentro de un App Service Linux con Python 3.12."),

  H("4.1 Resumen de la arquitectura desplegada", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Recurso", "Descripción"],
      ["App Service Plan", "Plan Linux (recomendado B1 o superior) en la región Central US."],
      ["App Service", "miadsatde – ejecuta Python 3.12 con Streamlit."],
      ["Runtime", "Python 3.12 (built-in en App Service)."],
      ["Comando de inicio", "python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0"],
      ["Almacenamiento", "El .pkl del modelo viaja con el repositorio (Code/modelo_xgboost_desercion_V1.1.1.pkl)."],
      ["URL pública", APP_URL],
    ],
    [2400, 6960]
  ),

  H("4.2 Pre-requisitos para desplegar", HeadingLevel.HEADING_2),
  BULLET("Cuenta de Azure con permisos de Contributor sobre la suscripción objetivo."),
  BULLET("Azure CLI instalado (az --version ≥ 2.60) o acceso al Portal de Azure."),
  BULLET("Git instalado localmente."),
  BULLET("Repositorio del proyecto con los archivos requirements.txt, app.py, src/ y Code/modelo_xgboost_desercion_V1.1.1.pkl."),

  H("4.3 Despliegue paso a paso con Azure CLI", HeadingLevel.HEADING_2),
  STEP("Autenticarse en Azure desde la terminal."),
  CODE("az login"),
  STEP("Definir variables de entorno locales para reutilizarlas (ajustar nombres si ya existen los recursos)."),
  CODE("$rg=\"rg-miad-satde\"; $loc=\"centralus\"; $plan=\"plan-miad-satde\"; $app=\"miadsatde\""),
  STEP("Crear el grupo de recursos."),
  CODE("az group create -n $rg -l $loc"),
  STEP("Crear el plan de App Service en Linux (sku B1)."),
  CODE("az appservice plan create -g $rg -n $plan --is-linux --sku B1"),
  STEP("Crear el Web App con runtime Python 3.12."),
  CODE("az webapp create -g $rg -p $plan -n $app --runtime \"PYTHON:3.12\""),
  STEP("Configurar el comando de inicio para que Streamlit escuche en el puerto 8000 expuesto por App Service."),
  CODE("az webapp config set -g $rg -n $app --startup-file \"python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0\""),
  STEP("Habilitar el build automático de dependencias al desplegar (App Service ejecuta pip install)."),
  CODE("az webapp config appsettings set -g $rg -n $app --settings SCM_DO_BUILD_DURING_DEPLOYMENT=true WEBSITES_PORT=8000"),
  STEP("Empaquetar el repositorio en un archivo zip y subirlo."),
  CODE("Compress-Archive -Path .\\* -DestinationPath app.zip -Force"),
  CODE("az webapp deploy -g $rg -n $app --src-path app.zip --type zip"),
  STEP("Esperar 2–4 minutos a que App Service instale dependencias y arranque Streamlit; luego abrir la URL pública."),
  CODE("Start-Process \"" + APP_URL + "\""),

  H("4.4 Despliegue desde Portal de Azure (alternativa visual)", HeadingLevel.HEADING_2),
  STEP("Crear un App Service: Portal → Create resource → Web App. Seleccionar Publish=Code, Runtime=Python 3.12, OS=Linux, Region=Central US."),
  STEP("En el plan, seleccionar B1 o superior."),
  STEP("Una vez creado, ir a Configuration → General Settings → Startup Command y pegar:  python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0"),
  STEP("En Configuration → Application Settings, agregar WEBSITES_PORT=8000 y SCM_DO_BUILD_DURING_DEPLOYMENT=true."),
  STEP("En Deployment Center, conectar el repositorio (GitHub, Azure Repos o ZIP) y disparar el deploy."),
  STEP("Esperar a que el contenedor se reinicie y validar la URL pública."),

  H("4.5 Seguridad y endurecimiento (recomendado)", HeadingLevel.HEADING_2),
  BULLET("Activar HTTPS Only en TLS/SSL settings."),
  BULLET("Para limitar el acceso, configurar Access Restrictions a un rango de IP corporativo o integrar con Azure AD (App Service Authentication / EasyAuth)."),
  BULLET("Habilitar Application Insights para monitorear errores y rendimiento."),
  BULLET("Configurar reglas de Auto-scale si la concurrencia esperada supera ~20 usuarios simultáneos (escalado horizontal del plan)."),
  BULLET("Cifrar variables sensibles vía Key Vault y referenciarlas desde App Settings si en el futuro se agregan credenciales o conexiones a bases externas."),

  H("4.6 Actualización del aplicativo desplegado", HeadingLevel.HEADING_2),
  STEP("Sincronizar localmente la nueva versión del repositorio (git pull o reemplazo manual)."),
  STEP("Volver a empaquetar y subir el zip."),
  CODE("Compress-Archive -Path .\\* -DestinationPath app.zip -Force"),
  CODE("az webapp deploy -g $rg -n $app --src-path app.zip --type zip"),
  STEP("Reiniciar el servicio (opcional, si el deploy no lo reinicia automáticamente)."),
  CODE("az webapp restart -g $rg -n $app"),
  STEP("Validar la URL pública. La primera carga tras un deploy puede tardar 30–60 segundos."),

  H("4.7 Verificación post-despliegue", HeadingLevel.HEADING_2),
  buildTable(
    [
      ["Prueba", "Resultado esperado"],
      ["Abrir " + APP_URL, "Carga la página de Inicio con el encabezado azul y la leyenda de 5 niveles."],
      ["Ir a Predicción individual y completar el formulario", "Devuelve insignia, probabilidad, decisión sugerida y top-5 factores."],
      ["Cargar Code/estudiantes_sinteticos_11500.csv en Carga masiva", "Procesa la cohorte sin errores y muestra los 4 KPIs."],
      ["Ir a Insights 360", "Muestra histograma, donut, KPIs y las 4 pestañas con gráficos."],
      ["Ir a Ayuda", "Muestra el diccionario y la tabla de las 5 bandas con la decisión sugerida."],
    ],
    [4200, 5160]
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 5. Anexo ----------
const sec5 = [
  H("5. Anexo · Esquema de variables", HeadingLevel.HEADING_1),
  P("Las 21 variables esperadas por SAT-DE, agrupadas por dimensión. Para el catálogo exhaustivo de categorías, consulte la página Ayuda dentro de la aplicación."),

  H("5.1 Variables académicas", HeadingLevel.HEADING_2),
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

  H("5.2 Variables demográficas y socioeconómicas", HeadingLevel.HEADING_2),
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

  H("5.3 Umbrales de riesgo (5 categorías)", HeadingLevel.HEADING_2),
  P("Cortes exactos calibrados sobre la Tabla de Eficiencia Operativa Real."),
  buildTable(
    [
      ["Banda", "Prob. mínima", "Prob. máxima", "% población esperado", "Tasa deserción interna", "Acción sugerida"],
      ["Bajo",    "0.0089", "0.3097", "53.07 %", "6.42 %",  "Permanencia"],
      ["Normal",  "0.3098", "0.5132", "13.00 %", "23.14 %", "Seguimiento"],
      ["Medio",   "0.5133", "0.6921", "8.54 %",  "38.59 %", "Plan de acompañamiento"],
      ["Alto",    "0.6927", "0.8486", "8.89 %",  "57.93 %", "Intervención focalizada"],
      ["Crítico", "0.8487", "0.9803", "16.50 %", "83.97 %", "Intervención inmediata"],
    ],
    [1300, 1300, 1300, 1700, 1900, 1860]
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// --------- 6. Soporte ----------
const sec6 = [
  H("6. Soporte y solución de problemas", HeadingLevel.HEADING_1),
  buildTable(
    [
      ["Síntoma", "Causa probable", "Acción sugerida"],
      ["La URL pública carga lenta la primera vez", "El App Service estaba en reposo (cold start).", "Esperar 30–60 segundos; cargas posteriores son rápidas."],
      ["La página no abre en localhost:8501", "Streamlit no quedó en ejecución.", "Revisar la terminal y volver a ejecutar streamlit run app.py."],
      ["Error 'Faltan columnas requeridas'", "El CSV no respeta el esquema.", "Descargar la plantilla desde Carga masiva y reutilizar sus columnas."],
      ["Advertencia 'Valores no vistos en …'", "Una fila usa un valor fuera del catálogo.", "Corregir el valor según el anexo (sección 5)."],
      ["Error 'Feature shape mismatch'", "Esquema del CSV inconsistente con el modelo.", "Verificar que las 21 columnas están presentes y con los nombres exactos."],
      ["Solo aparecen 1–2 bandas en la cohorte", "La cohorte tiene probabilidades muy concentradas.", "Es comportamiento esperado; revisar la distribución en Insights 360 → Distribución."],
      ["La predicción cambia mucho al editar un campo", "Variable de alto impacto (ej. materias reprobadas).", "Es comportamiento esperado; validar el valor real del estudiante."],
      ["El archivo .pkl no carga", "Pickle ausente o ruta incorrecta.", "Verificar que Code/modelo_xgboost_desercion_V1.1.1.pkl existe en el zip desplegado."],
      ["Despliegue en Azure devuelve 'Application Error'", "Falta el startup command o el WEBSITES_PORT.", "Re-aplicar sección 4.3 pasos 6 y 7; revisar logs con  az webapp log tail."],
    ],
    [2600, 3000, 3760]
  ),

  H("Contacto", HeadingLevel.HEADING_2),
  P("Equipo Grupo 10 · Maestría en Inteligencia Analítica de Datos (MIAD) · Universidad de los Andes. Para reportar incidencias, contactar al líder técnico del proyecto en el canal institucional asignado."),
];

// --------- documento ----------
const doc = new Document({
  creator: "MIAD Grupo 10",
  title: "Manual de Usuario · SAT-DE",
  description: "Manual de usuario del Sistema de Alerta Temprana de Deserción en Educación Superior (v2)",
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
            new TextRun({ text: "SAT-DE  ·  Manual de Usuario v2", color: BRAND, bold: true, size: 18 }),
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
    children: [...cover, ...toc, ...sec1, ...sec2, ...sec3, ...sec4, ...sec5, ...sec6],
  }],
});

const outPath = path.resolve(__dirname, "..", "Reportes", "Manual_Usuario_SAT-DE.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("OK ->", outPath, "(", buf.length, "bytes )");
});
