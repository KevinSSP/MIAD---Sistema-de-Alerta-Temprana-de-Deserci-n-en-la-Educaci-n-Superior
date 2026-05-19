/* Genera Reportes/Presentacion_SAT-DE.pptx
   Alineado con "Entrega Guia Anteproyecto.docx.pdf" */
const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Sanchez · Gómez · Peña · Porras";
pres.title = "SAT-DE · Modelo de Deserción en Educación Superior";
pres.company = "Universidad de los Andes · Ingeniería Industrial";

const W = 13.333, H = 7.5;

const NAVY = "1E3A8A";
const BLUE = "1D4ED8";
const SKY  = "0EA5E9";
const INK  = "0F172A";
const MUTED= "64748B";
const SOFT = "F1F5F9";
const BG   = "FFFFFF";
const OK   = "16A34A";
const WARN = "EAB308";
const BAD  = "DC2626";
const FONT_H = "Calibri";
const FONT_B = "Calibri";

function bg(slide, color = BG) { slide.background = { color }; }
function topBar(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.18, fill: { color: NAVY }, line: { color: NAVY } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.18, w: W, h: 0.06, fill: { color: SKY }, line: { color: SKY } });
}
function footer(slide, num, total) {
  slide.addText("SAT-DE · Modelo de Deserción en Educación Superior · Anteproyecto Grado 2025-2", {
    x: 0.4, y: H - 0.45, w: 10, h: 0.3, fontSize: 10, color: MUTED, fontFace: FONT_B, margin: 0,
  });
  slide.addText(`${num} / ${total}`, {
    x: W - 1.2, y: H - 0.45, w: 0.8, h: 0.3, fontSize: 10, color: MUTED,
    fontFace: FONT_B, align: "right", margin: 0,
  });
}
function title(slide, text, sub) {
  slide.addText(text, {
    x: 0.5, y: 0.45, w: W - 1, h: 0.75, fontSize: 28, bold: true,
    color: NAVY, fontFace: FONT_H, margin: 0,
  });
  if (sub) slide.addText(sub, {
    x: 0.5, y: 1.10, w: W - 1, h: 0.4, fontSize: 14, italic: true,
    color: MUTED, fontFace: FONT_B, margin: 0,
  });
}
function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: opts.fill || SOFT },
    line: { color: opts.border || "E2E8F0", width: 1 },
  });
}
function chip(slide, x, y, label, color = BLUE) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: 1.9, h: 0.32, fill: { color }, line: { color }, rectRadius: 0.16,
  });
  slide.addText(label.toUpperCase(), {
    x, y, w: 1.9, h: 0.32, fontSize: 9, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", fontFace: FONT_B, margin: 0, charSpacing: 2,
  });
}
function bullets(slide, items, x, y, w, h, opts = {}) {
  slide.addText(
    items.map((t, i) => ({ text: t, options: { bullet: { code: "25A0" }, breakLine: i < items.length - 1 } })),
    { x, y, w, h, fontSize: opts.size || 13, color: opts.color || INK,
      fontFace: FONT_B, valign: "top", paraSpaceAfter: 6 }
  );
}
function person(slide, x, y, initials) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: 1.45, h: 1.45, fill: { color: NAVY }, line: { color: SKY, width: 3 } });
  slide.addText(initials, {
    x, y, w: 1.45, h: 1.45, fontSize: 30, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", fontFace: FONT_H, margin: 0,
  });
}

const TOTAL = 10;

// ============== 1. PORTADA ==============
{
  const s = pres.addSlide(); bg(s, BG);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 5.4, h: H, fill: { color: NAVY }, line: { color: NAVY } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 0, w: 0.08, h: H, fill: { color: SKY }, line: { color: SKY } });

  s.addText("Universidad de los Andes  ·  Ingeniería Industrial", {
    x: 0.5, y: 0.55, w: 4.8, h: 0.35, fontSize: 11, color: SKY, bold: true,
    fontFace: FONT_B, charSpacing: 2, margin: 0,
  });
  s.addText("Anteproyecto Grado · Segundo Semestre 2025", {
    x: 0.5, y: 0.9, w: 4.8, h: 0.3, fontSize: 10, color: "CBD5F5",
    fontFace: FONT_B, italic: true, margin: 0,
  });
  s.addText("SAT-DE", {
    x: 0.5, y: 1.4, w: 4.8, h: 1.1, fontSize: 64, bold: true, color: "FFFFFF",
    fontFace: FONT_H, margin: 0,
  });
  s.addText("Sistema de Alerta Temprana\nde Deserción", {
    x: 0.5, y: 2.55, w: 4.8, h: 1.1, fontSize: 20, color: "FFFFFF", bold: true,
    fontFace: FONT_H, margin: 0,
  });
  s.addText("Modelo de Deserción en Educación Superior", {
    x: 0.5, y: 3.7, w: 4.8, h: 0.45, fontSize: 13, color: "CBD5F5", italic: true,
    fontFace: FONT_B, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 0.5, y: 4.2, w: 1.4, h: 0, line: { color: SKY, width: 3 } });
  s.addText("Sponsor:  Ministerio de Educación Nacional (MEN)\nFuentes:  SPADIES · DANE · ICFES · ICETEX\nAlineado con:  PND 2022–2026  ·  ODS 4 y 10", {
    x: 0.5, y: 4.4, w: 4.8, h: 1.7, fontSize: 12, color: "E2E8F0",
    fontFace: FONT_B, margin: 0, paraSpaceAfter: 6,
  });
  s.addText("Bogotá D.C. · 2025", {
    x: 0.5, y: H - 0.7, w: 4.8, h: 0.4, fontSize: 11, color: "94A3B8",
    fontFace: FONT_B, italic: true, margin: 0,
  });

  // Equipo
  s.addText("Equipo de trabajo", {
    x: 5.7, y: 0.55, w: 7.0, h: 0.35, fontSize: 12, color: MUTED, bold: true,
    fontFace: FONT_B, charSpacing: 3, margin: 0,
  });
  const team = [
    ["KS", "Kevin Snaider", "Sanchez Prieto"],
    ["JG", "Jorge Alberto", "Gómez Vigoya"],
    ["CP", "Carol Johana", "Peña Pico"],
    ["EP", "Elizabeth Lorena", "Porras Ortiz"],
  ];
  team.forEach((p, i) => {
    const xx = 5.8 + i * 1.78;
    person(s, xx, 1.0, p[0]);
    s.addText(p[1], { x: xx - 0.2, y: 2.55, w: 1.85, h: 0.3, fontSize: 11, color: INK, bold: true, align: "center", fontFace: FONT_B, margin: 0 });
    s.addText(p[2], { x: xx - 0.2, y: 2.82, w: 1.85, h: 0.3, fontSize: 9, color: MUTED, align: "center", fontFace: FONT_B, margin: 0 });
  });
  s.addText("Iniciales sustituibles por fotografías reales del equipo", {
    x: 5.7, y: 3.2, w: 7.0, h: 0.3, fontSize: 8, italic: true, color: MUTED, fontFace: FONT_B, margin: 0,
  });

  // Fachada vs Implementado
  s.addText("Prototipo Fachada  vs.  Prototipo Implementado", {
    x: 5.7, y: 3.6, w: 7.0, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: FONT_H, margin: 0,
  });

  card(s, 5.7, 4.05, 3.4, 2.7, { fill: SOFT, border: "CBD5F5" });
  s.addText("FACHADA · TAPA PROYECTO", {
    x: 5.85, y: 4.15, w: 3.2, h: 0.3, fontSize: 9, bold: true, color: MUTED,
    fontFace: FONT_B, charSpacing: 3, margin: 0,
  });
  s.addText("Bosquejo de pantallas\ny semáforo de alertas", {
    x: 5.85, y: 4.45, w: 3.2, h: 0.6, fontSize: 11, italic: true, color: INK, fontFace: FONT_B, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 5.15, w: 3.1, h: 1.5, fill: { color: "FFFFFF" }, line: { color: "94A3B8", width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.95, y: 5.15, w: 3.1, h: 0.25, fill: { color: NAVY }, line: { color: NAVY } });
  [BAD, WARN, OK].forEach((c, i) => {
    s.addShape(pres.shapes.OVAL, { x: 6.1 + i * 0.95, y: 5.55, w: 0.7, h: 0.7, fill: { color: c }, line: { color: c } });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.05, y: 6.32, w: 2.9, h: 0.25, fill: { color: "E2E8F0" }, line: { color: "E2E8F0" } });

  card(s, 9.25, 4.05, 3.4, 2.7, { fill: "EFF6FF", border: SKY });
  s.addText("IMPLEMENTADO · STREAMLIT", {
    x: 9.4, y: 4.15, w: 3.2, h: 0.3, fontSize: 9, bold: true, color: BLUE,
    fontFace: FONT_B, charSpacing: 3, margin: 0,
  });
  s.addText("App con XGBoost + SHAP\ny tablero Insights 360", {
    x: 9.4, y: 4.45, w: 3.2, h: 0.6, fontSize: 11, italic: true, color: INK, fontFace: FONT_B, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.5, y: 5.15, w: 3.1, h: 1.5, fill: { color: "FFFFFF" }, line: { color: SKY, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.5, y: 5.15, w: 3.1, h: 0.25, fill: { color: BLUE }, line: { color: BLUE } });
  for (let i = 0; i < 3; i++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 9.6 + i * 0.95, y: 5.5, w: 0.85, h: 0.45,
      fill: { color: i === 0 ? OK : (i === 1 ? WARN : BAD) },
      line: { color: i === 0 ? OK : (i === 1 ? WARN : BAD) },
    });
  }
  s.addShape(pres.shapes.RECTANGLE, { x: 9.6, y: 6.1, w: 2.9, h: 0.55, fill: { color: SOFT }, line: { color: SOFT } });
  [0.2, 0.5, 0.8, 0.6, 0.35, 0.25].forEach((hv, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 9.7 + i * 0.45, y: 6.1 + (0.55 - 0.55 * hv), w: 0.30, h: 0.55 * hv,
      fill: { color: BLUE }, line: { color: BLUE },
    });
  });
}

// ============== 2. PROBLEMA ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "El problema: ¿quién, qué y cómo lo medimos?",
    "Esquematización del caso (Parte 1) · Carta del Proyecto");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.6, w: W - 1, h: 1.05, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("La deserción acumulada en educación superior en Colombia supera el 45% (SPADIES, 2025), comprometiendo capital humano y eficiencia del gasto público. El PND 2022–2026 plantea sistemas de alertas tempranas, pero no existe una herramienta analítica que integre datos socioeconómicos y académicos y explique sus predicciones.", {
    x: 0.8, y: 1.65, w: W - 1.6, h: 0.95, fontSize: 12, color: "FFFFFF",
    fontFace: FONT_B, valign: "middle", margin: 0,
  });

  const cy = 2.95, ch = 3.55, cw = 4.0, gap = 0.18;
  const cxs = [0.5, 0.5 + cw + gap, 0.5 + 2 * (cw + gap)];

  card(s, cxs[0], cy, cw, ch);
  chip(s, cxs[0] + 0.2, cy + 0.2, "Usuarios", NAVY);
  s.addText("¿Quién lo usa?", { x: cxs[0] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 17, bold: true, color: INK, fontFace: FONT_H, margin: 0 });
  bullets(s, [
    "Sponsor: Ministerio de Educación Nacional (MEN)",
    "Directores de Bienestar Universitario de las IES",
    "Direcciones de Planeación / Permanencia",
    "MEN · Subdirección de Desarrollo Sectorial",
    "Proveedores de datos: DANE e ICFES",
  ], cxs[0] + 0.2, cy + 1.1, cw - 0.4, 2.3);

  card(s, cxs[1], cy, cw, ch);
  chip(s, cxs[1] + 0.2, cy + 0.2, "Requerimientos", BLUE);
  s.addText("¿Qué necesitan?", { x: cxs[1] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 17, bold: true, color: INK, fontFace: FONT_H, margin: 0 });
  bullets(s, [
    "Predecir riesgo individual (bajo · medio · alto)",
    "Procesar cohortes completas vía CSV",
    "Explicar la decisión por estudiante (SHAP)",
    "Dashboard interactivo con filtros regionales",
    "Latencia < 5 s y operación sin programar",
  ], cxs[1] + 0.2, cy + 1.1, cw - 0.4, 2.3);

  card(s, cxs[2], cy, cw, ch);
  chip(s, cxs[2] + 0.2, cy + 0.2, "Métricas", SKY);
  s.addText("¿Cómo lo medimos?", { x: cxs[2] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 17, bold: true, color: INK, fontFace: FONT_H, margin: 0 });
  bullets(s, [
    "Negocio: reducción ≥ 5 % deserción semestral",
    "Adopción: 100 % de las 21 IES piloto",
    "Técnicas: AUC ≥ 0.85 · F1 ≥ 0.80 (clase minoritaria)",
    "Calidad de datos: completitud > 95 %",
    "Satisfacción del usuario ≥ 85 %",
  ], cxs[2] + 0.2, cy + 1.1, cw - 0.4, 2.3);

  footer(s, 2, TOTAL);
}

// ============== 3. SOLUCIÓN GENERAL ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Cómo funciona SAT-DE en una mirada",
    "Tres bloques que convierten datos oficiales en decisiones accionables");

  const by = 2.4, bh = 3.6, bw = 3.8, gap = 0.4;
  const xs = [0.5, 0.5 + bw + gap, 0.5 + 2 * (bw + gap)];

  card(s, xs[0], by, bw, bh, { fill: "EFF6FF", border: SKY });
  s.addShape(pres.shapes.OVAL, { x: xs[0] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fill: { color: SKY }, line: { color: SKY } });
  s.addText("1", { x: xs[0] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fontSize: 28, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, fontFace: FONT_H });
  s.addText("Datos oficiales", { x: xs[0] + 1.15, y: by + 0.4, w: bw - 1.35, h: 0.5, fontSize: 18, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  s.addText("SPADIES · DANE · ICFES · ICETEX", { x: xs[0] + 1.15, y: by + 0.85, w: bw - 1.35, h: 0.3, fontSize: 11, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  bullets(s, [
    "Académicas (rendimiento, Saber 11)",
    "Socioeconómicas (estrato, ICETEX, becas)",
    "Demográficas y regionales",
    "21 campos por estudiante; lote o individual",
  ], xs[0] + 0.3, by + 1.4, bw - 0.6, 2.0, { size: 12 });

  card(s, xs[1], by, bw, bh, { fill: "DBEAFE", border: BLUE });
  s.addShape(pres.shapes.OVAL, { x: xs[1] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fill: { color: BLUE }, line: { color: BLUE } });
  s.addText("2", { x: xs[1] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fontSize: 28, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, fontFace: FONT_H });
  s.addText("Modelo predictivo", { x: xs[1] + 1.15, y: by + 0.4, w: bw - 1.35, h: 0.5, fontSize: 18, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  s.addText("XGBoost (ensamble) + SHAP", { x: xs[1] + 1.15, y: by + 0.85, w: bw - 1.35, h: 0.3, fontSize: 11, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  bullets(s, [
    "Maneja desbalance (scale_pos_weight)",
    "Robusto al ruido administrativo histórico",
    "Score de probabilidad (no solo sí/no)",
    "Explica las 5 variables que más pesaron",
  ], xs[1] + 0.3, by + 1.4, bw - 0.6, 2.0, { size: 12 });

  card(s, xs[2], by, bw, bh, { fill: "DCFCE7", border: OK });
  s.addShape(pres.shapes.OVAL, { x: xs[2] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fill: { color: OK }, line: { color: OK } });
  s.addText("3", { x: xs[2] + 0.3, y: by + 0.3, w: 0.7, h: 0.7, fontSize: 28, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, fontFace: FONT_H });
  s.addText("Sistema de alertas", { x: xs[2] + 1.15, y: by + 0.4, w: bw - 1.35, h: 0.5, fontSize: 18, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  s.addText("App web · Semáforo Rojo / Amarillo / Verde", { x: xs[2] + 1.15, y: by + 0.85, w: bw - 1.35, h: 0.3, fontSize: 11, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  bullets(s, [
    "Predicción individual con explicación",
    "Carga masiva de cohortes (CSV)",
    "Dashboard Insights 360 por regiones",
    "Recomendación de intervención focalizada",
  ], xs[2] + 0.3, by + 1.4, bw - 0.6, 2.0, { size: 12 });

  s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: xs[0] + bw + 0.05, y: by + 1.6, w: 0.3, h: 0.4, fill: { color: NAVY }, line: { color: NAVY }, rotate: 90 });
  s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: xs[1] + bw + 0.05, y: by + 1.6, w: 0.3, h: 0.4, fill: { color: NAVY }, line: { color: NAVY }, rotate: 90 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 6.4, w: W - 1, h: 0.55, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("Pasa de predictivo (qué pasará) a prescriptivo (qué hacer): cada alerta es accionable para los Directores de Bienestar.", {
    x: 0.7, y: 6.4, w: W - 1.4, h: 0.55, fontSize: 12, color: "FFFFFF",
    fontFace: FONT_B, valign: "middle", italic: true, margin: 0,
  });

  footer(s, 3, TOTAL);
}

// ============== 4. DEMO INDIVIDUAL ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Demo · Predicción individual",
    "Caso 1 — Score 0–1, semáforo y top-5 factores (SHAP)");

  card(s, 0.5, 1.7, 5.5, 5.1, { fill: SOFT });
  s.addText("Formulario guiado · 21 variables", { x: 0.7, y: 1.85, w: 5.1, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  ["Demográfico", "Socioeconómico (estrato · ICETEX · becas)", "Académico (Saber 11 · promedio)"].forEach((sec, i) => {
    const yy = 2.35 + i * 1.45;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: yy, w: 5.1, h: 1.3, fill: { color: "FFFFFF" }, line: { color: "E2E8F0", width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: yy, w: 0.12, h: 1.3, fill: { color: BLUE }, line: { color: BLUE } });
    s.addText(sec, { x: 0.95, y: yy + 0.08, w: 4.8, h: 0.3, fontSize: 11, bold: true, color: NAVY, fontFace: FONT_B, margin: 0 });
    for (let k = 0; k < 2; k++) {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.95 + k * 2.45, y: yy + 0.5, w: 2.3, h: 0.32, fill: { color: SOFT }, line: { color: "CBD5F5" } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.95 + k * 2.45, y: yy + 0.92, w: 2.3, h: 0.32, fill: { color: SOFT }, line: { color: "CBD5F5" } });
    }
  });

  card(s, 6.3, 1.7, 6.5, 5.1, { fill: "FFFFFF", border: BLUE });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.9, w: 2.2, h: 0.5, fill: { color: BAD }, line: { color: BAD }, rectRadius: 0.1 });
  s.addText("RIESGO ALTO", { x: 6.5, y: 1.9, w: 2.2, h: 0.5, fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT_H, margin: 0, charSpacing: 3 });
  s.addText("0.72", { x: 8.9, y: 1.78, w: 1.8, h: 0.75, fontSize: 42, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  s.addText("Probabilidad de deserción", { x: 10.7, y: 2.0, w: 2.0, h: 0.4, fontSize: 10, color: MUTED, fontFace: FONT_B, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 2.6, w: 6.1, h: 0.3, fill: { color: SOFT }, line: { color: SOFT } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 2.6, w: 6.1 * 0.72, h: 0.3, fill: { color: BAD }, line: { color: BAD } });

  s.addText("Top-5 factores (SHAP)", { x: 6.5, y: 3.1, w: 6.1, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  const factors = [
    { label: "Promedio académico", val: 0.28, dir: "Riesgo", color: BAD },
    { label: "Materias reprobadas", val: 0.22, dir: "Riesgo", color: BAD },
    { label: "Estrato socioeconómico", val: 0.15, dir: "Riesgo", color: BAD },
    { label: "Apoyo ICETEX / beca", val: 0.10, dir: "Protección", color: OK },
    { label: "Puntaje Saber 11", val: 0.08, dir: "Protección", color: OK },
  ];
  factors.forEach((f, i) => {
    const yy = 3.55 + i * 0.5;
    s.addText(f.label, { x: 6.5, y: yy, w: 2.6, h: 0.35, fontSize: 11, color: INK, fontFace: FONT_B, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 9.2, y: yy + 0.05, w: 2.4, h: 0.22, fill: { color: SOFT }, line: { color: SOFT } });
    s.addShape(pres.shapes.RECTANGLE, { x: 9.2, y: yy + 0.05, w: 2.4 * (f.val / 0.30), h: 0.22, fill: { color: f.color }, line: { color: f.color } });
    s.addText(f.dir, { x: 11.7, y: yy, w: 0.9, h: 0.32, fontSize: 9, bold: true, color: f.color, align: "right", fontFace: FONT_B, margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 6.15, w: 6.5, h: 0.65, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("Recomendación: tutoría académica, revisar carga de materias y validar apoyo ICETEX vigente.", {
    x: 6.5, y: 6.15, w: 6.1, h: 0.65, fontSize: 11, color: "FFFFFF", italic: true,
    fontFace: FONT_B, valign: "middle", margin: 0,
  });

  footer(s, 4, TOTAL);
}

// ============== 5. DEMO CARGA + INSIGHTS ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Demo · Carga masiva e Insights 360",
    "Casos 2 y 3 — De la cohorte completa al análisis regionalizado");

  card(s, 0.5, 1.7, 6.2, 5.1);
  chip(s, 0.7, 1.85, "Caso 2 · Carga masiva", BLUE);
  s.addText("Procesamiento por lote (CSV)", { x: 0.7, y: 2.25, w: 5.8, h: 0.4, fontSize: 17, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });

  const kpis = [
    { label: "Estudiantes", value: "11.500", color: NAVY },
    { label: "Alto riesgo", value: "1.842", color: BAD },
    { label: "Medio", value: "3.205", color: WARN },
    { label: "Bajo", value: "6.453", color: OK },
  ];
  kpis.forEach((k, i) => {
    const xx = 0.7 + i * 1.45;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 2.8, w: 1.35, h: 1.1, fill: { color: SOFT }, line: { color: "E2E8F0" } });
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 2.8, w: 1.35, h: 0.08, fill: { color: k.color }, line: { color: k.color } });
    s.addText(k.value, { x: xx, y: 2.95, w: 1.35, h: 0.5, fontSize: 18, bold: true, color: k.color, align: "center", fontFace: FONT_H, margin: 0 });
    s.addText(k.label, { x: xx, y: 3.45, w: 1.35, h: 0.4, fontSize: 9, color: MUTED, align: "center", fontFace: FONT_B, margin: 0 });
  });

  s.addText("Resultados descargables en CSV con score y riesgo", { x: 0.7, y: 4.05, w: 5.8, h: 0.3, fontSize: 10, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  s.addTable([
    [
      { text: "ID", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Prob.", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Riesgo", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Pred.", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
    ],
    ["E-00128", "0.84", { text: "Alto", options: { color: BAD, bold: true } }, "1"],
    ["E-04302", "0.71", { text: "Alto", options: { color: BAD, bold: true } }, "1"],
    ["E-01055", "0.48", { text: "Medio", options: { color: WARN, bold: true } }, "0"],
    ["E-09711", "0.22", { text: "Bajo", options: { color: OK, bold: true } }, "0"],
  ], {
    x: 0.7, y: 4.4, w: 5.8, colW: [1.6, 1.4, 1.6, 1.2],
    fontSize: 11, fontFace: FONT_B, color: INK, border: { pt: 0.5, color: "E2E8F0" }, rowH: 0.34,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 6.25, w: 2.0, h: 0.4, fill: { color: BLUE }, line: { color: BLUE }, rectRadius: 0.08 });
  s.addText("⬇  Descargar CSV", { x: 0.7, y: 6.25, w: 2.0, h: 0.4, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT_B, margin: 0 });

  card(s, 7.0, 1.7, 5.8, 5.1, { fill: "EFF6FF", border: BLUE });
  chip(s, 7.2, 1.85, "Caso 3 · Insights 360", SKY);
  s.addText("Análisis de cohorte por región", { x: 7.2, y: 2.25, w: 5.4, h: 0.4, fontSize: 17, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });

  s.addShape(pres.shapes.OVAL, { x: 7.4, y: 2.85, w: 2.0, h: 2.0, fill: { color: OK }, line: { color: "FFFFFF", width: 4 } });
  s.addShape(pres.shapes.PIE, { x: 7.4, y: 2.85, w: 2.0, h: 2.0, fill: { color: WARN }, line: { color: "FFFFFF", width: 4 }, rotate: 0 });
  s.addShape(pres.shapes.OVAL, { x: 7.85, y: 3.30, w: 1.1, h: 1.1, fill: { color: "EFF6FF" }, line: { color: "EFF6FF" } });
  s.addText("56%\nBajo", { x: 7.85, y: 3.30, w: 1.1, h: 1.1, fontSize: 11, bold: true, color: NAVY, align: "center", valign: "middle", fontFace: FONT_H, margin: 0 });

  s.addText("Riesgo por región (cohorte)", { x: 9.6, y: 2.85, w: 3.1, h: 0.3, fontSize: 10, bold: true, color: NAVY, fontFace: FONT_B, margin: 0 });
  ["Orinoquía", "Centro", "Costa Caribe", "Pacífico", "Andina", "Amazonía"].forEach((a, i) => {
    const yy = 3.2 + i * 0.27;
    const v = [0.46, 0.30, 0.42, 0.38, 0.28, 0.50][i];
    s.addText(a, { x: 9.6, y: yy, w: 1.3, h: 0.22, fontSize: 9, color: INK, fontFace: FONT_B, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.85, y: yy + 0.04, w: 1.5, h: 0.16, fill: { color: SOFT }, line: { color: SOFT } });
    s.addShape(pres.shapes.RECTANGLE, { x: 10.85, y: yy + 0.04, w: 1.5 * v, h: 0.16, fill: { color: BLUE }, line: { color: BLUE } });
    s.addText(`${Math.round(v * 100)}%`, { x: 12.4, y: yy, w: 0.35, h: 0.22, fontSize: 9, color: MUTED, fontFace: FONT_B, margin: 0 });
  });

  s.addText("4 pestañas · Distribución · Demográfico · Académico · Detalle", { x: 7.2, y: 5.05, w: 5.4, h: 0.35, fontSize: 10, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 5.45, w: 5.4, h: 1.2, fill: { color: "FFFFFF" }, line: { color: "E2E8F0" } });
  for (let i = 0; i < 40; i++) {
    const rx = 7.3 + Math.random() * 5.1;
    const ry = 5.55 + Math.random() * 1.0;
    const c = Math.random() > 0.5 ? OK : (Math.random() > 0.5 ? WARN : BAD);
    s.addShape(pres.shapes.OVAL, { x: rx, y: ry, w: 0.08, h: 0.08, fill: { color: c }, line: { color: c } });
  }

  footer(s, 5, TOTAL);
}

// ============== 6. GARANTÍAS vs LIMITACIONES ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Garantías y limitaciones del prototipo",
    "Contraste con las metas del anteproyecto y resultados de pruebas");

  const score = [
    ["Meta de modelo", "AUC ≥ 0.85", "AUC 0.89", true],
    ["Meta de modelo", "F1 ≥ 0.80 (clase minoritaria)", "F1 0.79", false],
    ["Latencia dashboard", "< 5 s", "< 5 s para 11.500 filas", true],
    ["Calidad de datos", "Completitud > 95 %", "Validación de columnas en CSV", true],
    ["Explicabilidad", "SHAP por instancia", "Top-5 factores por estudiante", true],
  ];
  s.addTable([
    [
      { text: "Dimensión", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Meta del anteproyecto", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Resultado del prototipo", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
      { text: "Estado", options: { bold: true, color: "FFFFFF", fill: { color: NAVY } } },
    ],
    ...score.map((r) => [
      r[0], r[1], r[2],
      { text: r[3] ? "Cumplido" : "Parcial", options: { color: r[3] ? OK : WARN, bold: true } },
    ]),
  ], {
    x: 0.5, y: 1.7, w: W - 1, colW: [2.6, 3.6, 4.5, 1.6],
    fontSize: 11, fontFace: FONT_B, color: INK, border: { pt: 0.5, color: "E2E8F0" }, rowH: 0.35,
  });

  card(s, 0.5, 4.05, 6.2, 2.8, { fill: "DCFCE7", border: OK });
  chip(s, 0.7, 4.2, "Garantías", OK);
  bullets(s, [
    "ROC AUC 0.89 supera la meta del anteproyecto (≥ 0.85)",
    "Score 0–1 traducido a semáforo (Rojo · Amarillo · Verde)",
    "Top-5 factores por estudiante para evitar caja negra",
    "Tolerante a categorías desconocidas (handle_unknown)",
    "Reproducible: semillas fijas, requirements y notebook",
  ], 0.7, 4.6, 5.8, 2.15, { size: 12 });

  card(s, 7.0, 4.05, 5.8, 2.8, { fill: "FEF3C7", border: WARN });
  chip(s, 7.2, 4.2, "Limitaciones", BAD);
  bullets(s, [
    "F1 = 0.79 ≈ 0.80 (rozando la meta): seguir optimizando",
    "Datos sintetizados para piloto; falta integración SPADIES en vivo",
    "Sin auth ni auditoría de usuarios (prototipo)",
    "No reemplaza juicio del consejero ni habilita sanciones",
    "Requiere los 21 campos del esquema",
  ], 7.2, 4.6, 5.4, 2.15, { size: 12 });

  footer(s, 6, TOTAL);
}

// ============== 7. PROPUESTA DE VALOR ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Propuesta de valor",
    "Impacto esperado para el MEN, las IES piloto y los estudiantes");

  const items = [
    { t: "Reducción del 5 %", d: "Meta del anteproyecto: bajar la deserción semestral en al menos 5 puntos en las IES piloto.", c: BAD },
    { t: "Eficiencia del gasto público", d: "Optimiza recursos del MEN al retener estudiantes con becas y subsidios ya invertidos.", c: BLUE },
    { t: "Decisión explicable", d: "SHAP responde el porqué de cada alerta, eliminando la percepción de caja negra.", c: SKY },
    { t: "Escala 21 IES → sistema", d: "Diseño replicable: del piloto de 21 IES al sistema de educación superior completo.", c: NAVY },
    { t: "Alineado con PND y ODS", d: "Apoya el PND 2022–2026 y los ODS 4 (educación de calidad) y 10 (reducción de desigualdades).", c: OK },
    { t: "Equidad regional", d: "Permite focalizar acciones por región (Orinoquía, Caribe, Pacífico, etc.) y por estrato.", c: BLUE },
  ];
  const cw = 4.0, ch = 1.55, gap = 0.25;
  items.forEach((it, idx) => {
    const col = idx % 3, row = Math.floor(idx / 3);
    const xx = 0.5 + col * (cw + gap);
    const yy = 1.85 + row * (ch + gap);
    card(s, xx, yy, cw, ch);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: yy, w: 0.12, h: ch, fill: { color: it.c }, line: { color: it.c } });
    s.addText(it.t, { x: xx + 0.25, y: yy + 0.12, w: cw - 0.4, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
    s.addText(it.d, { x: xx + 0.25, y: yy + 0.55, w: cw - 0.4, h: 0.95, fontSize: 11, color: INK, fontFace: FONT_B, margin: 0, valign: "top" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.6, w: W - 1, h: 1.2, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("89 %", { x: 0.7, y: 5.6, w: 2.5, h: 1.2, fontSize: 56, bold: true, color: SKY, align: "center", valign: "middle", fontFace: FONT_H, margin: 0 });
  s.addText("ROC AUC del prototipo, frente a la meta de 0.85 del anteproyecto. Capacidad demostrada para apoyar la meta nacional de reducir la deserción y cerrar la brecha del 45 % reportada por SPADIES (2025).", {
    x: 3.3, y: 5.6, w: W - 4.0, h: 1.2, fontSize: 12, color: "FFFFFF",
    fontFace: FONT_B, valign: "middle", italic: true, margin: 0,
  });

  footer(s, 7, TOTAL);
}

// ============== 8. COSTOS / RIESGOS / ADOPCIÓN ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Costos, riesgos y condiciones de adopción",
    "Recursos, AMEF (NPR) y stakeholders definidos en el anteproyecto");

  const cy = 1.75, ch = 5.1, cw = 4.05, gap = 0.18;
  const xs = [0.5, 0.5 + cw + gap, 0.5 + 2 * (cw + gap)];

  card(s, xs[0], cy, cw, ch);
  chip(s, xs[0] + 0.2, cy + 0.2, "Costos · Recursos", BLUE);
  s.addText("Inversión esperada", { x: xs[0] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 15, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  s.addTable([
    [{ text: "Concepto", options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11 } },
     { text: "Estimado", options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 11 } }],
    ["Líder de Proyecto", "16 semanas"],
    ["Científico de Datos Senior", "FTE 1.0"],
    ["Ingeniero de Datos (ETL)", "FTE 0.8"],
    ["Analista BI / Visualización", "FTE 0.6"],
    ["Cloud AWS/Azure", "~ USD 18 / mes"],
    ["Power BI / Tableau", "Lic. institucional"],
    ["Reentrenamiento semestral", "2 jornadas / semestre"],
  ], {
    x: xs[0] + 0.2, y: cy + 1.1, w: cw - 0.4, colW: [(cw - 0.4) * 0.65, (cw - 0.4) * 0.35],
    fontSize: 10, fontFace: FONT_B, color: INK, border: { pt: 0.5, color: "E2E8F0" }, rowH: 0.30,
  });

  card(s, xs[1], cy, cw, ch);
  chip(s, xs[1] + 0.2, cy + 0.2, "Riesgos · AMEF", BAD);
  s.addText("Top-3 con mitigación", { x: xs[1] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 15, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  const risks = [
    ["NPR 240 · Crítico", "Retraso integración SPADIES/ICFES por llaves inconsistentes", "PoC de integración (sem. 2) + llaves subrogadas", BAD],
    ["NPR 180 · Medio-Alto", "Model Drift por cambios en políticas de becas o socioeconómicas", "Pipeline MLOps con PSI > 0.2 → reentrenamiento", WARN],
    ["NPR 120 · Medio", "Resistencia / percepción de caja negra en las IES", "SHAP por estudiante + talleres de co-diseño", BLUE],
  ];
  risks.forEach((r, i) => {
    const yy = cy + 1.15 + i * 1.25;
    s.addShape(pres.shapes.RECTANGLE, { x: xs[1] + 0.2, y: yy, w: cw - 0.4, h: 0.32, fill: { color: r[3] }, line: { color: r[3] } });
    s.addText(r[0], { x: xs[1] + 0.3, y: yy, w: cw - 0.6, h: 0.32, fontSize: 10, bold: true, color: "FFFFFF", valign: "middle", fontFace: FONT_B, margin: 0, charSpacing: 1 });
    s.addText(r[1], { x: xs[1] + 0.2, y: yy + 0.36, w: cw - 0.4, h: 0.45, fontSize: 10, color: INK, fontFace: FONT_B, margin: 0 });
    s.addText("Mitigación: " + r[2], { x: xs[1] + 0.2, y: yy + 0.8, w: cw - 0.4, h: 0.4, fontSize: 9, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  });

  card(s, xs[2], cy, cw, ch);
  chip(s, xs[2] + 0.2, cy + 0.2, "Adopción", OK);
  s.addText("Condiciones de despliegue", { x: xs[2] + 0.2, y: cy + 0.6, w: cw - 0.4, h: 0.4, fontSize: 15, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
  bullets(s, [
    "Sponsor MEN · Subdirección de Desarrollo Sectorial",
    "Convenio con SPADIES, DANE e ICFES",
    "21 IES piloto con Director de Bienestar designado",
    "Pipeline MLOps (MLflow / Airflow) en cloud",
    "Manual de Usuario y Handover técnico al equipo TI",
    "Encuesta NPS / CSAT a stakeholders ≥ 85 %",
    "Cumplimiento Ley 1581 de 2012 (habeas data)",
  ], xs[2] + 0.3, cy + 1.15, cw - 0.5, 3.85, { size: 11 });

  footer(s, 8, TOTAL);
}

// ============== 9. HOJA DE RUTA ==============
{
  const s = pres.addSlide(); bg(s); topBar(s);
  title(s, "Hoja de ruta hacia producción",
    "Cinco fases del anteproyecto (CRISP-DM · 16 semanas) + escalamiento");

  const stages = [
    { t: "Sem. 1-2",   st: "Definir",     d: "Carta del Proyecto, alcance y validación con MEN.", c: OK,   r: "Líder / Sponsor" },
    { t: "Sem. 3-5",   st: "Medir",       d: "Ingesta SPADIES/DANE/ICFES, perfilamiento y línea base.", c: SKY,  r: "Ing. de Datos" },
    { t: "Sem. 6-9",   st: "Analizar",    d: "Feature engineering, RF/XGBoost, SHAP y validación.", c: BLUE, r: "Científico de Datos" },
    { t: "Sem. 10-14", st: "Implementar", d: "Dashboard, integración productiva y capacitación.", c: NAVY, r: "Analista BI / Líder" },
    { t: "Sem. 15-16", st: "Controlar",   d: "Monitoreo, KPIs de adopción y cierre formal.", c: BAD,  r: "Todo el equipo" },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 4.2, w: W - 1.4, h: 0.08, fill: { color: NAVY }, line: { color: NAVY } });

  const slotW = (W - 1.4) / 5;
  stages.forEach((stg, i) => {
    const xx = 0.7 + i * slotW + 0.15;
    const cardW = slotW - 0.3;
    s.addShape(pres.shapes.OVAL, { x: xx + cardW / 2 - 0.2, y: 4.05, w: 0.4, h: 0.4, fill: { color: stg.c }, line: { color: "FFFFFF", width: 3 } });
    s.addText(stg.t, { x: xx, y: 2.2, w: cardW, h: 0.3, fontSize: 10, bold: true, color: stg.c, align: "center", fontFace: FONT_B, charSpacing: 2, margin: 0 });
    card(s, xx, 2.55, cardW, 1.35);
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 2.55, w: cardW, h: 0.1, fill: { color: stg.c }, line: { color: stg.c } });
    s.addText(stg.st, { x: xx + 0.12, y: 2.7, w: cardW - 0.2, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: FONT_H, margin: 0 });
    s.addText(stg.d, { x: xx + 0.12, y: 3.05, w: cardW - 0.2, h: 0.6, fontSize: 9, color: INK, fontFace: FONT_B, margin: 0, valign: "top" });
    s.addText(stg.r, { x: xx + 0.12, y: 3.65, w: cardW - 0.2, h: 0.22, fontSize: 8, italic: true, color: MUTED, fontFace: FONT_B, margin: 0 });
  });

  s.addText("Tras las 16 semanas · Escalamiento", {
    x: 0.7, y: 4.7, w: W - 1.4, h: 0.35, fontSize: 14, bold: true, color: NAVY, fontFace: FONT_H, margin: 0,
  });
  bullets(s, [
    "Despliegue final en servidores del cliente (MEN / IES) con pruebas de estrés.",
    "Hoja de ruta para pasar de 21 IES piloto al sistema de educación superior completo.",
    "Reentrenamiento semestral con cortes académicos; monitoreo PSI + F1.",
    "Encuesta NPS/CSAT y documentación de Lecciones Aprendidas.",
  ], 0.8, 5.1, W - 1.6, 1.4, { size: 12 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 6.55, w: W - 1, h: 0.45, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("Próximo paso: acordar IES piloto, acceso a datos históricos SPADIES (3 años) y designar al responsable institucional del modelo.", {
    x: 0.7, y: 6.55, w: W - 1.4, h: 0.45, fontSize: 11, color: "FFFFFF", italic: true,
    fontFace: FONT_B, valign: "middle", margin: 0,
  });

  footer(s, 9, TOTAL);
}

// ============== 10. CIERRE ==============
{
  const s = pres.addSlide();
  bg(s, NAVY);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: H, fill: { color: SKY }, line: { color: SKY } });

  s.addText("Gracias", { x: 0.8, y: 1.6, w: 12, h: 1.8, fontSize: 96, bold: true, color: "FFFFFF", fontFace: FONT_H, margin: 0 });
  s.addText("¿Preguntas?", { x: 0.8, y: 3.5, w: 12, h: 0.7, fontSize: 28, color: SKY, fontFace: FONT_H, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 4.4, w: 2.4, h: 0, line: { color: SKY, width: 4 } });

  s.addText("SAT-DE · Sistema de Alerta Temprana de Deserción", { x: 0.8, y: 4.7, w: 12, h: 0.5, fontSize: 18, color: "CBD5F5", fontFace: FONT_B, margin: 0 });
  s.addText("Modelo de Deserción en Educación Superior · Anteproyecto Grado · 2025-2", { x: 0.8, y: 5.2, w: 12, h: 0.5, fontSize: 14, color: "94A3B8", fontFace: FONT_B, margin: 0 });
  s.addText("Kevin Sanchez · Jorge Gómez · Carol Peña · Elizabeth Porras", { x: 0.8, y: 5.65, w: 12, h: 0.4, fontSize: 12, color: "CBD5F5", fontFace: FONT_B, italic: true, margin: 0 });
  s.addText("Universidad de los Andes  ·  Facultad de Ingeniería  ·  Ingeniería Industrial", { x: 0.8, y: 6.0, w: 12, h: 0.4, fontSize: 11, color: "94A3B8", fontFace: FONT_B, margin: 0 });
  s.addText("Repositorio · Manual de Usuario · Anexo Técnico · Anteproyecto", { x: 0.8, y: 6.5, w: 12, h: 0.4, fontSize: 11, italic: true, color: "94A3B8", fontFace: FONT_B, margin: 0 });
}

const out = path.resolve(__dirname, "Presentacion_SAT-DE.pptx");
pres.writeFile({ fileName: out }).then((f) => console.log("OK ->", f));
