# 🎓 Sistema de Alerta Temprana de Deserción Estudiantil (SAT-DE)

[![Python](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/)
[![Machine Learning](https://img.shields.io/badge/Model-XGBoost-orange.svg)](https://xgboost.readthedocs.io/)
[![Status](https://img.shields.io/badge/Status-En%20Desarrollo-green.svg)]()

## 📌 Descripción del Proyecto
La deserción en la educación superior en Colombia es un fenómeno complejo con profundas raíces socioeconómicas y académicas[cite: 3]. Este proyecto, desarrollado como parte del **Proyecto Aplicado en Analítica de Datos (PAAD)** de la **Maestría en Inteligencia Analítica de Datos (MIAD)**, utiliza técnicas avanzadas de *Machine Learning* para identificar patrones de riesgo y emitir alertas tempranas que permitan intervenciones institucionales oportunas.

A través de este sistema, buscamos transformar datos históricos en decisiones estratégicas para mejorar la permanencia estudiantil.

---

## 🛠️ Stack Tecnológico
*   **Lenguaje:** Python 3.12.
*   **Análisis de Datos:** `Pandas`, `NumPy`.
*   **Modelado Predictivo:** `Scikit-learn`, `XGBoost`, `Random Forest`.
*   **Optimización:** `Optuna` para el ajuste fino de hiperparámetros.
*   **Visualización:** `Matplotlib`, `Seaborn`.

---

## 📊 Metodología y Datos
El sistema se entrena con un conjunto de datos robusto que comprende **31,567 registros** y **25 variables** que cubren dimensiones demográficas, socioeconómicas y académicas.

### Hallazgos Clave del Análisis (EDA):
*   **Factores Académicos:** El promedio académico acumulado y el número de materias reprobadas han sido identificados como los predictores con mayor importancia para el modelo.
*   **Perfil Socioeconómico:** Existe una correlación significativa entre el estrato socioeconómico, el acceso a becas y la persistencia del estudiante en el sistema.
*   **Tratamiento de Datos:** Se implementó un pipeline que incluye codificación categórica (`OneHotEncoder`), normalización numérica (`StandardScaler`) y manejo de desbalanceo de clases.

---

## 📂 Estructura del Repositorio
El repositorio está organizado para facilitar la reproducibilidad del análisis y el despliegue del prototipo:

*   📁 **`codigo/`**: Contiene los Jupyter Notebooks con el análisis exploratorio (EDA), limpieza de datos, entrenamiento de modelos competitivos y los scripts del aplicativo funcional.
*   📁 **`reportes/`**: Documentación técnica detallada, entregables del módulo 2 en formato PDF/Word y el Diccionario de Datos exhaustivo del proyecto.

---

## 🚀 Próximos Pasos
El proyecto avanza hacia la consolidación de un prototipo funcional que incluye:
1.  **Dashboard de Visualización:** Interfaz interactiva para que los consejeros académicos visualicen el nivel de riesgo en un semáforo (Rojo/Amarillo/Verde).

---

## 👥 Equipo de Trabajo - Grupo 10
*   **Carol Johana Peña Pico**.
*   **Kevin Snaider Sánchez Prieto**
*   **Jorge Alberto Gómez Vigoya**.
*   **Elizabeth Lorena Porras Ortiz**.

---
*Este repositorio es el resultado académico para el curso Proyecto Aplicado en Analítica de Datos (PAAD) - 2026.*
