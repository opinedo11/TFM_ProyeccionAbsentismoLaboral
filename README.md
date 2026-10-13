# Predicción del absentismo laboral mediante aprendizaje automático

Repositorio del Trabajo Fin de Máster: Aplicación de aprendizaje automático para predecir el absentismo en un grupo hospitalario.

El proyecto aborda la preparación de datos históricos de incapacidad temporal (IT), la construcción de variables predictoras y el entrenamiento y evaluación de modelos de aprendizaje automático para estimar la IT en horizontes semanales futuros.

## Estructura del repositorio

Los notebooks y scripts se organizan en una secuencia de trabajo:

| Archivo | Propósito |
|---|---|
| `01_generar_tabla_entrenamiento_IT.py` | Generación de la tabla de entrenamiento a partir de los datos preparados y las reglas definidas para construir la granularidad empleado-semana. |
| `02_Preparacion_Datos_IT.ipynb` | Preparación, revisión y transformación de los datos para su utilización en el proceso de modelado. |
| `03_Modelado_IT.ipynb` | Desarrollo y comparación inicial de modelos de aprendizaje automático para la predicción de IT. |
| `04_Optimizacion_IT_FINAL.ipynb` | Ajuste y optimización del modelo seleccionado y revisión de sus resultados. |
| `05_Entrenamiento_Final_IT.ipynb` | Entrenamiento final y generación del artefacto del modelo elegido. |

> Los nombres y el orden reflejan la organización actual del repositorio. Los detalles concretos de entradas, salidas y dependencias deben consultarse en cada archivo.

## Objetivo

Desarrollar y evaluar un enfoque de aprendizaje automático capaz de predecir la incapacidad temporal a partir de información histórica y variables laborales, con el propósito de explorar su utilidad como apoyo analítico para la planificación de los recursos humanos.

## Enfoque metodológico

El trabajo sigue, a grandes rasgos, estas etapas:

1. **Construcción de la base de entrenamiento:** organización de los registros con granularidad empleado-semana.
2. **Preparación de datos:** revisión de calidad, tratamiento de variables y preparación de las características predictoras.
3. **Modelado:** entrenamiento y comparación de algoritmos de regresión.
4. **Optimización:** ajuste de hiperparámetros y evaluación de alternativas.
5. **Entrenamiento final:** generación del modelo final y conservación de los artefactos necesarios para su uso posterior.

El modelo final del proyecto utiliza CatBoost para estimar cuatro horizontes semanales futuros de IT. Las métricas y los resultados experimentales deben consultarse en la memoria del TFM y en los archivos de resultados correspondientes.

## Autoría

Trabajo Fin de Máster desarrollado en el marco del Máster Universitario en Big Data.
