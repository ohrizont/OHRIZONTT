📊 Predicción del precio de portátiles (Kaggle)
Descripción del proyecto

Este proyecto aborda el problema de predicción del precio de portátiles a partir de sus especificaciones técnicas, utilizando modelos de machine learning aplicados a datos tabulares.

El dataset original procede de una competición antigua de Kaggle y contiene información sobre:

hardware (CPU, GPU, RAM, almacenamiento),

características físicas (peso, tamaño de pantalla),

tipo de equipo, sistema operativo y marca.

El objetivo no es solo maximizar la métrica, sino entender qué variables explican el precio, cómo transformarlas correctamente y comparar distintos enfoques de modelado de forma rigurosa.

Objetivo

Construir modelos predictivos estables y defendibles.

Comparar enfoques lineales y no lineales.

Analizar el impacto de la ingeniería de variables en el rendimiento.

Evaluar correctamente la generalización (sin fuga de información).

Mostrar cómo la introducción de benchmarks reales de CPU cambia radicalmente el problema.

Dataset

Dataset original: Kaggle (no incluido en el repositorio).

El repositorio no contiene los datos por motivos de licencia.

Para reproducir el proyecto:

Descargar el dataset desde Kaggle.

Colocar los ficheros localmente manteniendo los nombres originales.

Proceso de trabajo
1️⃣ Exploración y limpieza

Análisis de distribuciones de precios.

Identificación de variables redundantes o poco informativas.

Tratamiento de valores faltantes.

Normalización semántica de variables categóricas.

2️⃣ Ingeniería de variables

Transformación de CPU y GPU originales en:

variables categóricas de segmento / tier.

Selección progresiva de variables:

RAM, peso, pulgadas, almacenamiento, tipo de equipo, marca, SO.

Eliminación de variables irrelevantes o redundantes (por ejemplo, resolución exacta de pantalla).

3️⃣ Regresión lineal (modelo explicativo)

Construcción de varios modelos lineales sucesivos.

Eliminación iterativa de variables:

colineales,

sin significación,

sin impacto real en MAE.

Resultado:

modelo interpretable,

buen baseline,

MAE ≈ 200 €.

4️⃣ Random Forest (modelo predictivo)

Uso de variables “crudas” + categóricas ricas.

One-Hot Encoding controlado.

Ajuste progresivo del modelo:

número de árboles,

profundidad,

tamaño mínimo de hoja.

Validación:

hold-out 80/20 aleatorio,

comparación train vs test,

análisis de complejidad del bosque.

Resultado:

MAE ≈ 170 €,

R² ≈ 0.84,

modelo estable y defendible.

5️⃣ Importancia de variables

Uso de importancia por permutación (no sesgada).

Identificación de:

variables dominantes (RAM, peso, CPU/GPU),

variables redundantes (por ejemplo, peso vs tipo de equipo).

Simplificación del modelo sin pérdida de rendimiento.

6️⃣ Benchmarks reales de CPU (extensión)

Enriquecimiento del dataset cruzando el modelo exacto de CPU con:

benchmarks externos (PassMark).

Sustitución de proxies (cpu_tier) por una variable continua (cpu_rating).

Resultado:

MAE ≈ 15–20 €,

R² ≈ 0.99.

Interpretación:

no hay fuga de datos,

el problema se vuelve casi determinista,

representa un upper bound teórico del rendimiento del modelo.

Modelos finales

Se mantienen dos modelos finales, con objetivos distintos:

🔹 Modelo A — Sin benchmarks (generalista)

Variables técnicas y categóricas.

MAE ≈ 170 €.

Más robusto para escenarios reales incompletos.

Recomendado como modelo general.

🔹 Modelo B — Con benchmarks de CPU

Introduce rendimiento real del procesador.

MAE ≈ 15–20 €.

Representa el límite superior del problema con esta información.

Útil como modelo de referencia técnica.

Validación

Split aleatorio 80 % / 20 %.

Evaluación fuera de muestra.

Comparación train vs test para detectar sobreajuste.

Análisis adicional con OOB (cuando aplica).

Resultados consistentes y reproducibles.

Tecnologías usadas

Python

pandas, numpy

scikit-learn

matplotlib, seaborn

Ver requirements.txt para detalles.

Conclusión

El proyecto demuestra que:

la ingeniería de variables es más importante que el algoritmo,

Random Forest supera claramente a modelos lineales en este problema,

la introducción de benchmarks reales transforma por completo la capacidad predictiva,

el modelo final es técnicamente sólido y defendible a nivel profesional.
