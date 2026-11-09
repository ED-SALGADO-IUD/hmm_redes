# HMM de calidad de conectividad

Proyecto reproducible para estudiar cómo un modelo oculto de Markov (HMM) puede inferir tres regímenes de calidad de red a partir de observaciones categóricas conjuntas de latencia y pérdida de paquetes.

## Alcance y datos

El experimento genera datos sintéticos mediante un HMM especificado en `data/model_parameters.json`. No contiene mediciones de una red real y no debe interpretarse como una evaluación de un proveedor o dispositivo. Los estados sintéticos conocidos permiten medir el error de decodificación de Viterbi bajo cuatro condiciones controladas.

## Estructura

- `src/simulate_hmm.py`: simulador, algoritmo Forward en escala logarítmica, Viterbi, métricas y creación de figuras.
- `data/model_parameters.json`: estados, vocabulario, matrices, semillas y versiones del entorno.
- `data/replicate_metrics.csv`: métricas de cada una de las 1.600 réplicas.
- `data/summary_metrics.csv`: media, desviación estándar, mínimo, máximo e IC del 95 % de la media.
- `figures/`: figuras en PNG a 300 dpi.
- `docs/`: manuscrito y guion del pitch.

## Requisitos

Python 3.12 o compatible; NumPy, pandas y Matplotlib. Versiones usadas en la ejecución original: Python 3.12.14, NumPy 2.3.5, pandas 2.2.3 y Matplotlib 3.10.8.

Instalación opcional en un entorno virtual:

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install numpy pandas matplotlib
```

## Ejecución

Desde la carpeta del proyecto:

```bash
python src/simulate_hmm.py
```

La ejecución produce cuatro escenarios, 400 réplicas por escenario, secuencias de 300 pasos, 1.600 observaciones de métricas, tres gráficos analíticos de 300 dpi y los parámetros serializados. La figura de flujo incluida en `figures/fig4_workflow.png` completa la documentación visual del método. La semilla base es `20261109`; cada réplica usa una semilla derivada de forma determinista. Al ejecutar de nuevo en el mismo entorno deben obtenerse los mismos CSV y figuras (exceptuando metadatos internos de imagen).

## Escenarios

1. `S1_baseline`: matriz de transición y emisión de referencia.
2. `S2_persistent_states`: mayor permanencia en cada régimen.
3. `S3_overlapping_emissions`: mezcla de 55 % de distribución uniforme en las emisiones para representar señales poco discriminantes.
4. `S4_initial_state_stress`: distribución inicial desplazada desde el estado estable hacia degradado/caída.

## Interpretación correcta

El sistema mide la inferencia de estados generados por el propio modelo. Los intervalos de confianza describen la variabilidad Monte Carlo entre réplicas, no la incertidumbre de una población real ni el rendimiento en producción. Para llevar el modelo a una red real se necesitan series etiquetadas o criterios operativos validados, ajuste de parámetros, evaluación temporal fuera de muestra y calibración de umbrales.
