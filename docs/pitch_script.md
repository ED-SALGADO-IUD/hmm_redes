# Guion de pitch de investigación

**Duración estimada:** 2 minutos y 30 segundos. **Idioma:** español.  
**Título:** Detección de regímenes de calidad de red mediante un HMM reproducible.

## Guion con indicaciones visuales

**0:00–0:15 | Gancho**  
¿Una conexión está estable, degradándose o a punto de caer? Un único ping puede verse normal incluso cuando la red ya atraviesa una degradación. El reto es inferir el estado que no observamos directamente a partir de señales imperfectas.

**0:15–0:40 | Motivación e impacto**  
La latencia y la pérdida de paquetes afectan videollamadas, aplicaciones empresariales y servicios en la nube. La literatura ha aplicado modelos de Markov y HMM a pérdidas Wi-Fi, tráfico y disponibilidad de redes. Sin embargo, una exactitud promedio puede ocultar que el modelo no detecta precisamente las caídas, que son el evento menos frecuente.

**0:40–1:10 | Solución propuesta**  
Propongo un HMM categórico con tres estados ocultos: estable, degradado y caída. Las observaciones combinan nueve categorías formadas por tres niveles de latencia y tres niveles de pérdida de paquetes. La matriz A representa la persistencia y transición entre estados; la matriz B, la probabilidad de observar cada combinación; y pi, la condición inicial. Viterbi estima la secuencia más probable.

**1:10–1:35 | Metodología**  
Ejecuté 400 réplicas de 300 pasos en cuatro escenarios reproducibles: referencia, estados más persistentes, emisiones más parecidas y una condición inicial adversa. Se fijaron semillas y se calcularon exactitud, F1 macro, sensibilidad a caídas, falsas alarmas e intervalos de confianza del 95 %.

**1:35–2:05 | Resultado**  
En el escenario base, la exactitud fue 88,3 %, pero la sensibilidad media a caídas fue solo 42,0 %. Cuando las emisiones se hicieron más similares, la exactitud bajó a 75,0 % y la sensibilidad a caídas a 0,4 %. En contraste, con estados más persistentes, el F1 macro subió de 70,6 % a 77,2 %. Esto demuestra por qué debemos evaluar cada clase y no solo la exactitud global.

**2:05–2:30 | Aporte y cierre**  
El aporte es un experimento transparente que muestra cuándo un HMM puede ser útil y cuándo su alerta de caída deja de ser confiable. El código, los parámetros, las semillas y los resultados están disponibles para repetir el análisis. El siguiente paso es calibrarlo y validarlo con telemetría real antes de usarlo en producción.

## Apoyos visuales sugeridos

1. Un indicador de conexión con tres estados: estable, degradado y caída.
2. Esquema HMM: estado oculto → observación conjunta de latencia y pérdida.
3. Tabla/matriz A y matriz B de referencia.
4. Gráfico de sensibilidad: F1 macro y sensibilidad a caídas por escenario.
5. Cierre con resultados, repositorio y limitación: «simulación sintética; falta validación en red real».
