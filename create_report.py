from pathlib import Path
import json
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "docs" / "Articulo_HMM_Calidad_Red.docx"
PARAMS = json.loads((ROOT / "data/model_parameters.json").read_text(encoding="utf-8"))
SUMMARY = pd.read_csv(ROOT / "data/summary_metrics.csv")


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "5")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "D9D9D9")
        borders.append(element)
    tbl_pr.append(borders)


def style_run(run, size=10.5, bold=False, italic=False, color="000000", font="Aptos"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def add_body(doc, text, first_bold=None):
    # Use readable Unicode subscripts instead of raw LaTeX-style underscores.
    for raw, pretty in (
        ("X_{t−1}", "Xₜ₋₁"), ("X_{1:T}", "X₁:ₜ"), ("O_{1:T}", "O₁:ₜ"),
        ("X_{1:t}", "X₁:ₜ"), ("O_{1:t−1}", "O₁:ₜ₋₁"),
        ("X_{t−1},X_t", "Xₜ₋₁,Xₜ"), ("X_t", "Xₜ"), ("O_t", "Oₜ"),
        ("X_1", "X₁"), ("O_1", "O₁"), ("x_{1:T}", "x₁:ₜ"),
        ("o_{1:T}", "o₁:ₜ"), ("x1", "x₁"), ("o1", "o₁"),
        ("π_i", "πᵢ"), ("a_ij", "aᵢⱼ"), ("A_{i·}", "Aᵢ·"),
        ("a_ii", "aᵢᵢ"), ("a_ii)", "aᵢᵢ)"),
        ("B_{i·}", "Bᵢ·"), ("b_i(k)", "bᵢ(k)"), ("α_t", "αₜ"),
        ("α_1(i)", "α₁(i)"), ("α_{t−1}(i)", "αₜ₋₁(i)"),
        ("δ_1(i)", "δ₁(i)"), ("δ_t(j)", "δₜ(j)"), ("δ_{t−1}(i)", "δₜ₋₁(i)"),
        ("a_{x_{t−1},x_t}", "aₓₜ₋₁,ₓₜ"), ("b_{x_t}(o_t)", "bₓₜ(oₜ)"),
        ("P(x_{1:T},o_{1:T})", "P(x₁:ₜ,o₁:ₜ)"), ("π_{x1}", "πₓ₁"),
        ("b_{x1}(o1)", "bₓ₁(o₁)"), ("s_r", "sᵣ"), ("X̂_{1:T}", "X̂₁:ₜ"),
    ):
        text = text.replace(raw, pretty)
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if first_bold and text.startswith(first_bold):
        r = p.add_run(first_bold)
        style_run(r, bold=True)
        r = p.add_run(text[len(first_bold):])
        style_run(r)
    else:
        r = p.add_run(text)
        style_run(r)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text)
    style_run(r, size=14 if level == 1 else 12, bold=True, font="Aptos Display")
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run(text), size=9, bold=True)


def add_table(doc, headers, rows, widths=None, font_size=8.2):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_borders(table)
    header = table.rows[0]
    set_repeat_table_header(header)
    for j, text in enumerate(headers):
        cell = header.cells[j]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, "1F4E79")
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style_run(p.add_run(str(text)), size=font_size, bold=True, color="FFFFFF")
        if widths: cell.width = Inches(widths[j])
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        # Keep each table row intact when Word paginates a long table.
        tr_pr = table.rows[-1]._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        for j, value in enumerate(row):
            cells[j].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cells[j])
            if i % 2 == 1: set_cell_shading(cells[j], "EEF4F8")
            p = cells[j].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 or (len(headers) > 3 and j == 2) else WD_ALIGN_PARAGRAPH.CENTER
            style_run(p.add_run(str(value)), size=font_size)
            if widths: cells[j].width = Inches(widths[j])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def fmt_metric(metric, scenario):
    r = SUMMARY[(SUMMARY.scenario == scenario) & (SUMMARY.metric == metric)].iloc[0]
    return [f"{r['mean']:.3f} ± {r['sd']:.3f}",
            f"{r['min']:.3f}–{r['max']:.3f}",
            f"[{r['ci95_low']:.3f}, {r['ci95_high']:.3f}]"]


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin, sec.bottom_margin = Inches(0.72), Inches(0.72)
sec.left_margin, sec.right_margin = Inches(0.78), Inches(0.78)

normal = doc.styles["Normal"]
normal.font.name = "Aptos"
normal.font.size = Pt(10.5)
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
for stylename in ("Heading 1", "Heading 2", "Heading 3"):
    st = doc.styles[stylename]
    st.font.name = "Aptos Display"
    st.font.color.rgb = RGBColor(0, 0, 0)

# Header and page numbering
header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
style_run(header.add_run("MODELOS ESTOCÁSTICOS APLICADOS A REDES"), size=8, color="555555")
footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = footer.add_run("Página ")
style_run(r, size=8, color="555555")
fld = OxmlElement("w:fldSimple")
fld.set(qn("w:instr"), "PAGE")
footer._p.append(fld)

p = doc.add_paragraph(style="Title")
# Remove the default decorative rule from Word's Title style.
for node in list(doc.styles["Title"]._element.iter(qn("w:pBdr"))):
    node.getparent().remove(node)
if p._p.pPr is not None:
    for node in list(p._p.pPr.iter(qn("w:pBdr"))):
        node.getparent().remove(node)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
style_run(p.add_run("Evaluación reproducible de un modelo oculto de Markov para inferir regímenes de calidad de conectividad"), size=18, bold=True, font="Aptos Display")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(16)
style_run(p.add_run("Análisis de sensibilidad mediante simulación de latencia y pérdida de paquetes"), size=12, italic=True)

add_heading(doc, "Resumen", 1)
abstract = ("La degradación de una conexión puede desarrollarse antes de que se observe una interrupción completa, mientras que la latencia y la pérdida de paquetes ofrecen señales imperfectas del estado subyacente. Este trabajo diseña y evalúa un modelo oculto de Markov (HMM) categórico para inferir tres regímenes latentes: estable, degradado y caída. Las observaciones se forman cruzando tres niveles de latencia con tres niveles de pérdida de paquetes, lo que produce un alfabeto de nueve categorías. Se especifican la distribución inicial π, la matriz de transición A y la matriz de emisión B; la inferencia se realiza con Viterbi y la verosimilitud de las secuencias con el algoritmo Forward escalado. La evaluación es una simulación controlada, no una medición de una red real: se generaron 400 secuencias de 300 pasos en cada uno de cuatro escenarios, con semillas deterministas. En el escenario base, la exactitud media fue 88,3 % (IC 95 %: 87,97–88,59), mientras que la sensibilidad media a caídas fue 42,0 %. Al aumentar la superposición entre emisiones, la exactitud descendió a 75,0 % y la sensibilidad a caídas a 0,4 %. Una mayor persistencia de los estados elevó el F1 macro de 70,6 % a 77,2 %. El resultado central es que la exactitud global puede ocultar un rendimiento insuficiente en el estado crítico; por ello, la evaluación debe reportar métricas por clase y sensibilidad a los parámetros. El código, las matrices y los resultados crudos acompañan el manuscrito para permitir su reproducción.")
add_body(doc, abstract)
add_body(doc, "Palabras clave: modelo oculto de Markov, calidad de servicio, latencia, pérdida de paquetes, inferencia de estados, simulación reproducible.", first_bold="Palabras clave:")

add_heading(doc, "1. Introducción", 1)
for t in [
"La calidad de una conexión de red es temporal y no siempre puede resumirse con una observación aislada. Una demora elevada o una pérdida puntual pueden ser transitorias; una sucesión persistente de observaciones anómalas puede indicar congestión, interferencia o una falla en curso. Esta diferencia importa para la gestión de redes porque una alerta tardía reduce el tiempo disponible para diagnosticar el problema, mientras que una alerta excesivamente sensible puede generar falsas alarmas.",
"Los modelos de Markov representan sistemas cuyo estado evoluciona probabilísticamente. En una cadena visible, el estado se observa directamente; en un HMM, el estado relevante es latente y se infiere a partir de mediciones observables. Este último supuesto es adecuado para la calidad de red: el estado operativo real no siempre se registra como etiqueta, pero sí se observan indicadores como latencia, pérdida, ocupación del canal o métricas de señal. Trabajos recientes aplican HMM a pérdidas Wi-Fi, desempeño de celdas móviles, matrices de tráfico y detección de amenazas [1], [2], [7], [9], [12].",
"El problema se delimita a la inferencia secuencial de tres regímenes de calidad a partir de observaciones discretizadas de latencia y pérdida. No se estima causalidad, no se diagnostica el componente físico responsable y no se pretende certificar el rendimiento de una red real. El propósito es mostrar cómo las distribuciones de probabilidad, las matrices del HMM y la persistencia temporal afectan la inferencia, con experimentos enteramente reproducibles.",
"La pregunta general es cómo cambia la identificación de estados ocultos cuando varían la dinámica de transición, la separabilidad de las señales y las condiciones iniciales. El aporte consiste en una evaluación controlada que hace explícita la diferencia entre exactitud global y capacidad para reconocer la clase crítica, además de publicar los parámetros, semillas, datos sintéticos y scripts de análisis. No se propone un algoritmo HMM nuevo; la originalidad del ejercicio es el diseño comparativo, transparente y reproducible aplicado a esta delimitación.",
"Las preguntas de investigación son: P1) ¿Cómo puede representarse la calidad de conectividad mediante un HMM visible en sus observaciones y oculto en los regímenes? P2) ¿Cómo influyen los parámetros de transición y emisión en las simulaciones? P3) ¿Qué tan reproducible y sensible es el modelo frente a cambios en π, A y B? El objetivo general es implementar y evaluar el HMM mediante simulación Monte Carlo. Los objetivos específicos son formalizar sus distribuciones, ejecutar Forward y Viterbi, comparar cuatro escenarios y comunicar limitaciones para una eventual validación con telemetría real."
]: add_body(doc, t)

add_heading(doc, "2. Estado del arte", 1)
add_body(doc, "Se realizó una revisión dirigida de publicaciones académicas de 2021–2025 localizadas mediante consultas de términos relacionados con HMM, Markov, packet loss, network quality, reliability y traffic prediction, y consulta de metadatos en portales de editoriales, IEEE y repositorios universitarios. Se seleccionaron 12 artículos recientes con DOI o metadatos bibliográficos verificables, complementados con referencias clásicas. Esta revisión es una síntesis narrativa dirigida, no un protocolo PRISMA ni una extracción exhaustiva de Scopus y Web of Science; el cuartil y la indexación exacta deben confirmarse en el acceso institucional para el año y categoría exigidos por el curso.")
add_body(doc, "La literatura reciente cubre cuatro líneas. Primero, la caracterización de señales de red con HMM: Silva y Pedroso integran variables Wi-Fi de capa cruzada y modelan ráfagas de pérdidas usando datos de una red real [1]; Omer et al. infieren degradación de celdas LTE a partir de KPI de operador [9]. Segundo, el modelado de tráfico y su predicción, incluida la reconstrucción de matrices con valores ausentes mediante una combinación de autoencoders y HMM [7]. Tercero, los modelos de disponibilidad y confiabilidad de redes mediante cadenas de Markov [2], [10], [11]. Cuarto, el desarrollo metodológico para inferencia, sensibilidad y detección: se estudian la sensibilidad de modelos HMM de confiabilidad [6], el uso de HMM en mantenimiento [5], la paralelización de Forward/Viterbi [4] y la sensibilidad de Baum–Welch a la inicialización en detección de ataques [3], [8], [12].")
add_caption(doc, "Tabla I. Estudios recientes relacionados con modelos de Markov en redes y sistemas observados parcialmente.")
rows = [
["Silva y Pedroso (2022) [1]", "HMM de pérdidas Wi-Fi con SNR y ocupación", "Muestras reales de red; comparación con modelos de pérdida", "Motiva observaciones conjuntas; no evalúa nuestra matriz sintética."],
["Lu et al. (2022) [2]", "Cadena de Markov para disponibilidad de servicios de comunicación", "Simulación de demora y disponibilidad", "Apoya el uso de transiciones para métricas de red."],
["Zhang et al. (2022) [3]", "HMM para secuencias de ataques; inicialización semántica de Baum–Welch", "DARPA 2000, DEFCON21, ISCXIDS 2012", "Evidencia sensibilidad a inicialización y uso de inferencia latente."],
["Hassan et al. (2021) [4]", "Algoritmos paralelos Forward y Viterbi", "Comparación computacional en GPU", "Fundamenta algoritmos clásicos; este trabajo prioriza legibilidad."],
["Gámiz et al. (2023) [5]", "HMM para confiabilidad y mantenimiento", "Teoría de estimación y simulación", "Relaciona estados ocultos con fallas no observadas."],
["Gámiz et al. (2023) [6]", "HMM de doble cadena y análisis de sensibilidad", "Simulación y caso de bomba de agua", "Informa la evaluación de falsos positivos/negativos."],
["Sacco et al. (2023) [7]", "Autoencoders adversariales + HMM para matrices de tráfico", "Conjuntos reales y banco de pruebas", "Demuestra uso para completar y predecir tráfico."],
["Omer et al. (2024) [9]", "HMM para predecir degradación de celdas móviles", "KPI reales de operador", "Aplicación directa de regímenes de degradación; difiere en datos y estados."],
["Jalili et al. (2024) [10]; Heidari et al. (2024) [11]", "Cadenas de Markov para tolerancia a fallas y disponibilidad WSN", "Modelos de redundancia y fallas", "Apoyan la interpretación de estados y confiabilidad."],
["Mahboubi et al. (2025) [12]", "HMM continuo para decodificación de beaconing", "14 conjuntos de referencia y un conjunto sintético", "Ejemplo actual de inferencia de estados latentes en tráfico de red."],
]
add_table(doc, ["Estudio", "Modelo y problema", "Base de evaluación", "Aporte / diferencia"], rows, widths=[1.25, 1.65, 1.48, 2.15], font_size=7.4)
add_body(doc, "El vacío práctico para este ejercicio no es la ausencia de HMM en telecomunicaciones, sino la dificultad de comparar, de manera didáctica y reproducible, cómo la persistencia de estados y la ambigüedad de las observaciones cambian la detección de una clase crítica rara. Muchos trabajos ajustan sus modelos a datos operativos y optimizan un caso específico; el experimento aquí fija parámetros conocidos para aislar el efecto de las hipótesis. Su limitación deliberada es que no demuestra capacidad predictiva en producción y no sustituye los estudios empíricos anteriores.")

add_heading(doc, "3. Marco teórico", 1)
add_heading(doc, "3.1. Variables aleatorias y distribuciones", 2)
for t in [
"Sea Xₜ el estado oculto en el instante discreto t y Oₜ la observación categórica. Xₜ toma valores en S = {E, D, C}, donde E = estable, D = degradado y C = caída. Oₜ pertenece a V, el alfabeto de nueve combinaciones entre nivel de latencia y nivel de pérdida. Tanto la transición de estado como la emisión se modelan mediante distribuciones categóricas, equivalentes a una distribución multinomial de una sola prueba: X₁ ~ Cat(π), Xₜ | Xₜ₋₁=i ~ Cat(Aᵢ·) y Oₜ | Xₜ=i ~ Cat(Bᵢ·).",
"Las observaciones se obtendrían de medidas continuas y luego se discretizarían con umbrales operativos explícitos. En este experimento, latencia se clasifica como baja (<80 ms), moderada (80–200 ms) o alta (>200 ms), y pérdida de paquetes como baja (<1 %), moderada (1–5 %) o alta (>5 %). Los límites son supuestos del diseño de simulación, no umbrales universales de calidad; una aplicación real debe definirlos con requisitos de servicio y datos del entorno. Una extensión podría reemplazar las emisiones categóricas por distribuciones continuas, como lognormal para latencia positiva y beta para pérdida fraccional, o por una emisión multivariada que represente dependencia entre ambas métricas."
]: add_body(doc, t)
add_heading(doc, "3.2. Cadena de Markov y HMM", 2)
for t in [
"Una cadena de primer orden satisface la propiedad de Markov: P(Xₜ | X₁,…,Xₜ₋₁) = P(Xₜ | Xₜ₋₁). La matriz A = [aᵢⱼ] contiene aᵢⱼ = P(Xₜ=j | Xₜ₋₁=i), con aᵢⱼ ≥ 0 y suma de fila igual a uno. La distribución inicial πᵢ = P(X₁=i) también suma uno. Una cadena finita regular puede tener una distribución estacionaria μ que cumple μ = μA; esta distribución describe el comportamiento de largo plazo y no reemplaza π para iniciar una secuencia finita.",
"Un HMM añade observaciones cuyo valor depende del estado actual: bᵢ(k)=P(Oₜ=k | Xₜ=i), donde cada fila de B suma uno. Se asume independencia condicional de las observaciones dado el estado: P(Oₜ | X₁:ₜ,O₁:ₜ₋₁)=P(Oₜ | Xₜ). La especificación completa es λ=(S,V,A,B,π). La probabilidad conjunta de una trayectoria de estados y observaciones es P(x₁:ₜ,o₁:ₜ) = πₓ₁ bₓ₁(o₁) ∏ₜ₌₂ᵀ aₓₜ₋₁,ₓₜ bₓₜ(oₜ)."
]: add_body(doc, t)
add_caption(doc, "Tabla II. Definición del alfabeto de observaciones.")
obs_rows = [[f"{lat}-{loss}", lat_full, loss_full] for lat, lat_full in [("L", "Latencia baja"), ("M", "Latencia moderada"), ("H", "Latencia alta")] for loss, loss_full in [("low", "Pérdida baja"), ("moderate", "Pérdida moderada"), ("high", "Pérdida alta")]]
add_table(doc, ["Símbolo", "Nivel de latencia", "Nivel de pérdida"], obs_rows, widths=[1.2, 2.3, 2.3], font_size=8.2)
add_heading(doc, "3.3. Problemas clásicos de inferencia", 2)
for t in [
"El algoritmo Forward calcula P(O₁:ₜ|λ) sumando sobre todas las trayectorias latentes. Con αₜ(i)=P(O₁:ₜ,Xₜ=i), la recursión es α₁(i)=πᵢ bᵢ(O₁) y αₜ(j)=bⱼ(Oₜ) Σᵢ αₜ₋₁(i)aᵢⱼ. La implementación usa normalización por paso para evitar subdesbordamiento numérico y acumula la log-verosimilitud.",
"Viterbi encuentra la trayectoria más probable de estados: δ₁(i)=πᵢ bᵢ(O₁); δₜ(j)=bⱼ(Oₜ) maxᵢ[δₜ₋₁(i)aᵢⱼ], almacenando el argumento máximo para reconstruir el camino. En el programa, las operaciones se hacen en logaritmos. Baum–Welch, una forma del algoritmo EM, estima parámetros desconocidos maximizando iterativamente la verosimilitud mediante probabilidades suavizadas de estados y transiciones. El experimento no ejecuta Baum–Welch porque A, B y π se fijan antes de generar los datos: estimarlos en los mismos datos simulados cambiaría la pregunta de sensibilidad y podría introducir problemas de identificabilidad y óptimos locales. Se documenta como extensión y se mantiene el código enfocado en Forward y Viterbi [3]–[5], [13], [14]."
]: add_body(doc, t)

add_heading(doc, "4. Metodología", 1)
add_heading(doc, "4.1. Especificación del modelo", 2)
add_body(doc, "Los estados son S={E,D,C}; las nueve observaciones combinan L/M/H con pérdida baja/moderada/alta. En el orden E,D,C, la distribución inicial de referencia es π=(0,80; 0,18; 0,02). La matriz de transición de referencia y la matriz de emisión se muestran en la Fig. 1. Cada matriz es estocástica por filas. Las probabilidades son parámetros explícitos de un escenario sintético, no estimaciones de una red observada.")
add_caption(doc, "Figura 1. Matrices de transición y emisión del escenario de referencia.")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(str(ROOT / "figures/fig3_reference_matrices.png"), width=Inches(6.65))
add_body(doc, "La matriz A de referencia tiene probabilidades diagonales 0,94, 0,72 y 0,65; por ello, las duraciones medias geométricas de las rachas por estado son 1/(1−a_ii), equivalentes aproximadamente a 16,7, 3,6 y 2,9 intervalos. La distribución estacionaria calculada es μ=(0,748; 0,197; 0,056), mientras que π=(0,80; 0,18; 0,02) solo determina la distribución en el primer instante. Las emisiones se diseñaron para que las categorías de latencia y pérdida altas sean más probables en degradado o caída, sin volverlas deterministas.")

add_heading(doc, "4.2. Escenarios experimentales", 2)
scenario_rows = [
["S1 Baseline", "π=(0,80; 0,18; 0,02); A y B de referencia", "Comparación base"],
["S2 Persistent", "A cambia a diagonales 0,97; 0,84; 0,80", "Aísla el efecto de mayor permanencia"],
["S3 Overlap", "B' = 0,45B + 0,55U; U uniforme en 9 observaciones", "Evalúa la pérdida de separabilidad"],
["S4 Initial stress", "π=(0,45; 0,40; 0,15); A y B de referencia", "Evalúa sensibilidad al inicio"],
]
add_caption(doc, "Tabla III. Escenarios de simulación.")
add_table(doc, ["Escenario", "Cambio de parámetros", "Propósito"], scenario_rows, widths=[1.3, 3.2, 2.0], font_size=8.0)

add_heading(doc, "4.3. Algoritmo y flujo reproducible", 2)
for t in [
"Cada réplica inicia X₁ según π; después, en cada instante, muestrea Xₜ desde la fila correspondiente de A y Oₜ desde la fila de B asociada al estado. Se decodifica O₁:ₜ con Viterbi, se calcula la verosimilitud con Forward y se comparan estados estimados y verdaderos. El proceso se repite 400 veces por escenario, con T=300, usando una semilla base fija y una semilla derivada única para cada réplica.",
"Pseudocódigo: para cada escenario, validar que π, A y B sumen uno por fila; para r=1,…,400, fijar semilla sᵣ; muestrear X₁ y O₁; para t=2,…,300, muestrear Xₜ|Xₜ₋₁ y Oₜ|Xₜ; inferir X̂₁:ₜ por Viterbi; calcular métricas; agregar resultado. Finalmente, resumir media, desviación estándar, mínimo, máximo e intervalo de confianza de la media al 95 % sobre las 400 réplicas."
]: add_body(doc, t)
add_caption(doc, "Figura 2. Flujo del experimento reproducible.")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(str(ROOT / "figures/fig4_workflow.png"), width=Inches(6.5))

add_heading(doc, "4.4. Métricas y análisis estadístico", 2)
for t in [
"Se calculan exactitud global, F1 macro (promedio no ponderado de F1 para los tres estados), sensibilidad de la clase caída y tasa de falsa alarma entre intervalos verdaderamente estables que se decodifican como degradados o caída. El F1 macro evita que el estado estable, más prevalente, domine por completo el promedio. La sensibilidad a caída es particularmente relevante porque una clase rara puede quedar sin detección aunque la exactitud parezca alta.",
"Para cada métrica y escenario se resumen los 400 valores de réplica con media, desviación estándar muestral, mínimo y máximo. El IC del 95 % de la media se obtiene con la aproximación normal media ± 1,96·s/√n. Es un intervalo sobre el promedio de las réplicas Monte Carlo bajo los parámetros fijados; no es un intervalo de confianza de rendimiento en una red real ni incorpora incertidumbre de los parámetros. El entorno de ejecución reportado fue Python 3.12.14, NumPy 2.3.5, pandas 2.2.3 y Matplotlib 3.10.8."
]: add_body(doc, t)

add_heading(doc, "5. Resultados", 1)
add_body(doc, "El experimento completó 1.600 réplicas (400 por escenario), cada una con 300 observaciones. El archivo de métricas crudas contiene una fila por réplica y permite regenerar las tablas. La Fig. 3 muestra la matriz de confusión normalizada del escenario base; la Fig. 4 compara F1 macro y sensibilidad a caídas. Las figuras se generan en PNG con resolución de 300 dpi.")
add_caption(doc, "Figura 3. Matriz de confusión normalizada del escenario base.")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(str(ROOT / "figures/fig1_baseline_confusion.png"), width=Inches(4.8))
add_caption(doc, "Figura 4. Sensibilidad de la inferencia a cambios de parámetros y distribuciones.")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(str(ROOT / "figures/fig2_scenario_sensitivity.png"), width=Inches(6.6))

metrics_labels = [("accuracy", "Exactitud"), ("macro_f1", "F1 macro"), ("outage_recall", "Sensibilidad a caída"), ("stable_false_alarm_rate", "Falsa alarma en estable")]
scenarios = [("S1_baseline", "S1 Base"), ("S2_persistent_states", "S2 Persistente"), ("S3_overlapping_emissions", "S3 Emisiones solapadas"), ("S4_initial_state_stress", "S4 Inicio alterado")]
for table_no, (metric, label) in enumerate(metrics_labels, start=4):
    rows = []
    for scenario, display in scenarios:
        mean_sd, min_max, ci = fmt_metric(metric, scenario)
        rows.append([display, mean_sd, min_max, ci])
    add_caption(doc, f"Tabla {table_no}. {label}: resumen entre réplicas (n=400 por escenario).")
    add_table(doc, ["Escenario", "Media ± DE", "Mín.–máx.", "IC 95 % de la media"], rows, widths=[1.7, 1.45, 1.25, 1.8], font_size=8.0)

add_body(doc, "En S1, la exactitud media fue 0,883, pero el F1 macro fue 0,706 y la sensibilidad a caída 0,420. La diferencia muestra el efecto de la prevalencia y la dificultad de separar estados: la exactitud resume todos los intervalos, mientras que el F1 macro asigna el mismo peso a cada clase. El intervalo de confianza de la sensibilidad a caída fue más ancho (0,394–0,446) y su desviación estándar fue 0,265, lo que refleja el bajo número de episodios de caída en muchas secuencias.")
add_body(doc, "En S2, la mayor persistencia se asoció con mayor exactitud (0,924) y F1 macro (0,772), además de una sensibilidad a caída de 0,521. La explicación es que una secuencia más persistente aporta contexto temporal que ayuda a compensar observaciones ambiguas. Este resultado depende de las matrices especificadas y no implica que hacer persistentes las transiciones mejore automáticamente cualquier red.")
add_body(doc, "En S3, mezclar 55 % de uniformidad en B redujo la exactitud a 0,750 y el F1 macro a 0,293. La sensibilidad a caída fue 0,004, aunque la falsa alarma en el estado estable también fue baja (0,001). El decodificador optó, en la práctica, por estados frecuentes y no identificó el estado raro. Esta es la evidencia cuantitativa principal de que una exactitud moderada puede coexistir con una alerta crítica prácticamente inoperante.")
add_body(doc, "En S4, cambiar π produjo resultados cercanos a S1 para secuencias de 300 pasos: exactitud 0,884 y F1 macro 0,707. La dinámica de A domina progresivamente las condiciones iniciales en secuencias largas de esta cadena. No debe generalizarse a ventanas cortas, cadenas con transiciones más lentas o estimación de parámetros; en esos contextos π puede tener un efecto mayor.")

add_heading(doc, "6. Discusión", 1)
for t in [
"Los resultados responden a la pregunta de cómo las probabilidades influyen en la simulación: la persistencia temporal puede facilitar la decodificación, mientras que la superposición de emisiones puede destruir la capacidad de identificar la clase crítica. La matriz B es central para la discriminación; si las distribuciones condicionadas por estado son parecidas, más datos temporales no garantizan una detección útil. La matriz A aporta estructura temporal, pero puede favorecer estados de larga duración y dejar eventos cortos sin reconocer.",
"La literatura de pérdidas Wi-Fi muestra que combinar indicadores de red y modelar su dependencia temporal puede mejorar el ajuste respecto de modelos más simples [1]. Los estudios sobre degradación móvil también usan KPI reales para predecir estados de severidad [9]. Los modelos de confiabilidad y mantenimiento enfatizan que el estado real puede ser latente y que falsos positivos y negativos importan para la decisión [5], [6]. Este trabajo coincide en el uso de estados no observables y evaluación de inferencia, pero difiere al fijar parámetros y datos sintéticos para aislar sensibilidad. No compara contra un umbral operacional o un predictor neuronal, porque el objetivo es estudiar la mecánica probabilística del HMM, no declarar superioridad frente a métodos alternativos.",
"La validez interna es adecuada para la pregunta de simulación: las transiciones y emisiones generan los datos y el decodificador usa exactamente las matrices conocidas. La semilla y los parámetros permiten reproducir los resultados. La validez externa es limitada: las matrices se eligieron para ilustrar hipótesis, las mediciones continuas se discretizaron por umbrales hipotéticos, no se modeló autocorrelación dentro de cada emisión ni deriva temporal, y no existe un conjunto de telemetría real o etiquetas de fallas. Además, los intervalos de confianza reflejan variabilidad entre secuencias generadas, no incertidumbre de una población operacional.",
"Una implementación práctica debe definir ventanas y umbrales desde acuerdos de nivel de servicio, recoger latencia y pérdida sincronizadas, revisar valores ausentes, ajustar A y B en un periodo de entrenamiento y evaluar en otro periodo separado. Debe medir sensibilidad, precisión, falsas alarmas por hora, tiempo de anticipación y calibración probabilística. Si se aprende con Baum–Welch, es necesario usar varias inicializaciones, controlar permutaciones de etiquetas y validar que el modelo no converja a soluciones locales, un problema documentado en HMM para detección de secuencias [3]. El diseño de emisiones continuas o mixtas podría evitar pérdida de información por discretización."
]: add_body(doc, t)

add_heading(doc, "7. Conclusiones y trabajo futuro", 1)
add_body(doc, "P1. El problema se representa mediante un HMM con estados ocultos estable, degradado y caída, observaciones conjuntas de niveles de latencia y pérdida, matriz de transición A, matriz de emisión B y distribución inicial π. La propiedad de Markov modela la evolución de los regímenes; las distribuciones categóricas modelan las transiciones y emisiones.")
add_body(doc, "P2. Los parámetros sí cambiaron los resultados. Al aumentar la permanencia se observó un aumento del F1 macro de 0,706 a 0,772. Al hacer más similares las emisiones, el F1 macro cayó a 0,293 y la sensibilidad a caída a 0,004. Las métricas por clase y los intervalos de confianza evidencian que la separabilidad de B es más determinante que la exactitud aislada.")
add_body(doc, "P3. El experimento es reproducible mediante semillas fijas, especificación publicada y código ejecutable. Cambiar π produjo cambios pequeños para T=300, mientras cambiar A y sobre todo B alteró sustancialmente el rendimiento. Esta conclusión está limitada a la longitud de secuencia y parámetros seleccionados; debe reexaminarse con ventanas cortas y telemetría real.")
add_body(doc, "El trabajo aporta un punto de partida ejecutable para explicar diseño, inferencia y sensibilidad de un HMM aplicado a calidad de conectividad. No demuestra que este modelo esté listo para producción ni presenta un resultado empírico de operador. El trabajo futuro debe validar umbrales y parámetros con datos de red reales, comparar contra reglas de umbral y otros modelos secuenciales, incorporar emisiones continuas, estimar incertidumbre de parámetros y evaluar el costo operativo de falsas alarmas y fallas no detectadas.")

add_heading(doc, "Referencias", 1)
refs = [
"[1] C. A. Gouvea da Silva and C. M. Pedroso, “Packet Loss Characterization Using Cross Layer Information and HMM for Wi-Fi Networks,” Sensors, vol. 22, no. 22, Art. no. 8592, 2022, doi: 10.3390/s22228592.",
"[2] Y. Lu, Y. Chen, J. Liu, Y. Fu, Q. Guo, et al., “A network business availability modeling method based on Markov chain,” EURASIP Journal on Wireless Communications and Networking, vol. 2022, Art. no. 72, 2022, doi: 10.1186/s13638-022-02154-y.",
"[3] X. Zhang, T. Wu, Q. Zheng, L. Zhai, H. Hu, W. Yin, Y. Zeng, and C. Cheng, “Multi-Step Attack Detection Based on Pre-Trained Hidden Markov Models,” Sensors, vol. 22, no. 8, Art. no. 2874, 2022, doi: 10.3390/s22082874.",
"[4] S. S. Hassan, S. Särkkä, and Á. García-Fernández, “Temporal Parallelization of Inference in Hidden Markov Models,” IEEE Transactions on Signal Processing, vol. 69, pp. 4875–4887, 2021, doi: 10.1109/TSP.2021.3103338.",
"[5] M. L. Gámiz, N. Limnios, and M. C. Segovia-García, “Hidden Markov models in reliability and maintenance,” European Journal of Operational Research, vol. 304, no. 3, pp. 1242–1255, 2023, doi: 10.1016/j.ejor.2022.05.006.",
"[6] M. L. Gámiz, F. Navas-Gómez, R. Raya-Miranda, and M. C. Segovia-García, “Dynamic reliability and sensitivity analysis based on HMM models with Markovian signal process,” Reliability Engineering & System Safety, vol. 239, Art. no. 109498, 2023, doi: 10.1016/j.ress.2023.109498.",
"[7] A. Sacco, F. Esposito, and G. Marchetto, “Completing and Predicting Internet Traffic Matrices Using Adversarial Autoencoders and Hidden Markov Models,” IEEE Transactions on Network and Service Management, vol. 20, no. 3, pp. 2244–2258, 2023, doi: 10.1109/TNSM.2023.3270166.",
"[8] Y. Feng, W. Xu, Z. Zhang, et al., “Continuous Hidden Markov Model Based Spectrum Sensing with Estimated SNR for Cognitive UAV Networks,” Sensors, vol. 22, no. 7, Art. no. 2620, 2022, doi: 10.3390/s22072620.",
"[9] A. S. Omer, A. D. Tufa, T. T. Debella, and D. H. Woldegebreal, “Hidden Markov models for predicting cell-level mobile networks performance degradation,” e-Prime - Advances in Electrical Engineering, Electronics and Energy, vol. 5, Art. no. 100742, 2024, doi: 10.1016/j.prime.2024.100742.",
"[10] A. Jalili, J. A. A. Alzubi, R. Rezaei, J. L. Webber, F. R. Alzubi, and C. Fernández-Campusano, “Markov chain-based analysis and fault tolerance technique for enhancing chain-based routing in WSNs,” Concurrency and Computation: Practice and Experience, vol. 36, no. 12, Art. no. e8032, 2024, doi: 10.1002/cpe.8032.",
"[11] M. Heidari et al., “Assessment of reliability and availability of wireless sensor networks in industrial applications by considering permanent faults,” Concurrency and Computation: Practice and Experience, vol. 36, no. 27, Art. no. e8252, 2024, doi: 10.1002/cpe.8252.",
"[12] A. Mahboubi, K. Luong, G. Jarrad, S. Camtepe, M. Bewong, M. Bahutair, and G. Pogrebna, “Lurking in the shadows: Unsupervised decoding of beaconing communication for enhanced cyber threat hunting,” Journal of Network and Computer Applications, vol. 236, Art. no. 104127, 2025, doi: 10.1016/j.jnca.2025.104127.",
"[13] L. E. Baum, T. Petrie, G. Soules, and N. Weiss, “A maximization technique occurring in the statistical analysis of probabilistic functions of Markov chains,” Annals of Mathematical Statistics, vol. 41, no. 1, pp. 164–171, 1970, doi: 10.1214/aoms/1177697196.",
"[14] L. R. Rabiner, “A tutorial on hidden Markov models and selected applications in speech recognition,” Proceedings of the IEEE, vol. 77, no. 2, pp. 257–286, 1989, doi: 10.1109/5.18626.",
"[15] E. N. Gilbert, “Capacity of a burst-noise channel,” Bell System Technical Journal, vol. 39, no. 5, pp. 1253–1265, 1960, doi: 10.1002/j.1538-7305.1960.tb03959.x."
]
for ref in refs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(4)
    style_run(p.add_run(ref), size=8.5)

add_heading(doc, "Anexo A. Archivos y reproducción", 1)
add_body(doc, "El proyecto adjunto contiene el código en src/simulate_hmm.py; los parámetros y versiones del entorno en data/model_parameters.json; las 1.600 filas de métricas por réplica en data/replicate_metrics.csv; las tablas estadísticas resumidas en data/summary_metrics.csv; y las figuras PNG a 300 dpi en figures/. Para repetir el análisis, instalar las dependencias indicadas en requirements.txt y ejecutar python src/simulate_hmm.py desde la carpeta raíz del proyecto. La reproducción requiere Python 3.12 o una versión compatible con las dependencias señaladas.")
add_body(doc, "Nota de alcance bibliográfico: las publicaciones se presentan con sus DOI y datos de revista para facilitar su verificación. La pertenencia a Q1 depende del año, categoría y base de evaluación; debe comprobarse en Scopus Sources o Journal Citation Reports/Web of Science con los accesos de la institución antes de afirmar el cuartil de cada revista.")

doc.save(OUT)
print(OUT)
