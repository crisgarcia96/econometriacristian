# -*- coding: utf-8 -*-  # Permite escribir comentarios, textos y títulos en español.
"""Análisis Box-Jenkins del café: nueve secciones, resultados solamente en pantalla."""  # Presenta el objetivo general.

# SECCIÓN 0. DEPENDENCIAS, PARÁMETROS Y FUNCIONES AUXILIARES.  # Identifica la configuración inicial.
import argparse  # Permite configurar el programa desde la terminal sin modificar el código.
import inspect  # Comprueba qué argumentos admite la versión instalada de statsmodels.
import sys  # Permite comprobar si la entrada por terminal está disponible.
import warnings  # Permite capturar advertencias de convergencia durante la estimación.
from pathlib import Path  # Construye rutas relativas compatibles con Windows, macOS y Linux.
import numpy as np  # Realiza cálculos numéricos, logaritmos y transformaciones exponenciales.
import pandas as pd  # Lee el Excel y organiza los datos como serie temporal mensual.
import matplotlib  # Controla el motor de generación de gráficos.
import matplotlib.pyplot as plt  # Dibuja y guarda los gráficos del ejercicio.
from scipy import stats  # Calcula normalidad, cuantiles y bandas de confianza.
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf  # Dibuja FAC y FACP con intervalos.
from statsmodels.stats.diagnostic import acorr_ljungbox, het_arch  # Ejecuta Ljung-Box y ARCH-LM.
from statsmodels.tsa.seasonal import STL  # Descompone la serie en tendencia, estacionalidad y residuo.
from statsmodels.tsa.statespace.sarimax import SARIMAX  # Estima ARIMA y SARIMA sin selección automática.
from statsmodels.tsa.stattools import adfuller, kpss  # Evalúa estacionariedad con hipótesis opuestas.
from statsmodels.tools.sm_exceptions import ConvergenceWarning  # Identifica alertas de convergencia.

RAIZ = Path(__file__).resolve().parent  # Toma como referencia la ubicación real de este archivo.
CARPETA_DATOS = RAIZ / "datos"  # Define la carpeta local para las bases entregadas.
ARCHIVO_EJEMPLO = CARPETA_DATOS / "Exportaciones de café.xlsx"  # Indica la base incluida en el ZIP.
HORIZONTE = 10  # Establece diez meses conforme a las instrucciones académicas.
ALFA = 0.05  # Fija el nivel de significancia de referencia en cinco por ciento.
MES_NOMBRE = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}  # Abrevia meses para gráficas.
plt.rcParams.update({"figure.dpi": 120, "axes.grid": True, "grid.alpha": 0.18, "font.size": 10})  # Mejora la visualización de las figuras en pantalla.


def opciones_programa():  # Define las opciones de ejecución interactiva y automática.
    parser = argparse.ArgumentParser(description="Box-Jenkins del volumen exportado de café sin tostar por Colombia.")  # Describe el script.
    parser.add_argument("--excel", type=str, help="Ruta de un Excel; evita abrir el selector de archivos.")  # Facilita pruebas reproducibles.
    parser.add_argument("--hoja", type=str, default="Detalle", help="Nombre de la hoja de datos (por defecto: Detalle).")  # Permite cambiar la hoja.
    parser.add_argument("--hasta", type=int, choices=range(1, 10), default=9, metavar="1-9", help="Última etapa visible que debe ejecutarse.")  # Permite estudiar por bloques.
    parser.add_argument("--paso-a-paso", action="store_true", help="Hace una pausa después de cada sección.")  # Permite revisar los resultados.
    parser.add_argument("--sin-graficas", action="store_true", help="Omite las ventanas gráficas; conserva las tablas en consola.")  # Permite comprobaciones sin interfaz gráfica.
    parser.add_argument("--sin-dialogo", action="store_true", help="Utiliza la base incluida en datos/, si existe.")  # Permite automatizar pruebas con Excel adjunto.
    return parser.parse_args()  # Devuelve las opciones especificadas por el usuario.


def elegir_excel(args):  # Pregunta por el archivo de Excel sin depender de una ruta fija.
    if args.excel:  # Comprueba si el usuario indicó explícitamente un Excel.
        ruta = Path(args.excel).expanduser().resolve()  # Convierte el argumento en una ruta absoluta verificable.
    elif args.sin_dialogo:  # Comprueba si debe utilizarse directamente el ejemplo incluido.
        ruta = ARCHIVO_EJEMPLO  # Conserva la ruta relativa al proyecto empaquetado.
    else:  # Abre el selector gráfico en una ejecución normal.
        ruta = None  # Inicializa el resultado por si el selector no está disponible.
        try:  # Intenta utilizar la ventana nativa del sistema operativo.
            from tkinter import Tk, filedialog  # Importa el selector gráfico disponible en Python estándar.
            ventana = Tk()  # Crea una ventana auxiliar.
            ventana.withdraw()  # Oculta la ventana auxiliar y conserva solo el diálogo.
            ventana.attributes("-topmost", True)  # Trae el selector al frente en sistemas que lo permiten.
            elegido = filedialog.askopenfilename(title="Selecciona el Excel de exportaciones de café", initialdir=str(CARPETA_DATOS if CARPETA_DATOS.is_dir() else RAIZ), filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")])  # Solicita el archivo al usuario.
            ventana.destroy()  # Cierra correctamente la ventana auxiliar.
            if elegido:  # Verifica que se seleccionó un documento.
                ruta = Path(elegido).expanduser().resolve()  # Convierte la selección en una ruta utilizable.
        except Exception as exc:  # Captura fallos del entorno gráfico sin ocultar el motivo.
            print(f"Selector gráfico no disponible: {exc}")  # Explica por qué se usará la terminal.
        if ruta is None and sys.stdin.isatty():  # Ofrece alternativa únicamente si hay una terminal interactiva.
            entrada = input("Escribe la ruta completa del Excel: ").strip().strip('"')  # Solicita un Excel situado en cualquier carpeta.
            if not entrada and ARCHIVO_EJEMPLO.is_file():  # Ofrece el ejemplo solo cuando realmente está incluido.
                ruta = ARCHIVO_EJEMPLO  # Reutiliza el Excel de ejemplo disponible.
            elif entrada:  # Comprueba que se haya escrito una ruta.
                ruta = Path(entrada).expanduser().resolve()  # Usa el archivo elegido sin depender del proyecto.
            else:  # Evita intentar abrir un ejemplo inexistente.
                raise ValueError("Debes seleccionar o indicar la ruta de un Excel válido.")  # Facilita el uso con solo el archivo Python.
        if ruta is None:  # Comprueba si no hubo diálogo ni terminal interactiva.
            raise RuntimeError("No fue posible seleccionar un Excel. Ejecuta con --excel 'ruta/datos.xlsx'.")  # Evita cargar otro archivo sin autorización.
    if not ruta.is_file():  # Comprueba que el archivo realmente exista.
        raise FileNotFoundError(f"No existe el Excel seleccionado: {ruta}")  # Da un mensaje comprensible.
    if ruta.suffix.lower() not in {".xlsx", ".xls"}:  # Comprueba las extensiones admitidas.
        raise ValueError("Selecciona un archivo .xlsx o .xls.")  # Evita errores de formato difíciles de diagnosticar.
    print(f"\nExcel seleccionado: {ruta}")  # Deja constancia del documento utilizado.
    return ruta  # Devuelve la ruta validada a la etapa de lectura.



def mostrar_figura(nombre, args):  # Prepara la figura sin escribir archivos en disco.
    plt.tight_layout()  # Evita superposición de títulos y etiquetas.
    print(f"  Gráfica preparada para visualizar: {nombre.replace('.png', '')}")  # Identifica la imagen que se abrirá en PyCharm.
    if args.sin_graficas:  # Permite omitir ventanas en ejecuciones de comprobación.
        plt.close(plt.gcf())  # Descarta la figura sin exportarla cuando el usuario lo solicite.


def mostrar_tabla(df, nombre):  # Presenta la tabla completa directamente en la consola.
    print(f"\n  TABLA: {nombre.replace('.csv', '').replace('_', ' ')}")  # Etiqueta la tabla sin sugerir que existe un CSV.
    print(df.to_string(index=False, max_rows=None, max_cols=None, float_format=lambda v: f"{v:,.5f}"))  # Imprime todas las filas y columnas, sin guardar archivos.


def mostrar_etapa(numero, nombre):  # Imprime un encabezado identificable para cada bloque.
    print(f"\n{'=' * 75}\nSECCIÓN {numero}. {nombre}\n{'=' * 75}")  # Separa visualmente las etapas.


def preguntar_pausa(args, numero):  # Permite avanzar manualmente entre las secciones.
    if args.paso_a_paso and numero < args.hasta:  # Solo interrumpe cuando se solicitó el modo didáctico.
        if not args.sin_graficas and matplotlib.get_backend().lower() != "agg":  # Permite revisar gráficas al terminar cada bloque.
            plt.show()  # Abre las figuras del bloque hasta que se cierren las ventanas.
            plt.close("all")  # Libera las figuras ya revisadas antes del siguiente bloque.
        input("\nRevisa los resultados de esta sección y pulsa Enter para continuar...")  # Detiene la ejecución hasta que el estudiante decida avanzar.


# SECCIÓN 1. CARGA, COMPROBACIÓN E INTEGRIDAD TEMPORAL.  # Identifica las tareas de preparación de datos.
def seccion_1(ruta, args):  # Lee las observaciones mensuales y comprueba que no haya errores.
    mostrar_etapa(1, "Carga y revisión de datos")  # Presenta la primera etapa.
    base = pd.read_excel(ruta, sheet_name=args.hoja)  # Lee directamente la hoja suministrada por el usuario.
    base.columns = base.columns.astype(str).str.strip()  # Elimina espacios innecesarios de los encabezados.
    fecha_col = "Mes"  # Nombra la columna de fechas esperada según el archivo suministrado.
    valor_col = "Peso en Toneladas Netas"  # Nombra la variable económica que se modelará.
    faltan = [x for x in (fecha_col, valor_col) if x not in base.columns]  # Identifica columnas ausentes.
    if faltan:  # Interrumpe si los encabezados no corresponden al formato de la base.
        raise ValueError(f"Faltan columnas {faltan}. Columnas disponibles: {base.columns.tolist()}")  # Explica cómo corregir la base.
    base[fecha_col] = pd.to_datetime(base[fecha_col], errors="coerce")  # Convierte fechas de Excel a fechas reales.
    base[valor_col] = pd.to_numeric(base[valor_col], errors="coerce")  # Convierte toneladas en números decimales.
    if base[[fecha_col, valor_col]].isna().any().any():  # Detecta fechas vacías o valores no numéricos.
        raise ValueError("Hay fechas o toneladas vacías/no numéricas: revisar la base antes de estimar.")  # Impide interpolaciones silenciosas.
    if (base[valor_col] <= 0).any():  # Comprueba la condición necesaria para trabajar con logaritmos.
        raise ValueError("Se encontraron toneladas menores o iguales a cero; el logaritmo no está definido.")  # Evita transformar valores inválidos.
    if base[fecha_col].dt.day.ne(1).any():  # Revisa si las fechas están identificadas por el primer día del mes.
        print("  Nota: normalizando las fechas al primer día de cada mes.")  # Documenta el ajuste de calendario.
    base[fecha_col] = base[fecha_col].dt.to_period("M").dt.to_timestamp()  # Normaliza las fechas a frecuencia mensual.
    base = base.sort_values(fecha_col).reset_index(drop=True)  # Ordena cronológicamente los registros.
    if base[fecha_col].duplicated().any():  # Comprueba si algún mes aparece más de una vez.
        raise ValueError("Existen meses duplicados; la base debe presentar un total mensual único.")  # No agrega automáticamente registros.
    indice = pd.date_range(base[fecha_col].min(), base[fecha_col].max(), freq="MS")  # Genera el calendario mensual esperado.
    if len(indice) != len(base) or not indice.equals(pd.DatetimeIndex(base[fecha_col])):  # Comprueba continuidad sin huecos.
        raise ValueError("La base tiene meses faltantes o fechas discontinuas; corregir antes de continuar.")  # Impide series temporales artificiales.
    if len(base) < 100:  # Comprueba el requisito mínimo de observaciones del curso.
        raise ValueError("La serie debe tener por lo menos 100 observaciones mensuales.")  # Alinea el ejercicio con la guía.
    serie = pd.Series(base[valor_col].to_numpy(dtype=float), index=indice, name="Toneladas netas")  # Construye la serie original.
    serie.index.name = "Mes"  # Asigna el nombre de la dimensión temporal.
    mostrar_tabla(serie.rename("Toneladas_netas").reset_index(), "01_base_validada.csv")  # Imprime los registros procesados sin modificar el Excel.
    print(f"  Periodo: {serie.index[0]:%Y-%m} a {serie.index[-1]:%Y-%m}")  # Reporta la cobertura temporal.
    print(f"  Observaciones: {len(serie)}; sin faltantes ni duplicados.")  # Documenta la calidad de datos.
    print(f"  Toneladas: promedio = {serie.mean():,.2f}; último mes = {serie.iloc[-1]:,.2f}")  # Resume magnitudes económicas.
    print("  Fuente: agregado aportado por el usuario de las subpartidas 0901111000, 0901119000 y 0901120000.")  # Documenta el universo arancelario.
    return serie  # Entrega datos confiables a las siguientes secciones.


# SECCIÓN 2. EXPLORACIÓN Y ESTACIONALIDAD.  # Distingue estacionalidad de estacionariedad.
def seccion_2(serie, args):  # Describe los datos e inspecciona posibles patrones anuales recurrentes.
    mostrar_etapa(2, "Exploración y estacionalidad (no estacionariedad)")  # Identifica el propósito exploratorio.
    fig, ax = plt.subplots(figsize=(11.5, 4.6))  # Abre un gráfico panorámico adecuado para póster.
    ax.plot(serie.index, serie, color="#244b73", linewidth=1.6)  # Dibuja el volumen exportado mes a mes.
    ax.set(title="Exportaciones de café sin tostar de Colombia", xlabel="Año", ylabel="Toneladas netas")  # Identifica variable y unidades.
    mostrar_figura("02_serie_original.png", args)  # Presenta la figura histórica.
    grupos = [serie[serie.index.month == m].to_numpy() for m in range(1, 13)]  # Agrupa observaciones por mes calendario.
    fig, ax = plt.subplots(figsize=(11, 4.8))  # Crea un gráfico para investigar estacionalidad descriptiva.
    ax.boxplot(grupos, tick_labels=[MES_NOMBRE[m] for m in range(1, 13)], showfliers=False)  # Compara las distribuciones mensuales.
    ax.set(title="Distribución de toneladas exportadas por mes calendario", xlabel="Mes", ylabel="Toneladas netas")  # Identifica los ejes.
    mostrar_figura("03_estacionalidad_mensual.png", args)  # Presenta evidencia visual de estacionalidad.
    stl = STL(serie, period=12, robust=True).fit()  # Descompone la serie; no modifica ni desestacionaliza los datos usados en los modelos.
    fig = stl.plot()  # Muestra nivel, tendencia, componente estacional y residuo.
    fig.set_size_inches(11, 8)  # Aumenta el tamaño del diagnóstico temporal.
    mostrar_figura("04_descomposicion_STL.png", args)  # Presenta la descomposición exploratoria.
    fuerza = max(0.0, 1.0 - np.var(stl.resid) / np.var(stl.resid + stl.seasonal))  # Calcula fuerza estacional descriptiva de STL.
    print(f"  Fuerza estacional descriptiva (STL, 0 a 1): {fuerza:.3f}")  # Presenta el indicador sin confundirlo con una prueba formal.
    print("  Importante: estacionalidad no equivale a no estacionariedad.")  # Explicita la distinción conceptual.
    return {"fuerza_stl": fuerza}  # Facilita reportar el diagnóstico más adelante.


# SECCIÓN 3. TRANSFORMACIONES Y PRUEBAS DE ESTACIONARIEDAD.  # Prepara la serie modelada.
def pruebas_estacionariedad(serie, etiqueta):  # Aplica ADF y KPSS e informa sus hipótesis nulas.
    if "result_object" in inspect.signature(adfuller).parameters:  # Identifica las versiones recientes de statsmodels.
        adf = adfuller(serie, autolag="AIC", result_object=False)  # Mantiene el resultado tradicional y evita la advertencia futura.
    else:  # Conserva compatibilidad con versiones anteriores de statsmodels.
        adf = adfuller(serie, autolag="AIC")  # Contrasta H0: la serie contiene raíz unitaria.
    with warnings.catch_warnings():  # Captura avisos por límites tabulados de valores p en KPSS.
        warnings.simplefilter("ignore", category=Warning)  # Evita que las advertencias tapen las cifras de la prueba.
        kp = kpss(serie, regression="c", nlags="auto")  # Contrasta H0: la serie es estacionaria alrededor de una constante.
    return {"Serie": etiqueta, "ADF_p": float(adf[1]), "KPSS_p": float(kp[1]), "ADF_rechaza_raiz": bool(adf[1] < ALFA), "KPSS_no_rechaza_estac": bool(kp[1] >= ALFA)}  # Resume pruebas con hipótesis diferentes.


def seccion_3(serie, args):  # Evalúa transformaciones sin imponer conclusiones antes de observar los resultados.
    mostrar_etapa(3, "Logaritmo, primera diferencia y pruebas ADF/KPSS")  # Presenta el bloque de transformación.
    log_serie = np.log(serie)  # Transforma el volumen en logaritmo natural de toneladas.
    dlog = log_serie.diff().dropna()  # Calcula el crecimiento logarítmico mensual y elimina el primer NA.
    comparacion = pd.DataFrame([pruebas_estacionariedad(serie, "Toneladas originales"), pruebas_estacionariedad(log_serie, "Logaritmo de toneladas"), pruebas_estacionariedad(dlog, "Primera diferencia logarítmica")])  # Evalúa las tres representaciones.
    mostrar_tabla(comparacion, "05_pruebas_estacionariedad.csv")  # Presenta resultados sin redondear excesivamente.
    if not (comparacion.iloc[2]["ADF_rechaza_raiz"] and comparacion.iloc[2]["KPSS_no_rechaza_estac"]):  # Evalúa si la primera diferencia logró evidencia compatible con estacionariedad.
        print("  ADVERTENCIA: ADF y KPSS no respaldan conjuntamente d=1; revisar antes de validar el modelo.")  # Evita afirmar estacionariedad sin prueba.
    else:  # Informa una lectura favorable de las dos pruebas.
        print("  Evidencia compatible con estacionariedad de la primera diferencia logarítmica.")  # Resume los contrastes sin declararlos demostración absoluta.
    fig, axes = plt.subplots(2, 1, figsize=(11, 7), sharex=True)  # Compara serie logarítmica y diferencia logarítmica.
    axes[0].plot(log_serie.index, log_serie, color="#244b73", linewidth=1.2)  # Muestra el logaritmo de las toneladas.
    axes[0].set(ylabel="ln(toneladas)", title="Transformación logarítmica")  # Identifica la escala transformada.
    axes[1].plot(dlog.index, dlog, color="#447b8c", linewidth=1.0)  # Muestra los cambios logarítmicos mensuales.
    axes[1].set(xlabel="Año", ylabel="Δ ln(toneladas)", title="Primera diferencia del logaritmo")  # Identifica las variaciones.
    mostrar_figura("06_logaritmos_diferenciacion.png", args)  # Presenta las dos transformaciones.
    return {"log": log_serie, "dlog": dlog, "pruebas": comparacion}  # Devuelve series y evidencias estadísticas.


# SECCIÓN 4. IDENTIFICACIÓN VISUAL (FAC/FACP).  # Examina órdenes y persistencia estacional.
def seccion_4(transformacion, args):  # Construye correlogramas de la serie logarítmica y diferenciada.
    mostrar_etapa(4, "FAC/FACP e identificación de posibles componentes estacionales")  # Presenta el diagnóstico.
    for sufijo, nombre, valores in [("log", "Logaritmo en niveles", transformacion["log"]), ("dlog", "Diferencia logarítmica", transformacion["dlog"])]:  # Ejecuta el mismo diagnóstico para ambas series.
        fig, axes = plt.subplots(1, 2, figsize=(11.5, 4.6))  # Prepara FAC y FACP lado a lado.
        plot_acf(valores, ax=axes[0], lags=36, alpha=0.05, zero=False)  # Calcula FAC con bandas de confianza al 95 %.
        plot_pacf(valores, ax=axes[1], lags=36, alpha=0.05, zero=False, method="ywm")  # Calcula FACP con bandas de confianza.
        axes[0].set(title=f"FAC: {nombre}", ylim=(-1, 1), xlabel="Rezago mensual")  # Establece escala de correlación correcta.
        axes[1].set(title=f"FACP: {nombre}", ylim=(-1, 1), xlabel="Rezago mensual")  # Permite reconocer posibles órdenes p y q.
        mostrar_figura(f"07_fac_facp_{sufijo}.png", args)  # Presenta la evidencia gráfica.
    autocorr_12 = float(transformacion["dlog"].autocorr(lag=12))  # Calcula la asociación en el rezago anual de la serie estacionarizada.
    umbral = float(1.96 / np.sqrt(len(transformacion["dlog"])))  # Aproxima el umbral visual de correlación al 95 %.
    print(f"  Autocorrelación de Δlog en rezago 12: {autocorr_12:.4f}; banda aproximada: ±{umbral:.4f}")  # Da evidencia para evaluar estacionalidad.
    print("  FAC/FACP orientan p, q, P y Q; no son pruebas formales de raíz unitaria.")  # Evita confundir identificación con prueba de estacionariedad.
    return {"autocorr_12": autocorr_12, "banda": umbral}  # Entrega un indicador descriptivo para el informe.


# SECCIÓN 5. ESTIMACIÓN MANUAL DE MODELOS CANDIDATOS.  # Excluye cualquier función de selección automática.
def candidatos():  # Define explícitamente los órdenes que justificaremos mediante FAC/FACP y parsimonia.
    lista = [  # Enumera los candidatos manuales; no llama a funciones de selección automática.
        ("ARIMA(0,1,1)", (0, 1, 1), (0, 0, 0, 12)),  # Referencia sencilla: una media móvil.
        ("ARIMA(1,1,0)", (1, 1, 0), (0, 0, 0, 12)),  # Referencia sencilla: un autorregresivo.
        ("ARIMA(1,1,1)", (1, 1, 1), (0, 0, 0, 12)),  # Modelo AR y MA de primer orden.
        ("ARIMA(2,1,0)", (2, 1, 0), (0, 0, 0, 12)),  # Evalúa persistencia autorregresiva de orden dos.
        ("ARIMA(0,1,2)", (0, 1, 2), (0, 0, 0, 12)),  # Evalúa medias móviles de orden dos.
        ("ARIMA(2,1,1)", (2, 1, 1), (0, 0, 0, 12)),  # Combina dos términos AR con uno MA.
        ("ARIMA(1,1,2)", (1, 1, 2), (0, 0, 0, 12)),  # Combina un término AR con dos MA.
        ("ARIMA(2,1,2)", (2, 1, 2), (0, 0, 0, 12)),  # Incluye la especificación utilizada en el trabajo anterior.
        ("SARIMA(0,1,1)(1,0,0)[12]", (0, 1, 1), (1, 0, 0, 12)),  # Introduce AR estacional anual.
        ("SARIMA(1,1,1)(1,0,0)[12]", (1, 1, 1), (1, 0, 0, 12)),  # Combina ARMA de orden uno y AR estacional.
        ("SARIMA(1,1,1)(0,0,1)[12]", (1, 1, 1), (0, 0, 1, 12)),  # Alternativa con media móvil estacional.
    ]  # Concluye la lista de modelos estimados sobre la misma serie logarítmica.
    return lista  # Entrega una lista fija, legible y modificable manualmente por el estudiante.


def estimar(serie_log, orden, estacional):  # Estima una especificación fija, sin buscar modelos automáticamente.
    modelo = SARIMAX(serie_log, order=orden, seasonal_order=estacional, trend="n", enforce_stationarity=True, enforce_invertibility=True)  # Estima sobre logaritmos en niveles, con d=1 interno y sin deriva.
    with warnings.catch_warnings():  # Evita bloquear la exploración por avisos de optimización.
        warnings.simplefilter("ignore", ConvergenceWarning)  # La convergencia se inspecciona por separado en el resultado.
        ajuste = modelo.fit(disp=False, maxiter=100)  # Estima parámetros por máxima verosimilitud en el espacio de estados.
    return ajuste  # Conserva parámetros, información de convergencia y estructura residual.


def residuales_utiles(ajuste, orden, estacional):  # Elimina residuos iniciales afectados por la inicialización difusa.
    quemar = max(int(getattr(ajuste, "loglikelihood_burn", 0)), orden[1] + estacional[1] * 12, 1)  # Determina observaciones de inicialización.
    return pd.Series(ajuste.resid, dtype=float).iloc[quemar:].dropna()  # Devuelve residuos de una predicción paso a paso.


def evaluar_candidato(nombre, orden, estacional, ajuste):  # Prepara una fila comparable de resultados y diagnóstico.
    convergio = bool(ajuste.mle_retvals.get("converged", True))  # Comprueba el resultado del optimizador.
    gdl = orden[0] + orden[2] + estacional[0] + estacional[2]  # Cuenta términos AR/MA para corrección de Ljung-Box.
    resid = residuales_utiles(ajuste, orden, estacional)  # Selecciona residuos útiles para diagnóstico.
    p_lb = float(acorr_ljungbox(resid, lags=[24], model_df=gdl, return_df=True)["lb_pvalue"].iloc[0])  # Contrasta autocorrelación residual hasta 24 meses.
    return {  # Entrega una fila homogénea para comparar los modelos candidatos.
        "Modelo": nombre,  # Nombre de la especificación estimada.
        "p": orden[0], "d": orden[1], "q": orden[2],  # Órdenes no estacionales.
        "P": estacional[0], "D": estacional[1], "Q": estacional[2],  # Órdenes estacionales.
        "s": 12,  # Periodicidad anual de la serie mensual.
        "AIC": float(ajuste.aic),  # Criterio de información de Akaike.
        "BIC": float(ajuste.bic),  # Criterio de información bayesiano.
        "Ljung_Box_p24": p_lb,  # Valor p de autocorrelación residual hasta 24 rezagos.
        "Convergió": convergio,  # Indica si la optimización fue satisfactoria.
    }  # Termina el registro de comparación.


def seccion_5(log_serie, args):  # Selecciona entre especificaciones explícitas usando solo entrenamiento.
    mostrar_etapa(5, "Estimación de candidatos y selección sin auto.arima")  # Señala que la selección es manual.
    entrenamiento = log_serie.iloc[:-HORIZONTE]  # Reserva las últimas diez observaciones antes de comparar modelos.
    prueba = log_serie.iloc[-HORIZONTE:]  # Separa los diez meses que servirán para validar pronósticos.
    filas = []  # Inicializa la tabla de comparación de modelos.
    ajustes = {}  # Almacena temporalmente los ajustes que convergieron.
    for nombre, orden, estacional in candidatos():  # Recorre únicamente los modelos definidos explícitamente.
        try:  # Impide que un modelo numéricamente inestable detenga los demás.
            ajuste = estimar(entrenamiento, orden, estacional)  # Ajusta cada modelo sin ver la prueba final.
            fila = evaluar_candidato(nombre, orden, estacional, ajuste)  # Calcula información y residuos.
            filas.append(fila)  # Añade el modelo evaluado a la tabla.
            if fila["Convergió"] and np.isfinite(fila["BIC"]):  # Solo admite modelos numéricamente viables.
                ajustes[nombre] = (orden, estacional, ajuste)  # Presenta el candidato para la selección.
            print(f"  {nombre:<32} BIC={fila['BIC']:>10.2f}  Ljung-Box p={fila['Ljung_Box_p24']:.4f}  convergencia={fila['Convergió']}")  # Expone resultados útiles para discusión.
        except Exception as exc:  # Registra fallos sin sustituirlos por números inventados.
            print(f"  {nombre}: no estimable ({type(exc).__name__}: {exc})")  # Documenta problemas de la estimación.
    if not ajustes:  # Comprueba si al menos una especificación convergió.
        raise RuntimeError("No convergió ningún candidato. Revisar datos y órdenes definidos manualmente.")  # Detiene un procedimiento sin resultados válidos.
    tabla = pd.DataFrame(filas).sort_values("BIC").reset_index(drop=True)  # Ordena los modelos por parsimonia penalizada.
    mostrar_tabla(tabla, "08_comparacion_modelos_entrenamiento.csv")  # Presenta la tabla principal para el póster.
    aceptables = [fila for fila in filas if fila["Modelo"] in ajustes and fila["Ljung_Box_p24"] > ALFA]  # Prioriza ausencia de autocorrelación detectable.
    elegibles = aceptables if aceptables else [fila for fila in filas if fila["Modelo"] in ajustes]  # Admite alternativa solo si ningún modelo supera la prueba.
    elegido = min(elegibles, key=lambda fila: fila["BIC"])  # Selecciona el menor BIC dentro del conjunto con buen diagnóstico.
    nombre = elegido["Modelo"]  # Identifica el modelo final elegido sin ocultar el criterio aplicado.
    orden, estacional, ajuste_train = ajustes[nombre]  # Recupera el ajuste de entrenamiento sin filtración temporal.
    if not aceptables:  # Reconoce si ningún candidato pasó el contraste residual.
        print("  ADVERTENCIA: ningún candidato pasó Ljung-Box p>0.05 en entrenamiento; la selección requiere revisión.")  # Evita declarar validado un modelo deficiente.
    print(f"\n  Modelo elegido: {nombre} (BIC entre candidatos convergentes con Ljung-Box p>0,05, si existe).")  # Explica la regla aplicada.
    print("  IMPORTANTE: candidatos definidos manualmente; la tabla no utiliza selección automática.")  # Alinea el procedimiento con la guía.
    return {"entrenamiento": entrenamiento, "prueba": prueba, "tabla": tabla, "nombre": nombre, "orden": orden, "estacional": estacional, "ajuste_train": ajuste_train}  # Devuelve la selección documentada.


# SECCIÓN 6. VALIDACIÓN PREDICTIVA A DIEZ MESES.  # Usa observaciones nunca vistas en el ajuste.
def seccion_6(seleccion, args):  # Evalúa la capacidad predictiva en el horizonte oficial.
    mostrar_etapa(6, "Validación fuera de muestra: diez meses")  # Identifica la prueba predictiva.
    reales = np.exp(seleccion["prueba"])  # Devuelve las observaciones reservadas a toneladas originales.
    anticipados = seleccion["ajuste_train"].get_forecast(steps=HORIZONTE)  # Pronostica diez meses desde el final del entrenamiento.
    pred_mediana = np.exp(anticipados.predicted_mean)  # Retransforma pronósticos logarítmicos en medianas condicionales.
    pred_media = np.exp(anticipados.predicted_mean + 0.5 * anticipados.var_pred_mean)  # Aproxima la media bajo supuesto lognormal.
    error = reales.to_numpy() - pred_media.to_numpy()  # Calcula diferencias entre toneladas reales y medias previstas.
    rmse = float(np.sqrt(np.mean(error ** 2)))  # Calcula raíz del error cuadrático medio en toneladas.
    mae = float(np.mean(np.abs(error)))  # Calcula error absoluto medio en toneladas.
    mape = float(100.0 * np.mean(np.abs(error / reales.to_numpy())))  # Calcula error porcentual absoluto medio.
    escala = float(np.mean(np.abs(np.diff(np.exp(seleccion["entrenamiento"])))))  # Calcula referencia ingenua usando solo entrenamiento.
    mase = float(mae / escala) if escala > 0 else float("nan")  # Escala MAE contra la variación ingenua del entrenamiento.
    base_ingenua = float(np.exp(seleccion["entrenamiento"].iloc[-1]))  # Obtiene el último valor conocido al momento de pronosticar.
    pronostico_ingenuo = np.full(HORIZONTE, base_ingenua)  # Usa el mismo último dato para los diez horizontes como referencia.
    rmse_ingenuo = float(np.sqrt(np.mean((reales.to_numpy() - pronostico_ingenuo) ** 2)))  # Calcula el RMSE ingenuo fuera de muestra.
    mae_ingenuo = float(np.mean(np.abs(reales.to_numpy() - pronostico_ingenuo)))  # Calcula el MAE ingenuo fuera de muestra.
    metricas = pd.DataFrame([{  # Construye una tabla con métricas del modelo y su punto de referencia.
        "Modelo": seleccion["nombre"],  # Identifica el modelo que se validó.
        "RMSE_ton": rmse,  # Presenta error cuadrático del modelo en toneladas.
        "MAE_ton": mae,  # Presenta error absoluto del modelo en toneladas.
        "MAPE_pct": mape,  # Presenta error porcentual absoluto medio.
        "MASE": mase,  # Presenta error absoluto escalado según entrenamiento.
        "RMSE_ingenuo_ton": rmse_ingenuo,  # Permite comparación con referencia genuinamente fuera de muestra.
        "MAE_ingenuo_ton": mae_ingenuo,  # Permite comparación sobre el mismo horizonte de diez meses.
        "Meses_prueba": HORIZONTE,  # Documenta tamaño y horizonte de validación.
    }])  # Concluye las métricas de validación.
    mostrar_tabla(metricas, "09_metricas_validacion_10_meses.csv")  # Presenta métricas del conjunto de prueba.
    tabla = pd.DataFrame({"Mes": reales.index, "Real_ton": reales.to_numpy(), "Pronostico_puntual_media_ton": pred_media.to_numpy(), "Pronostico_mediana_ton": pred_mediana.to_numpy(), "Error_ton": error})  # Construye validación mensual.
    mostrar_tabla(tabla, "10_validacion_mensual.csv")  # Permite revisar mes por mes en consola.
    fig, ax = plt.subplots(figsize=(10.5, 4.8))  # Abre gráfico de evaluación fuera de muestra.
    ax.plot(reales.index, reales.to_numpy(), label="Toneladas observadas", marker="o", color="#244b73")  # Dibuja datos que no se usaron al estimar.
    ax.plot(reales.index, pred_media.to_numpy(), label="Pronóstico puntual (media)", marker="o", color="#b85c38")  # Contrasta el modelo.
    ax.set(title="Evaluación fuera de muestra: últimos diez meses", xlabel="Mes", ylabel="Toneladas netas")  # Explica periodo y unidades.
    ax.legend()  # Diferencia valores observados de pronosticados.
    mostrar_figura("11_validacion_10_meses.png", args)  # Presenta figura para discusión de desempeño.
    if mase >= 1.0:  # Reconoce la escala MASE sin confundirla con prueba directa contra el ingenuo multihorizonte.
        print("  NOTA: MASE >= 1 significa MAE superior al error ingenuo medio de entrenamiento (escala MASE).")  # Evita una comparación predictiva incorrecta.
    if mae >= mae_ingenuo:  # Compara errores fuera de muestra en idénticos diez meses.
        print("  ADVERTENCIA: el MAE del modelo no mejora el pronóstico ingenuo de diez meses.")  # Informa una limitación sustentada.
    else:  # Reporta cuando sí supera la referencia del mismo horizonte.
        print("  El MAE del modelo mejora el pronóstico ingenuo en los diez meses reservados.")  # Distingue evaluación real de escala MASE.
    return {"metricas": metricas, "validacion": tabla}  # Devuelve cifras para informe y póster.


# SECCIÓN 7. REESTIMACIÓN FINAL, COEFICIENTES Y ECUACIÓN.  # Incorpora todos los meses después de validar.
def seccion_7(log_serie, seleccion, args):  # Obtiene el modelo para pronosticar meses futuros reales.
    mostrar_etapa(7, "Reestimación de modelo final y ecuación")  # Presenta resultados de estimación final.
    ajuste = estimar(log_serie, seleccion["orden"], seleccion["estacional"])  # Reestima mismos órdenes usando las 212 observaciones.
    parametros = pd.DataFrame({  # Construye la tabla estadística del modelo final.
        "Parametro": ajuste.params.index.astype(str),  # Identifica cada término AR, MA, estacional y varianza.
        "Coeficiente": ajuste.params.to_numpy(),  # Presenta las estimaciones por máxima verosimilitud.
        "Error_estandar": ajuste.bse.to_numpy(),  # Presenta la incertidumbre estimada de los coeficientes.
        "Estadistico_z": ajuste.params.to_numpy() / ajuste.bse.to_numpy(),  # Calcula los estadísticos normales aproximados.
        "Valor_p": ajuste.pvalues.to_numpy(),  # Presenta la significancia individual aproximada.
    })  # Completa la tabla de parámetros.
    mostrar_tabla(parametros, "12_coeficientes_modelo_final.csv")  # Presenta coeficientes para la ecuación del póster.
    coef = ajuste.params  # Permite buscar cada estimación por el nombre de su rezago.
    p_ar, d, q_ma = seleccion["orden"]  # Lee órdenes no estacionales de la especificación seleccionada.
    p_sar, d_s, q_sma, periodo = seleccion["estacional"]  # Lee órdenes estacionales y frecuencia.
    phi = "1" + "".join(f" {-float(coef[f'ar.L{i}']):+.5f} B^{i}" for i in range(1, p_ar + 1))  # Sustituye términos AR con signo negativo.
    theta = "1" + "".join(f" {float(coef[f'ma.L{i}']):+.5f} B^{i}" for i in range(1, q_ma + 1))  # Sustituye términos MA con signo positivo.
    phi_s = "1" + "".join(f" {-float(coef[f'ar.S.L{i * periodo}']):+.5f} B^{i * periodo}" for i in range(1, p_sar + 1))  # Sustituye términos AR estacionales.
    theta_s = "1" + "".join(f" {float(coef[f'ma.S.L{i * periodo}']):+.5f} B^{i * periodo}" for i in range(1, q_sma + 1))  # Sustituye términos MA estacionales.
    izquierda = f"({phi})({phi_s})(1-B)^{d}(1-B^{periodo})^{d_s} ln(Y_t)"  # Compone el lado de las diferencias autorregresivas.
    derecha = f"({theta})({theta_s}) e_t"  # Compone el lado de los errores y sus rezagos.
    ecuacion = [  # Prepara texto listo para incorporarse al póster.
        "ECUACIÓN DEL MODELO (notación de polinomios de rezago)",  # Encabezado de la expresión estimada.
        f"Modelo: {seleccion['nombre']}",  # Identifica los órdenes principales del modelo.
        "Variable: Y_t = toneladas exportadas de café sin tostar",  # Precisa la variable económica.
        "Variable modelada: z_t = ln(Y_t)",  # Aclara el uso de transformación logarítmica.
        "Operador de rezago: B z_t = z_(t-1)",  # Define B para interpretar polinomios.
        "Ecuación con coeficientes sustituidos:",  # Introduce la ecuación numérica, no solamente una fórmula genérica.
        f"{izquierda} = {derecha}",  # Presenta la ecuación final con los parámetros efectivamente estimados.
        "Constante/deriva: excluida (trend='n')",  # Documenta la especificación determinística.
        "Coeficientes y errores estándar:",  # Introduce la tabla textual con incertidumbres.
    ]  # Concluye la introducción de la ecuación.
    ecuacion.extend([f"  {fila.Parametro} = {fila.Coeficiente:+.6f} (EE = {fila.Error_estandar:.6f}; p = {fila.Valor_p:.6g})" for fila in parametros.itertuples(index=False)])  # Sustituye coeficientes y errores estándar.
    ecuacion.append("Interpretación: los parámetros ar y ma describen persistencia en las variaciones del logaritmo; los componentes estacionales operan en rezagos de 12 meses.")  # Vincula la ecuación con la variable.
    print("\n".join(ecuacion))  # Presenta la ecuación y sus coeficientes directamente en la consola.
    print(f"  AIC final: {ajuste.aic:.2f}; BIC final: {ajuste.bic:.2f}; convergió: {ajuste.mle_retvals.get('converged', True)}")  # Indica criterios finales.
    return {"ajuste": ajuste, "parametros": parametros, "ecuacion": ecuacion}  # Entrega estimación y ecuación a etapas siguientes.


# SECCIÓN 8. DIAGNÓSTICO DE RESIDUALES.  # Valida independencia, normalidad y heterocedasticidad.
def seccion_8(seleccion, modelo, args):  # Aplica pruebas estadísticas sobre residuos del modelo final.
    mostrar_etapa(8, "Diagnóstico de residuos del modelo final")  # Señala los supuestos evaluados.
    residuos = residuales_utiles(modelo["ajuste"], seleccion["orden"], seleccion["estacional"])  # Evita los periodos iniciales de estimación.
    gdl = seleccion["orden"][0] + seleccion["orden"][2] + seleccion["estacional"][0] + seleccion["estacional"][2]  # Ajusta los grados de libertad de Ljung-Box.
    ljung = acorr_ljungbox(residuos, lags=[12, 24], model_df=gdl, return_df=True)  # Evalúa correlación residual anual y bianual.
    p_ljung_12 = float(ljung.loc[12, "lb_pvalue"])  # Recupera evidencia de autocorrelación hasta un año.
    p_ljung_24 = float(ljung.loc[24, "lb_pvalue"])  # Recupera evidencia de autocorrelación hasta dos años.
    jb = stats.jarque_bera(residuos)  # Contrasta hipótesis nula de normalidad de errores.
    arch = het_arch(residuos, nlags=12, ddof=gdl)  # Contrasta heterocedasticidad condicional con 12 rezagos.
    tabla = pd.DataFrame([  # Construye una fila por supuesto analizado.
        {"Prueba": "Ljung-Box 12 rezagos", "Valor_p": p_ljung_12, "Hipotesis_nula": "No autocorrelación hasta 12 rezagos", "Rechazar_H0_5pct": p_ljung_12 < ALFA},  # Contrasta el primer año.
        {"Prueba": "Ljung-Box 24 rezagos", "Valor_p": p_ljung_24, "Hipotesis_nula": "No autocorrelación hasta 24 rezagos", "Rechazar_H0_5pct": p_ljung_24 < ALFA},  # Contrasta dos años.
        {"Prueba": "Jarque-Bera", "Valor_p": float(jb.pvalue), "Hipotesis_nula": "Residuos normalmente distribuidos", "Rechazar_H0_5pct": jb.pvalue < ALFA},  # Contrasta normalidad.
        {"Prueba": "ARCH-LM 12 rezagos", "Valor_p": float(arch[1]), "Hipotesis_nula": "Sin efectos ARCH hasta 12 rezagos", "Rechazar_H0_5pct": float(arch[1]) < ALFA},  # Contrasta varianza condicional.
    ])  # Termina la tabla de diagnóstico.
    mostrar_tabla(tabla, "14_pruebas_residuales.csv")  # Presenta tabla resumida para el póster.
    fig, axes = plt.subplots(3, 2, figsize=(12, 11))  # Crea seis diagnósticos visuales en una sola lámina.
    axes[0, 0].plot(residuos.index, residuos, linewidth=0.9, color="#244b73")  # Visualiza cambios de magnitud y choques atípicos.
    axes[0, 0].axhline(0, color="gray", linestyle="--")  # Marca el centro teórico de errores.
    axes[0, 0].set(title="Residuos a través del tiempo", ylabel="Error logarítmico")  # Describe la escala residual.
    plot_acf(residuos, ax=axes[0, 1], lags=36, alpha=0.05, zero=False)  # Busca autocorrelación sin capturar por el modelo.
    axes[0, 1].set(title="FAC de residuos", ylim=(-1, 1))  # Fija la escala de correlación.
    plot_pacf(residuos, ax=axes[1, 0], lags=36, alpha=0.05, zero=False, method="ywm")  # Busca dependencia parcial residual.
    axes[1, 0].set(title="FACP de residuos", ylim=(-1, 1))  # Presenta los rezagos residuales.
    plot_acf(residuos ** 2, ax=axes[1, 1], lags=36, alpha=0.05, zero=False)  # Busca correlación en varianza o volatilidad.
    axes[1, 1].set(title="FAC de residuos al cuadrado", ylim=(-1, 1))  # Facilita detectar componentes ARCH.
    axes[2, 0].hist(residuos, bins=25, density=True, color="#8bb5b7", edgecolor="white", alpha=0.85)  # Presenta la distribución empírica.
    xx = np.linspace(float(residuos.min()), float(residuos.max()), 250)  # Crea una malla para una normal de referencia.
    axes[2, 0].plot(xx, stats.norm.pdf(xx, loc=float(residuos.mean()), scale=float(residuos.std(ddof=1))), color="#b85c38")  # Superpone densidad normal comparable.
    axes[2, 0].set(title="Histograma y distribución normal de referencia", xlabel="Residuo")  # Permite identificar asimetría y colas.
    stats.probplot(residuos, dist="norm", plot=axes[2, 1])  # Compara cuantiles empíricos contra cuantiles normales.
    axes[2, 1].set(title="Gráfico de cuantiles Q-Q")  # Explica el último diagnóstico.
    mostrar_figura("15_diagnostico_residuos.png", args)  # Presenta los seis gráficos para el póster.
    if p_ljung_24 < ALFA:  # Advierte si persiste autocorrelación residual.
        print("  ADVERTENCIA: el modelo NO supera Ljung-Box a 24 rezagos; revisar especificación antes de sostener validez.")  # Evita declarar ruido blanco sin fundamento.
    if float(jb.pvalue) < ALFA:  # Advierte si la normalidad es incompatible con los residuos.
        print("  ADVERTENCIA: Jarque-Bera rechaza normalidad; intervalos gaussianos son aproximados.")  # Explica limitaciones de pronóstico.
    if float(arch[1]) < ALFA:  # Advierte si hay evidencia de heterocedasticidad condicional.
        print("  ADVERTENCIA: ARCH-LM sugiere dependencia de volatilidad; revisar intervalos.")  # Señala necesidad de diagnóstico adicional.
    return {"tabla": tabla, "residuos": residuos}  # Devuelve diagnósticos para conclusiones.


# SECCIÓN 9. PRONÓSTICO PUNTUAL E INTERVALOS A DIEZ MESES.  # Produce resultados entendibles en toneladas.
def seccion_9(serie, modelo, args):  # Pronostica el nivel de toneladas aunque el ARIMA trabaje en logaritmos.
    mostrar_etapa(9, "Pronóstico puntual a diez meses e intervalos 80 % / 95 %")  # Presenta la etapa principal del ejercicio.
    resultado = modelo["ajuste"].get_forecast(steps=HORIZONTE)  # Genera pronósticos para diez meses posteriores a la última fecha.
    media_log = resultado.predicted_mean  # Obtiene la media condicional del logaritmo futuro.
    var_log = resultado.var_pred_mean  # Obtiene la varianza de predicción logarítmica por horizonte.
    mediana = np.exp(media_log)  # Retransforma a la mediana condicional de toneladas.
    media_corregida = np.exp(media_log + 0.5 * var_log)  # Calcula la media aproximada bajo un supuesto lognormal.
    c80 = resultado.conf_int(alpha=0.20)  # Genera límites de predicción al 80 % sobre escala logarítmica.
    c95 = resultado.conf_int(alpha=0.05)  # Genera límites de predicción al 95 % sobre escala logarítmica.
    fechas = pd.date_range(serie.index[-1] + pd.offsets.MonthBegin(1), periods=HORIZONTE, freq="MS")  # Crea los diez meses futuros.
    tabla = pd.DataFrame({  # Crea una fila por mes pronosticado.
        "Mes": fechas,  # Identifica los diez meses siguientes al último dato observado.
        "Pronostico_puntual_media_ton": media_corregida.to_numpy(),  # Media corregida en toneladas.
        "Mediana_pronostico_ton": mediana.to_numpy(),  # Mediana condicional, también expresada en toneladas.
        "Limite_inferior_80_ton": np.exp(c80.iloc[:, 0].to_numpy()),  # Límite inferior al 80 %.
        "Limite_superior_80_ton": np.exp(c80.iloc[:, 1].to_numpy()),  # Límite superior al 80 %.
        "Limite_inferior_95_ton": np.exp(c95.iloc[:, 0].to_numpy()),  # Límite inferior al 95 %.
        "Limite_superior_95_ton": np.exp(c95.iloc[:, 1].to_numpy()),  # Límite superior al 95 %.
    })  # Completa la tabla de pronóstico puntual.
    mostrar_tabla(tabla, "16_PRONOSTICO_PUNTUAL_10_MESES.csv")  # Presenta la tabla central del póster.
    historicos = serie.iloc[-48:]  # Muestra cuatro años para que el pronóstico de diez meses sea visible.
    fig, ax = plt.subplots(figsize=(12, 5.6))  # Crea un gráfico legible para impresión.
    ax.plot(historicos.index, historicos.to_numpy(), label="Observado", color="#244b73", linewidth=1.8)  # Dibuja el historial reciente.
    ax.fill_between(fechas, tabla["Limite_inferior_95_ton"], tabla["Limite_superior_95_ton"], color="#8bb5b7", alpha=0.23, label="Intervalo 95 %")  # Representa incertidumbre más amplia.
    ax.fill_between(fechas, tabla["Limite_inferior_80_ton"], tabla["Limite_superior_80_ton"], color="#8bb5b7", alpha=0.45, label="Intervalo 80 %")  # Representa incertidumbre intermedia.
    ax.plot(fechas, media_corregida.to_numpy(), label="Pronóstico puntual (media)", color="#b85c38", linewidth=2.0, marker="o", markersize=3)  # Destaca el dato puntual por mes.
    ax.axvline(serie.index[-1], color="gray", linestyle="--", linewidth=1, label="Fin de datos reales")  # Marca frontera entre historia y futuro.
    ax.set(title="Exportaciones mensuales de café sin tostar: pronóstico a diez meses", xlabel="Mes", ylabel="Toneladas netas")  # Identifica variable, periodo y unidades.
    ax.legend(ncol=2, fontsize=9, loc="upper left")  # Explica curvas y bandas.
    mostrar_figura("17_PRONOSTICO_10_MESES.png", args)  # Presenta la visualización del pronóstico.
    print("  NOTA: los intervalos gaussianos y la corrección lognormal de la media son aproximaciones condicionadas al modelo.")  # Reconoce límites ante falta de normalidad.
    return tabla  # Entrega los diez pronósticos al informe final.


# ORQUESTACIÓN.  # Ejecuta todo o detiene el proceso en la etapa indicada por el estudiante.
def main():  # Coordina las nueve secciones visibles del análisis.
    args = opciones_programa()  # Lee el modo de operación elegido.
    print("\n>>> VERSIÓN SOLO PANTALLA: SIN CSV, SIN PNG, SIN CARPETAS DE RESULTADOS <<<")  # Identifica el archivo correcto incluso antes de elegir el Excel.
    ruta = elegir_excel(args)  # Selecciona el Excel externo sin necesidad de una carpeta datos/ preexistente.
    print("Tablas: consola | Gráficas: panel Plots de PyCharm o ventanas Matplotlib.")  # Aclara cómo se visualizan las salidas.
    contexto = {}  # Presenta objetos necesarios para etapas siguientes.
    pasos = [  # Enumera funciones que el estudiante puede ejecutar por bloque.
        (1, lambda: seccion_1(ruta, args)),  # Importa y valida Excel.
        (2, lambda: seccion_2(contexto[1], args)),  # Explora la estacionalidad descriptiva.
        (3, lambda: seccion_3(contexto[1], args)),  # Transforma y contrasta estacionariedad.
        (4, lambda: seccion_4(contexto[3], args)),  # Examina FAC y FACP.
        (5, lambda: seccion_5(contexto[3]["log"], args)),  # Estima candidatos definidos manualmente.
        (6, lambda: seccion_6(contexto[5], args)),  # Evalúa diez meses reservados.
        (7, lambda: seccion_7(contexto[3]["log"], contexto[5], args)),  # Reestima y presenta ecuación.
        (8, lambda: seccion_8(contexto[5], contexto[7], args)),  # Verifica diagnósticos residuales.
        (9, lambda: seccion_9(contexto[1], contexto[7], args)),  # Pronostica diez meses siguientes.
    ]  # Concluye las nueve etapas visibles de trabajo.
    for numero, tarea in pasos:  # Ejecuta etapas en el orden metodológico Box-Jenkins.
        if numero > args.hasta:  # Comprueba si el usuario pidió detenerse en una sección.
            break  # Evita ejecutar bloques posteriores.
        contexto[numero] = tarea()  # Ejecuta el bloque y conserva su resultado.
        preguntar_pausa(args, numero)  # Ofrece una pausa didáctica entre etapas.
    print(f"\nEjecución terminada hasta la sección {args.hasta}. No se guardó ningún archivo.")  # Confirma que todas las salidas quedaron en pantalla.
    if not args.sin_graficas:  # Comprueba si el usuario desea observar las figuras en pantalla.
        if matplotlib.get_backend().lower() == "agg":  # Detecta un entorno sin interfaz gráfica.
            print("ATENCIÓN: backend Agg no interactivo. Configura un backend gráfico en PyCharm para ver las figuras.")  # Evita prometer ventanas en servidores.
        else:  # El intérprete ofrece un motor gráfico disponible.
            print("Abriendo todas las gráficas; cierra las ventanas cuando termines de revisarlas.")  # Explica el comportamiento final.
            plt.show()  # Presenta todas las figuras juntas, sin crear PNG ni alterar el Excel.


if __name__ == "__main__":  # Evita ejecutar automáticamente al importar este archivo desde otro programa.
    try:  # Centraliza errores para que se puedan interpretar en la terminal.
        main()  # Ejecuta el análisis con los argumentos indicados.
    except (ValueError, FileNotFoundError, RuntimeError, KeyError, ImportError) as error:  # Captura fallos esperables de datos o configuración.
        print(f"\nERROR: {error}")  # Explica el motivo de la interrupción.
        raise SystemExit(1) from error  # Devuelve un código de salida no exitoso para el sistema operativo.
