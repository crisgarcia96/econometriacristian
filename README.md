# Exportaciones colombianas de café sin tostar: metodología Box-Jenkins

Proyecto académico de Econometría II. Analiza el peso neto mensual exportado de café sin tostar por Colombia, correspondiente al agregado de las subpartidas **0901111000, 0901119000 y 0901120000**.

## Datos

- Fuente de trabajo: `datos/Exportaciones de café.xlsx` (agregado mensual provisto por el equipo, hoja `Detalle`).
- Frecuencia mensual, enero de 2009 a agosto de 2026, 212 observaciones.
- Variable: `Peso en Toneladas Netas` (no valor monetario).
- Archivos en `resultados/`: tablas y gráficos de la ejecución académica original, conservados para consulta. El script principal se ejecuta en pantalla y **no sobrescribe esos archivos**.

## Preparación

Es necesario instalar Python y tener conectividad para instalar las dependencias:

```bash
python -m pip install -r requirements.txt
```

## Ejecución reproducible desde el repositorio clonado

Ejecución sin ventanas gráficas, apta para la verificación rápida por terminal:

```bash
python analisis_cafe.py --sin-dialogo --sin-graficas
```

Ejecución habitual: solicita seleccionar la base Excel y abre las figuras en ventanas (o en el panel de gráficos del IDE):

```bash
python analisis_cafe.py
```

Para revisar las etapas manualmente:

```bash
python analisis_cafe.py --sin-dialogo --paso-a-paso
```

Para indicar explícitamente otro archivo:

```bash
python analisis_cafe.py --excel "ruta/a/otra_base.xlsx" --sin-graficas
```

El archivo debe contener una hoja `Detalle` con columnas `Mes` y `Peso en Toneladas Netas`; no se requieren rutas absolutas ni modificar el código. El script valida frecuencia, continuidad y valores positivos.

## Metodología

Nueve secciones: carga y validación; estacionalidad; transformaciones logarítmicas y ADF/KPSS; FAC/FACP; comparación manual de ARIMA con componentes estacionales; validación fuera de muestra a diez meses; estimación definitiva y ecuación; diagnóstico residual; pronósticos puntuales e intervalos 80 % y 95 %.

La elección de especificaciones no utiliza `auto_arima`. Los resultados pueden variar ligeramente según las versiones de los paquetes y el entorno de cómputo.

## Estructura

- `analisis_cafe.py`: código completo y comentado por instrucciones; **solo muestra** tablas y gráficos.
- `datos/`: hoja Excel que permite reproducir el ejercicio.
- `resultados/`: figuras y tablas históricas que acompañan el póster, no se regeneran automáticamente en la versión de pantalla.
- `requirements.txt`: dependencias para instalar el ambiente.

## Alcance y limitaciones

El modelo es un ejercicio econométrico académico, no una proyección oficial. La validación a diez meses reportó limitaciones predictivas; los resultados deberán interpretarse con sus intervalos y supuestos. Es necesario citar en la sustentación la fuente estadística original de la base agregada.
