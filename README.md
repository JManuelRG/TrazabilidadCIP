# TrazabilidadCIP
#Sistema digital de trazabilidad de ciclos CIP y control de la vigencia sanitaria de tanques móviles en manufactura cosmética
## Presenta:
## Dalia Itzel Coss López


	# Asesores	
## Dr. José Manuel Rivera Garnica
## M. en C. Juan Manuel Salgado Román
## Lic. Alejandro Chávez Araujo


# Reporte automatizado de la bitácora digital de tanques móviles

Script en Python que lee la base de datos exportada de una aplicación de Power Apps (lista de SharePoint en formato `.xlsx`) usada para registrar los ciclos de limpieza de tanques móviles en una planta de producción, y genera de forma automática un **reporte en PDF** con indicadores de uso, gráficos y una revisión de la calidad de los datos.

## Características

- Lectura y limpieza de la base: fechas con mes abreviado en español (`04-ago-2026`), horas en formato AM/PM y unificación de nombres de usuario.
- Indicadores de uso: registros totales, tanques distintos, días con captura, usuarios activos y registros por tanque.
- Gráficos de uso:
  - Calendarios de actividad diaria con [`calplot`](https://github.com/tomkwok/calplot) (total y por tipo de limpieza).
  - Registros por día de la semana y por hora.
  - Registros por usuario y tipo de limpieza.
  - Tanques con más registros e intervalos entre registros consecutivos del mismo tanque.
- Revisión de la calidad de los datos de sesión: horas faltantes, sesiones abiertas y cierres idénticos repetidos.
- Generación del PDF con [`reportlab`](https://www.reportlab.com/).
- Si `calplot` no está instalado o falla, el script avisa en consola y dibuja un calendario equivalente con `matplotlib`, de modo que el reporte se genera de todos modos.

## Requisitos

- Python 3.9 o superior
- Librerías:

```bash
pip install pandas openpyxl matplotlib reportlab calplot scipy
```

## Datos de entrada

El script espera un archivo `.xlsx` (una sola hoja, una fila por ciclo registrado) con, al menos, estas columnas:

| Columna | Descripción |
|---|---|
| `Fecha_Registro` | Fecha del registro (por ejemplo, `04-ago-2026`) |
| `Hora_Registro` | Hora del registro (por ejemplo, `10:57:58 AM`) |
| `Número_TM` | Número del tanque móvil |
| `Tipo_limpieza` | Tipo de limpieza registrada |
| `Registro_Usuario` | Usuario que capturó el registro |
| `Hora_Inicio de sesion` | Hora de inicio de sesión en la app |
| `Hora_Cierre de sesion` | Hora de cierre de sesión en la app |

Si existe una columna `ID`, se eliminan los registros duplicados por ese campo. Si falta alguna de las columnas requeridas, el script se detiene e indica cuáles faltan.

> **Importante:** la base de datos proviene de una operación industrial real y **no se incluye en este repositorio**. Antes de subir cambios, confirma que no estás agregando archivos `.xlsx`, PDF generados ni datos con nombres de personas. Se recomienda añadirlos al `.gitignore`:
>
> ```
> *.xlsx
> Reporte_tanques_moviles_*.pdf
> ```

## Uso

### En Spyder

1. Abre `reporte_tanques_moviles.py`.
2. Ajusta `RUTA_EXCEL` con la ruta completa de tu archivo. Si el Excel está en el directorio de trabajo de Spyder, basta con el nombre.
3. Ejecuta el script completo (`F5`).

### Desde la terminal

```bash
python reporte_tanques_moviles.py
```

El PDF se guarda en `CARPETA_SALIDA` con la fecha y hora en el nombre (`Reporte_tanques_moviles_AAAAMMDD_HHMM.pdf`), así que cada ejecución genera un archivo nuevo. Para actualizar el reporte con datos nuevos, reemplaza el Excel y vuelve a ejecutar.

## Configuración

Los parámetros se editan al inicio del script:

| Parámetro | Descripción | Valor por defecto |
|---|---|---|
| `RUTA_EXCEL` | Ruta del archivo de datos | `Copia_de_monitoreo_de_tanques_moviles_CL.xlsx` |
| `CARPETA_SALIDA` | Carpeta donde se guarda el PDF | Directorio actual |
| `UMBRAL_DUPLICADO_H` | Horas por debajo de las cuales dos registros del mismo tanque se consideran una probable captura duplicada | `1` |
| `UMBRAL_REPROCESO_H` | Horas por debajo de las cuales dos registros del mismo tanque se marcan como intervalo corto a revisar | `24` |
| `MAX_SESION_MIN` | Duración máxima (min) de una sesión para considerarla confiable | `120` |
| `UMBRAL_CIERRE_REPETIDO` | Veces que debe repetirse una misma hora de cierre para marcarla como posible edición masiva | `10` |
| `ALIAS_USUARIOS` | Diccionario para unificar nombres de usuario escritos de forma distinta | Un ejemplo editable |

Los umbrales son criterios de análisis configurables, no definiciones normativas. Ajústalos según el proceso real de la planta.

## Contenido del reporte

1. Encabezado con indicadores clave y hallazgos principales redactados automáticamente.
2. Uso de la aplicación en el tiempo (calendarios, día de la semana y hora).
3. Registros por usuario.
4. Tanques móviles e intervalos entre registros.
5. Calidad de los datos de sesión.
6. Notas metodológicas.

## Estructura del script

| Bloque | Función principal |
|---|---|
| Lectura y limpieza | `cargar_datos()` |
| Indicadores | `calcular()` |
| Gráficos | `generar_graficos()`, `calendario()` |
| PDF | `construir_pdf()` |
| Ejecución | `main()` |

## Limitaciones

- La base solo registra uso de la aplicación. No incluye consumo de agua, receta CIP ni producto previo, por lo que el reporte mide uso y frecuencia de registro, no ahorro de agua.
- Un intervalo corto entre registros del mismo tanque es un indicio que debe validarse con el personal de operación; no prueba por sí solo un re-lavado evitable.
- `calplot` dibuja el año calendario completo, por lo que con pocos meses de datos gran parte del calendario aparece vacío.

## Licencia

Por definir.
