# Tutorial básico de `openpyxl` (Fundamentos)

Este material está pensado para estudiantes que ya manejan:
- `list`
- `set`
- `dict`
- funciones

> En los ejemplos usamos *typing* moderno de Python 3.13+ (`list[str]`, `dict[str, int]`, etc.), sin crear clases.

---

## 1) Instalación

```bash
python -m pip install openpyxl
```

Verificar versión instalada:

```bash
python -c "import openpyxl; print(openpyxl.__version__)"
```

---

## 2) Teoría mínima: workbook, worksheet, celdas y rangos

- **Workbook**: el archivo Excel completo (`.xlsx`).
- **Worksheet**: una hoja dentro del workbook.
- **Celda**: intersección fila/columna (ej: `A1`).
- **Rango**: grupo de celdas (ej: `A1:C10`).

```python
from openpyxl import Workbook

def teoria_basica() -> None:
    wb: Workbook = Workbook()            # workbook
    ws = wb.active                       # worksheet activa
    ws.title = "Datos"                  # cambiar nombre de hoja

    ws["A1"] = "Nombre"                # celda por referencia A1
    ws.cell(row=1, column=2, value="Edad")  # celda por fila/columna

    rango = ws["A1:B1"]                 # rango
    print(rango)

    print(wb.sheetnames)                 # nombres de hojas
    print(len(wb.worksheets))            # cantidad de hojas

teoria_basica()
```

### Acceso a hojas por nombre y por número

```python
from openpyxl import Workbook

def hojas() -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Resumen"
    wb.create_sheet("Notas")
    wb.create_sheet("Promedios")

    por_nombre = wb["Notas"]
    por_indice = wb.worksheets[0]  # primera hoja

    print(por_nombre.title)
    print(por_indice.title)

hojas()
```

---

## 3) Escribir archivos (`.xlsx`)

```python
from openpyxl import Workbook

def escribir_archivo() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"

    # Referencia por letra+número
    ws["A1"] = "Producto"
    ws["B1"] = "Cantidad"

    # Referencia por fila/columna
    ws.cell(row=2, column=1, value="Cuaderno")
    ws.cell(row=2, column=2, value=12)

    ws.cell(row=3, column=1, value="Lápiz")
    ws.cell(row=3, column=2, value=20)

    wb.save("ventas.xlsx")

escribir_archivo()
```

---

## 4) Tipos de datos básicos: texto, números, booleanos, fechas y moneda

```python
from datetime import datetime
from openpyxl import Workbook

def tipos_datos() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tipos"

    ws["A1"] = "Texto"
    ws["B1"] = "Hola"

    ws["A2"] = "Entero"
    ws["B2"] = 42

    ws["A3"] = "Decimal"
    ws["B3"] = 3.1416

    ws["A4"] = "Booleano"
    ws["B4"] = True

    ws["A5"] = "Fecha"
    ws["B5"] = datetime(2026, 6, 7)
    ws["B5"].number_format = "DD/MM/YYYY"

    ws["A6"] = "Moneda"
    ws["B6"] = 1234.5
    ws["B6"].number_format = '"$"#,##0.00'

    wb.save("tipos.xlsx")

tipos_datos()
```

---

## 5) Fórmulas básicas con celdas

```python
from openpyxl import Workbook

def formulas_basicas() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Fórmulas"

    ws["A1"] = "N1"
    ws["B1"] = "N2"
    ws["C1"] = "Suma"

    ws["A2"] = 10
    ws["B2"] = 15
    ws["C2"] = "=A2+B2"

    ws["A3"] = 7
    ws["B3"] = 3
    ws["C3"] = "=SUM(A2:B3)"

    wb.save("formulas.xlsx")

formulas_basicas()
```

---

## 6) Leer archivos y detectar zonas con información

```python
from openpyxl import load_workbook

def leer_archivo(ruta: str) -> None:
    wb = load_workbook(ruta)

    # Seleccionar hoja por nombre
    ws = wb["Ventas"]

    print("Filas con datos:", ws.max_row)
    print("Columnas con datos:", ws.max_column)

    # Recorrer zona con datos
    for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column, values_only=True):
        print(fila)

# leer_archivo("ventas.xlsx")
```

---

## 7) Rangos, rangos con nombre, estilos, bordes y merge

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName

def estilos_y_rangos() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Estilos"

    datos: list[tuple[str, int]] = [("Ana", 85), ("Luis", 92), ("Marta", 78)]

    ws["A1"] = "Nombre"
    ws["B1"] = "Nota"

    for i, (nombre, nota) in enumerate(datos, start=2):
        ws.cell(row=i, column=1, value=nombre)
        ws.cell(row=i, column=2, value=nota)

    # Estilos básicos
    ws["A1"].font = Font(bold=True, italic=True, color="FFFFFF")
    ws["B1"].font = Font(bold=True, color="FFFFFF")

    encabezado_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    ws["A1"].fill = encabezado_fill
    ws["B1"].fill = encabezado_fill

    borde_simple = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for fila in ws["A1:B4"]:
        for celda in fila:
            celda.border = borde_simple

    # Rango con nombre (A2:B4)
    rango_notas = DefinedName(name="RangoNotas", attr_text="Estilos!$A$2:$B$4")
    wb.defined_names.add(rango_notas)

    # Merge
    ws.merge_cells("D1:E1")
    ws["D1"] = "Resumen"
    ws["D1"].font = Font(bold=True)

    wb.save("estilos_rangos.xlsx")

estilos_y_rangos()
```

---

## 8) Tamaños de columnas/filas, orientación, tamaño de hoja y área de impresión

```python
from openpyxl import Workbook
from openpyxl.worksheet.page import PageMargins

def formato_impresion() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Impresión"

    filas: list[dict[str, str | int]] = [
        {"producto": "Cuaderno universitario", "cantidad": 12},
        {"producto": "Lápiz HB", "cantidad": 20},
        {"producto": "Borrador", "cantidad": 7},
    ]

    ws["A1"] = "Producto"
    ws["B1"] = "Cantidad"

    for i, item in enumerate(filas, start=2):
        ws.cell(row=i, column=1, value=item["producto"])
        ws.cell(row=i, column=2, value=item["cantidad"])

    # Ancho manual de columnas
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12

    # Alto manual de fila
    ws.row_dimensions[1].height = 24

    # Ajuste "automático" básico (calculado)
    for col in ["A", "B"]:
        max_len = 0
        for celda in ws[col]:
            valor = "" if celda.value is None else str(celda.value)
            max_len = max(max_len, len(valor))
        ws.column_dimensions[col].width = max_len + 2

    # Configuración de página
    ws.page_setup.orientation = "landscape"   # horizontal
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)

    # Área de impresión
    ws.print_area = "A1:B4"

    wb.save("impresion.xlsx")

formato_impresion()
```

---

## 9) Gráficos con datos

```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

def grafico_barras() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Gráfico"

    ws.append(["Mes", "Ventas"])
    ws.append(["Enero", 1200])
    ws.append(["Febrero", 1500])
    ws.append(["Marzo", 1700])
    ws.append(["Abril", 1600])

    datos = Reference(ws, min_col=2, min_row=1, max_row=5)
    categorias = Reference(ws, min_col=1, min_row=2, max_row=5)

    chart = BarChart()
    chart.title = "Ventas por mes"
    chart.y_axis.title = "Monto"
    chart.x_axis.title = "Mes"
    chart.add_data(datos, titles_from_data=True)
    chart.set_categories(categorias)

    ws.add_chart(chart, "D2")

    wb.save("grafico.xlsx")

grafico_barras()
```

---

## Sugerencia didáctica de orden para clase

1. Instalación y teoría básica (`Workbook`, `Worksheet`, celdas, rangos)
2. Escritura y guardado de archivos
3. Tipos de datos y formato de moneda/fecha
4. Lectura de archivos y recorrido de datos
5. Fórmulas y rangos con nombre
6. Estilos, bordes y merge
7. Diseño de impresión (ancho, alto, orientación, área)
8. Gráficos

Con esto tienes una ruta completa, desde lo más básico hasta reportes visuales en Excel usando `openpyxl`.
