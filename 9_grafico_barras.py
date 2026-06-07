"""
Crear un gráfico de barras con openpyxl.
Ejecutar: python 9_grafico_barras.py
"""
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
    print("Archivo grafico.xlsx guardado.")


if __name__ == "__main__":
    grafico_barras()
