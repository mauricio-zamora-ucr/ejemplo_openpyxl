"""
Fórmulas básicas usando celdas.
Ejecutar: python 5_formulas_basicas.py
"""
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
    print("Archivo formulas.xlsx guardado.")


if __name__ == "__main__":
    formulas_basicas()
