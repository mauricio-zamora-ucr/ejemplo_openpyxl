
def mostrar_comandos() -> None:
    print("Comandos para instalar openpyxl:\n")
    print("python -m pip install openpyxl\n")
    print('Para verificar la versión:')
    print('python -c "import openpyxl; print(openpyxl.__version__)"')


def verificar_version() -> None:
    import openpyxl

    print("\nopenpyxl instalado. Versión:", openpyxl.__version__)



if __name__ == "__main__":
    mostrar_comandos()
    verificar_version()
