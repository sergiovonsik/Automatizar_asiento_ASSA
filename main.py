import pyinputplus as pyip
from package.module import AsientoContable, ListadosAlamcenadosxls, NuevoListadoxls

def main():
    choice = pyip.inputChoice(["merge", "asiento"])
    if choice == "asiento":
        asiento_ASSA = AsientoContable()
        asiento_ASSA.cargar_aportes_en_el_asiento(asiento_ASSA.suma_aportes_cada_region())
        asiento_ASSA.guardar_excel_en_pdf()
    elif choice == "merge":
        listados_base = ListadosAlamcenadosxls()
        listados_nuevo_para_mergear = NuevoListadoxls()
        listados_nuevo_para_mergear.guardar_pandas_en_xlsx()
    main()


if __name__ == '__main__':
    main()