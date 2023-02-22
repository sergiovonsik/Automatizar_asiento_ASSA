import pyinputplus as pyip

from package.module import AsientoContable, ListadosAlamcenadosxls, NuevoListadoxls



if __name__ == '__main__':
    choice = pyip.inputChoice(["Mergear_listado_a_Padron", "Crear_asiento_contable"])
    if choice == "Crear_asiento_contable":
        asiento_ASSA = AsientoContable()
        asiento_ASSA.cargar_aportes_en_el_asiento(asiento_ASSA.suma_aportes_cada_region())
        asiento_ASSA.guardar_excel_en_pdf()
    else:
        listados_base = ListadosAlamcenadosxls()
        listados_nuevo_para_mergear = NuevoListadoxls()
        listados_nuevo_para_mergear.guardar_pandas_en_xlsx()
