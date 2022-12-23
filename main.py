from pprint import pprint
import os
import openpyxl
import xlrd
import logging
import pandas as pd
import xlsxwriter
import pyinputplus as pyip
from win32com import client


def find_file():
    archivos_disponibles = [entry.name for entry in os.scandir('Listados2022') if entry.is_file()]
    return pyip.inputChoice([*archivos_disponibles])


class ListadosAlamcenadosxls:
    def __init__(self):
        # file = find_file()
        self.df_principal_raw = pd.read_excel(f'Listados2022\\ASSATotalDeMeses.xls')
        self.main_processed_df = self.df_principal_raw.loc[::,
                                 ['LEGAJO', 'APELLIDO', 'REGION', self.df_principal_raw.columns[-1]]]

    def mostrar_elementos_del_array(self):
        list_tabla_agentes = self.main_processed_df.values.tolist()
        pprint(list_tabla_agentes)


class NuevoListadoxls(ListadosAlamcenadosxls):
    def __init__(self):
        super().__init__()
        self.file = find_file().replace(".xls", "")
        self.df_for_merge_raw = pd.read_excel(f'Listados2022\\{self.file}.xls',
                                              f'{self.obtener_sheet_name()}')  # -->'Cuota Febrero 2022')
        self.listados_listos_para_procesar = self.df_for_merge_raw.iloc[4:-1, :].values
        self.df_listados_nuevos = self.sacar_sueldos_de_activos()
        self.df_central_ya_mergeado = self.mergear_registros()
        self.df_central_ya_mergeado = self.rellenar_datos_faltantes()

    @property
    def regiones_posibles(self):
        regiones_posibles = self.df_principal_raw['REGION'].unique()
        return regiones_posibles

    def obtener_sheet_name(self):
        workbook = xlrd.open_workbook(f'Listados2022\\{self.file}.xls')
        sheet_names = workbook.sheet_names()
        return sheet_names[-2]

    def verificar_regiones(self, str_region, sujeto):
        region = str_region.replace(" ", "").upper()
        if region not in self.regiones_posibles:
            print(f'la region \'{region}\' del aportante  \'{sujeto}\' no existe, ingresa unas de las posibles:')
            if str_region == 'Santa Fe ':
                return "SANTAFE"
            region = pyip.inputChoice([*self.regiones_posibles])
            return region
        return region

    def __str__(self):
        return f"{self.listados_listos_para_procesar}"

    def sacar_sueldos_de_activos(self):
        sueldos_personal_registrado = list()
        for legajo, apellido, sueldo in self.listados_listos_para_procesar:
            sueldos_personal_registrado.append([legajo, sueldo])
            if legajo not in self.main_processed_df.loc[::, "LEGAJO"].values:
                print(f"NUEVITOS: {legajo, apellido, sueldo}")

        df_padron_completo = pd.DataFrame(sueldos_personal_registrado,
                                          columns=["LEGAJO", self.file])  # ->"FEBRERO 2022"
        return df_padron_completo

    def mergear_registros(self):
        dataframe_para_guardar = self.df_principal_raw.merge(self.df_listados_nuevos, on='LEGAJO', how='outer')
        return dataframe_para_guardar

    def rellenar_datos_faltantes(self):
        padron_con_datos_raw = pd.read_excel(f'Listados2022\\PADRON ASSA.xls')
        padron_con_datos_procesados = padron_con_datos_raw.loc[::, ['Legajo', 'Apellido y Nombre', 'Ubicación']]
        listado_activos_apellido_regiones = list()
        for legajo, apellido, region in padron_con_datos_procesados.values:
            pattern = region.split('-')
            region = pattern[0]
            listado_activos_apellido_regiones.append([legajo, apellido, region])

        df_padron_completo = pd.DataFrame(listado_activos_apellido_regiones, columns=["LEGAJO", "APELLIDO", "REGION"])

        for i, row in self.df_central_ya_mergeado.iterrows():

            if self.df_central_ya_mergeado.isnull().loc[i, "APELLIDO"] or self.df_central_ya_mergeado.isnull().loc[i, "REGION"]:
                data_para_transferir = df_padron_completo.loc[
                    df_padron_completo['LEGAJO'] == row["LEGAJO"], ["APELLIDO", "REGION"]]
                self.df_central_ya_mergeado.loc[i, "APELLIDO"] = data_para_transferir.values[0][0]
                self.df_central_ya_mergeado.loc[i, "REGION"] = self.verificar_regiones(
                    data_para_transferir.values[0][1], data_para_transferir.values[0][0])

        return self.df_central_ya_mergeado

    def guardar_pandas_en_xlsx(self):
        writer = pd.ExcelWriter(f'{self.file} PADRON_PYTHON_ASSA.xlsx', engine='xlsxwriter')
        self.df_central_ya_mergeado.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()


class AsientoContable:
    def __init__(self):
        self.file = None
        self.excel_asiento_contable = openpyxl.load_workbook(f'AsientosContables\\Template Asiento.xlsx')
        self.template_asiento = self.excel_asiento_contable['ASIENTO']
        self.aportes_de_listados = self.obtener_aportes_de_listados

    @property
    def obtener_aportes_de_listados(self):
        archivos_disponibles = [entry.name for entry in os.scandir() if entry.is_file()]
        file = pyip.inputChoice([*archivos_disponibles])  # 'Registro_ASSA_FEBRERO 2022.xlsx'  #
        df = pd.read_excel(f'{file}', sheet_name=0)
        self.file = file.replace(" PADRON_PYTHON_ASSA.xlsx", "")
        return df.iloc[::, [2, -1]]

    @property
    def regiones(self):
        columna_regiones = self.aportes_de_listados['REGION'].unique()
        localidades = dict()
        for i in columna_regiones:
            localidades[i] = 0
        return localidades
    def suma_aportes_cada_region(self):
        grouped = self.aportes_de_listados.groupby('REGION')
        ultima_columna = self.aportes_de_listados.columns[-1]
        result = grouped[ultima_columna].sum()
        aportes_por_localidad_dict = result.to_dict()
        return aportes_por_localidad_dict

    def cargar_aportes_en_el_asiento(self, hash_table):
        for i in range(9, 21):
            print(self.template_asiento[f'D{i}'].value)
            self.template_asiento[f'D{i}'] = hash_table[self.template_asiento[f'D{i}'].value]
            print(self.template_asiento[f'D{i}'].value)
        print(self.template_asiento[f'C26'].value)
        self.template_asiento[f'C26'] = self.file
        print(self.template_asiento[f'C26'].value)
        self.excel_asiento_contable.save(f'AsientosContables\\asiento_ASSA_mes_{self.file} .xlsx')
        print("***ARCHIVO GUARDADO***")


    def guardar_excel_en_pdf(self):

        # Open Microsoft Excel
        excel = client.Dispatch('Excel.Application')

        # Read Excel File
        sheets = excel.Workbooks.Open(f"C:\\Users\\Sergio Nicolas\\PycharmProjects\\pythonProject\AsientosContables"
                                      f"\\asiento_ASSA_mes_{self.file} .xlsx")
        work_sheets = sheets.Worksheets[0]

        # Convert into PDF File
        work_sheets.ExportAsFixedFormat(0, f"C:\\Users\\Sergio Nicolas\\PycharmProjects\\pythonProject\AsientosContables"
                                      f"\\PDF_para_imprimir{self.file}")



if __name__ == '__main__':
    choice = pyip.inputChoice(["Cargar listado a padron", "Cargar asiento contable"])
    if choice == "Cargar asiento contable":
        asiento_ASSA = AsientoContable()
        asiento_ASSA.cargar_aportes_en_el_asiento(asiento_ASSA.suma_aportes_cada_region())
        asiento_ASSA.guardar_excel_en_pdf()
    else:
        listados_base = ListadosAlamcenadosxls()
        listados_nuevo_para_mergear = NuevoListadoxls()
        listados_nuevo_para_mergear.guardar_pandas_en_xlsx()
