from pprint import pprint
import os
import xlrd
import logging
import pandas as pd
import xlsxwriter
import pyinputplus as pyip


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
        self.file = find_file().replace(".xls", "")
        super().__init__()

        self.df_for_merge_raw = pd.read_excel(f'Listados2022\\{self.file}.xls',
                                              f'{self.simil_camel_case()}')  # -->'Cuota Febrero 2022')
        self.listados_listos_para_procesar = self.df_for_merge_raw.iloc[4:-1, :].values
        self.df_listados_nuevos = self.sacar_sueldos_de_activos()
        self.df_central_ya_mergeado = self.mergear_registros()
        self.df_central_ya_mergeado = self.rellenar_datos_faltantes()

    def simil_camel_case(self):
        workbook = xlrd.open_workbook(f'Listados2022\\{self.file}.xls')
        sheet_names = workbook.sheet_names()
        return sheet_names[-2]

    @property
    def regiones_posibles(self):
        regiones_posibles = self.df_principal_raw['REGION'].unique()
        return regiones_posibles

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

            if self.df_central_ya_mergeado.isnull().loc[i, "APELLIDO"] or self.df_central_ya_mergeado.isnull().loc[
                i, "REGION"]:
                data_para_transferir = df_padron_completo.loc[
                    df_padron_completo['LEGAJO'] == row["LEGAJO"], ["APELLIDO", "REGION"]]
                self.df_central_ya_mergeado.loc[i, "APELLIDO"] = data_para_transferir.values[0][0]
                self.df_central_ya_mergeado.loc[i, "REGION"] = self.verificar_regiones(
                    data_para_transferir.values[0][1], data_para_transferir.values[0][0])

        return self.df_central_ya_mergeado

    def guardar_pandas_en_xlsx(self):
        writer = pd.ExcelWriter(f'Registro_ASSA_{self.file}.xlsx', engine='xlsxwriter')
        self.df_central_ya_mergeado.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()


class AsientoContable:
    def __init__(self):
        self.file = find_file()
        self.df_principal_raw = pd.read_excel(f'Listados2022\\{self.file}')
        self.main_processed_df = self.df_principal_raw.loc[::,
                                 ['LEGAJO', 'APELLIDO', 'REGION', self.df_principal_raw.columns[-1]]]
    def __str__(self):
        return f"esta es una prueba para git nada mas brotha"


if __name__ == '__main__':
    asiento_ASSA = AsientoContable()
    print(asiento_ASSA)


    exit()
    listados_base = ListadosAlamcenadosxls()
    listados_nuevo_para_mergear = NuevoListadoxls()
    listados_nuevo_para_mergear.guardar_pandas_en_xlsx()
