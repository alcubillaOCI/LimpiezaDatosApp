import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import zipfile
import io

# Título de la aplicación
st.title("Aplicación de Carga y Modificación de Archivos")

# %%% Sección para cargar el archivo
st.subheader("Etiqueta de denuncias", divider='orange')
df_up = st.file_uploader(
    "Archivo xlsx con las páginas 'Etiquetas de denuncias', 'Denuncias' y 'Denuncias desechadas' del reporte de IA", type=["xlsx"])

st.subheader("Zip file", divider='violet')
zip = st.file_uploader("Cargar el archivo zip", type=["zip"])
on = st.toggle('Archivos xlsx')
if on:
    st.write('Todas las bases de datos tienen formato donde el primer renglón es el nombre de columna y solo una hoja')
    st.write('* **Relacion Medidas formativas:** relación entre valores de Medidas formativas y su limpieza')
    st.write('* **Regiones:** relación entre campus y regiones actualizadas')
    st.write('* **Relacion Sanciones por FIA:** relación entre descripción, categorización por individual o grupal')
    st.write('* **BD Insc-Alu:** alumnos inscritos')
    st.write('* **Carreras:** valor anterior y nuevo de las siglas de carreras')

if zip:
    # Descomprimir el archivo .zip
    with zipfile.ZipFile(zip, 'r') as zip_ref:
        zip_ref.extractall(".")

    # Nombre carpeta comprimida
    carpeta = zip_ref.namelist()[0]

    # Crear dataframes según el nombre
    name = ['Relacion Medidas formativas', 'Regiones',
            'Relacion Sanciones por FIA', 'BD Insc-Alu', 'Carreras']

    if carpeta + name[0] + '.xlsx' in zip_ref.namelist():
        med = pd.read_excel(carpeta + name[0] + '.xlsx')
    else:
        st.write('El archivo "' + name[0] +
                 '" no se encuentra en el archivo .zip.')

    if carpeta + name[1] + '.xlsx' in zip_ref.namelist():
        regiones = pd.read_excel(carpeta + name[1] + '.xlsx')
    else:
        st.write('El archivo "' + name[1] +
                 '" no se encuentra en el archivo .zip.')

    if carpeta + name[2] + '.xlsx' in zip_ref.namelist():
        cierre = pd.read_excel(carpeta + name[2] + '.xlsx')
    else:
        st.write('El archivo "' + name[2] +
                 '" no se encuentra en el archivo .zip.')

    if carpeta + name[3] + '.xlsx' in zip_ref.namelist():
        insc_al = pd.read_excel(carpeta + name[3] + '.xlsx')
    else:
        st.write('El archivo "' + name[3] +
                 '" no se encuentra en el archivo .zip.')
    if carpeta + name[4] + '.xlsx' in zip_ref.namelist():
        carreras = pd.read_excel(carpeta + name[4] + '.xlsx')
    else:
        st.write('El archivo "' + name[4] +
                 '" no se encuentra en el archivo .zip.')

# %%% Si están los archivos cargados
if (df_up is not None) and (zip is not None):
    # %%% Leer archivos cargados
    df = pd.read_excel(df_up, sheet_name='Etiquetas de denuncias')
    den = pd.read_excel(df_up, sheet_name='Denuncias')

    # %%% DATOS --------
    df.columns = ['Folio', 'Cat', 'Etiqueta']

    # Función para juntar los valores de una serie con '-'
    def concat_values(series):
        return '-'.join(map(str, series))

    df = df.pivot_table(index='Folio', columns='Cat',
                        values='Etiqueta', aggfunc=concat_values)

    col = ['C Nómina Del Reportado', 'E Calificación Asignada', 'E Cantidad De Personas Implicadas En La Falta', 'E Consecuencia Disciplinaria', 'E Correo Institucional De La Persona Que Reporta', 'E Crn De La Clase', 'E Lugar/Actividad Donde Ocurrió La Falta', 'E Materia/Unidad De Formación Donde Ocurrió La Falta',
           'E Matrícula Del Estudiante Reportado', 'E Medida Formativa Asignada Por El Ciac', 'E Modelo Educativo', 'E Programa O Carrera Del Estudiante Reportado', 'E Semestre (#)', 'Ethos Previos', 'Nombre Del Reportado', 'Nombre Y Id Del Que Reporta', 'Periodo/Ciclo Recepción', 'Región', 'Sexo De La Persona Reportada', 'Tipo De Cierre/Resolución']
    df = df[col]
    den.rename(columns={'Folio Interno de Denuncia': 'Folio', 'Mes de recepción': 'Mes',
               'Año de recepción': 'Year', 'Unidad de Negocio': 'Nivel'}, inplace=True)

    # Desechar registros que no sean de Tec de Mty
    folio_no_profesional = den[den['Empresa'] !=
                               'ETHOS Estudiantes - Tec de Monterrey']['Folio']
    desechadas = den[den['Estatus de la denuncia'] == 'Desechada']['Folio']

    df = df[(~df.index.isin(folio_no_profesional))
            & (~df.index.isin(desechadas))]
    den = den[(~den.Folio.isin(folio_no_profesional))
              & (~den.Folio.isin(desechadas))]

    excel_file = pd.ExcelFile(df_up)
    if 'Denuncias desechadas' in excel_file.sheet_names:
        desecho = pd.read_excel(df_up, sheet_name='Denuncias desechadas')
        df = df[~df.index.isin(desecho.Folio)]
        den = den[~den.Folio.isin(desecho.Folio)]

    # %%% FUNCIONES
    # Función para modificar ciertas palabras en los valores para fines de formato
    def modificar_palabra(cadena, val):
        if type(cadena) is str:
            palabra_modificada = cadena
            for palabra, sustituto in val.items():
                # Se modifica cada palabra con los valores del diccionario
                palabra_modificada = palabra_modificada.replace(
                    palabra, sustituto)
            return palabra_modificada
        else:
            return cadena

    # Función para quitar '_' de las columnas y dar formato de mayuscula
    def formato_strings_cap(nombre):
        if (type(nombre) is str):
            # Reemplazar '_' por ' ' y empezar con mayusuclas
            return nombre.replace('_', ' ').title()
        else:
            return nombre

    # Función para quitar '_' de las columnas
    def formato_strings(nombre):
        if (type(nombre) is str):
            # Reemplazar '_' por ' ' y empezar con mayusuclas
            return nombre.replace('_', ' ')
        else:
            return nombre
        
    # %%% Formato de la localización apropiado
    den['Localización'].replace({'Cd. Juárez': 'Ciudad Juárez',
                                'Celaya': 'Prepa Celaya',
                                 'Cumbres': 'Prepa Cumbres',
                                 'Garza Laguera': 'Prepa Eugenio Garza Lagüera',
                                 'Garza Sada': 'Prepa Eugenio Garza Sada',
                                 'Santa Catarina': 'Prepa Santa Catarina',
                                 'Valle Alto': 'Prepa Valle Alto'}, inplace=True)

    insc_al = insc_al.set_index('Matrícula')  # alumnos inscritos

    # %%% 1 - E Calificación Asignada
    df['E Calificación Asignada'].replace({'na': np.nan, 'n_a': np.nan, 'nr': np.nan, 'cero': '0',
                                           'np': np.nan, 'da': np.nan, 'ne': np.nan, 'sc': np.nan,
                                           'in': np.nan, 'sn': np.nan, 'se': np.nan, 'nd': np.nan,
                                           'n_e': np.nan, 'no_aplica': np.nan, 'no_se_involucra': np.nan}, inplace=True)

    # %%% 2 - E Cantidad De Personas Implicadas En La Falta
    df['E Cantidad De Personas Implicadas En La Falta'] = df['E Cantidad De Personas Implicadas En La Falta'].str.extract(
        r'(\d+)')[0]
    df['E Cantidad De Personas Implicadas En La Falta'] = df['E Cantidad De Personas Implicadas En La Falta'].astype(
        'Int64')

    # %%% 4 - E Correo Institucional De La Persona Que Reporta
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_tec_mx", r"\1@tec.mx", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_tec_com_mx", r"\1@tec.mx", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_tecmx", r"\1@tec.mx", regex=True, inplace = True)

    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_itesm_mx", r"\1@itesm.mx", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_tecmilenio_mx", r"\1@tecmilenio.mx", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_hotmail_com", r"\1@hotmail.com", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace("(.*)_gmail_com", r"\1@gmail.com", regex=True, inplace = True)
    df['E Correo Institucional De La Persona Que Reporta'].replace(['na','sin_informacion'], np.nan, inplace = True)

    # %%% 8 - E Matrícula Del Estudiante Reportado
    # Pasar a nulo los valores que no son matrículas
    # ------------------------------------------------
    # Df con condicional si la columna Matrícula contiene 'ax' donde x es un número
    a = pd.DataFrame(
        df['E Matrícula Del Estudiante Reportado'].str.contains(r'a\d+'))
    # Índices de valores no válidos
    index = a[a['E Matrícula Del Estudiante Reportado'] == False].index
    # Convertir a nulo
    df.loc[index, 'E Matrícula Del Estudiante Reportado'] = np.nan

    # ------------------------------------------------
    # Formato general a minuscula
    df['E Matrícula Del Estudiante Reportado'] = df['E Matrícula Del Estudiante Reportado'].str.lower()

    # Si contiene letras distintas a 'a'
    # Índices
    index_mat_no_a = df[df['E Matrícula Del Estudiante Reportado'].str.contains(
        '|'.join(list(map(chr, range(98, 123))))) == True].index  # minusculas sin la a)
    # Reemplazar a valores correctos
    for i in index_mat_no_a:
        # reasignar valor correcto de matrículas
        df.loc[i, 'E Matrícula Del Estudiante Reportado'] = '_'.join(
            re.findall(r'(a\d{8})', df.loc[i, 'E Matrícula Del Estudiante Reportado']))

    # Cantidad de personas nulo y con matrículas registradas
    # ------------------------------------------------
    # Folios de los registros con cantidad de personas nulo y matrículas no nulas
    index_cantidad_matriculas = df[(df['E Cantidad De Personas Implicadas En La Falta'].isnull()) & (
        df['E Matrícula Del Estudiante Reportado'].notnull())].index
    df.loc[index_cantidad_matriculas, 'E Cantidad De Personas Implicadas En La Falta'] = df.loc[index_cantidad_matriculas,
                                                                                                'E Matrícula Del Estudiante Reportado'].str.count('a').astype('Int64')

    # Matrículas registradas > cantidad de personas. Se cambia por la cantidad de matriculas registradas
    # ------------------------------------------------
    index_mat_cant = df[df['E Matrícula Del Estudiante Reportado'].str.count(
        'a').astype('Int64') > df['E Cantidad De Personas Implicadas En La Falta']].index
    df.loc[index_mat_cant, 'E Cantidad De Personas Implicadas En La Falta'] = df.loc[index_mat_cant,
                                                                                     'E Matrícula Del Estudiante Reportado'].str.count('a')

    # %%% 9 - E Medida Formativa Asignada Por El Ciac
    # Diccionario con la medida y categoría correspondiente
    med_dict = med.set_index('Descripción')['Categoría'].to_dict()
    df['E Medida Formativa Asignada Por El Ciac'].replace(
        med_dict, inplace=True)

    # %%%10 - E Modelo Educativo
    # indices de no nulos
    ind = df[~df['E Modelo Educativo'].isnull()].index.values
    mod_nonull = df.loc[ind, 'E Modelo Educativo']  # series de los no nulos

    # True/False de cambiar según la escritura
    tec21 = (mod_nonull.str.contains("tec") & mod_nonull.str.contains(
        "21") & ~(mod_nonull.str.contains("planes")))
    planes = (mod_nonull.str.contains("planes") & mod_nonull.str.contains(
        "anteriores") & ~(mod_nonull.str.contains("21")))
    tec_planes = (mod_nonull.str.contains("tec") & mod_nonull.str.contains(
        "21") & mod_nonull.str.contains("planes") & mod_nonull.str.contains("anteriores"))

    # Cambiar los valores
    mod_nonull = mod_nonull.where(~tec21, 'tec21')
    mod_nonull = mod_nonull.where(~planes, 'planes_anteriores')
    mod_nonull = mod_nonull.where(~tec_planes, 'tec21-planes_anteriores')
    mod_nonull = mod_nonull.where(tec21 | tec_planes | planes, np.nan)

    df.loc[ind, 'E Modelo Educativo'] = mod_nonull
    # %%% 11 - E Programa O Carrera Del Estudiante Reportado
    # Función para limpiar y juntar la carrera del estudiante reportado, devuelve las carreras solo con letras
    def limpiar_programa(cadena):
        # ignorar nan
        if (type(cadena) is str):
            # solo dejar letras
            cadena_letra = ''.join(
                letra for letra in cadena if letra.isalpha())
            return cadena_letra
        else:
            return cadena
        
    df['E Programa O Carrera Del Estudiante Reportado'] = df['E Programa O Carrera Del Estudiante Reportado'].apply(
        limpiar_programa)
    df['E Programa O Carrera Del Estudiante Reportado'].replace(
        ['no_identificado', 'ernestodiezmartinezguzmanneg', 'na', 'noidentificado', 'pendiente'], np.nan, inplace=True)

    carr_dict = carreras.set_index('Valor')['Nuevo'].to_dict()
    df['E Programa O Carrera Del Estudiante Reportado'].replace(
        carr_dict, inplace=True)

    # %%% 12 - E Semestre (#)
    def formato_semestre(valor):
        # Si es un número
        if isinstance(valor, str) and valor.isdigit():
            valor_numerico = int(valor)
            # Si es mayor a 10, se anula
            if valor_numerico > 10:
                return np.nan  # Reemplazar por el valor que desees
            else:
                return valor
        # Si no es un número y tiene 'remedial'
        elif isinstance(valor, str) and ('remedial' in valor):
            valor = valor.replace('remedial', '')
        # Cualquier otro valor se hace nulo
        else:
            return np.nan

    df['E Semestre (#)'] = df['E Semestre (#)'].apply(formato_semestre)

    df['E Semestre (#)'] = df['E Semestre (#)'].astype('Int64')

    # %%% 13 - Ethos previos
    ind = df[~df['Ethos Previos'].isnull()].index.values
    et_nonull = df.loc[ind, 'Ethos Previos']  # series de los no nulos

    # Valores donde si tienen casos como 'si-no' o 'no-identificado'
    nans = (et_nonull.str.contains("no") & ((et_nonull.str.contains(
        "identificado")) | et_nonull.str.contains("si")))
    # Valores donde se tienen casos con 'si' con otra estructura que no incluya el 'no'
    si = (et_nonull.str.contains("si") & ~(
        et_nonull.str.contains("no")))
    et_nonull = et_nonull.where(~nans, np.nan)
    et_nonull = et_nonull.where(~si, 'si')

    df.loc[ind, 'Ethos Previos'] = et_nonull

    # %%% 16 - Periodo/Ciclo Recepción
    # Valores nulos
    df['Periodo/Ciclo Recepción'].replace(
        {'4': np.nan, 'no_identificado': np.nan}, inplace=True)

    # Se corrige el formato del periodo profesional
    # Las keys representan el pedazo a modificar y el valor el reemplazo
    val = {'febrero-junio': 'fj',
           'agosto-diciembre': 'ad',
           '-': '',
           '_': '',
           '2020': '20',
           'ptm': '',
           'verano': 'v',
           'invierno': 'i',
           'semestre': '',
           '202': '2',
           '201': '1'}

    val2 = {'febrero-junio': 'fj',
            'agosto-diciembre': 'ad',
            '1fj': 'fj'}

    # Primera ronda para formato
    df['Periodo/Ciclo Recepción'] = df['Periodo/Ciclo Recepción'].apply(
        lambda x: modificar_palabra(x, val))
    # Se hace nuevamente para valores nuevos obtenidos de aplicar el formato anterior
    df['Periodo/Ciclo Recepción'] = df['Periodo/Ciclo Recepción'].apply(
        lambda x: modificar_palabra(x, val2))

    def periodo(i):
        # De la base de denuncias, se obtiene el mes y año de recepción del registro i
        mes = den[den.Folio == i].Mes.values[0]
        year = str(den[den.Folio == i].Year.values[0] - 2000)
        # Nivel Preparatoria
        if den[den.Folio == i].Nivel.values[0] == 'Preparatoria':
            # Meses de cada periodo
            if mes >= 1 and mes <= 5:  # Enero Mayo
                a = 'EM' + year
            elif mes == 6:  # Verano
                a = 'V' + year + 'P'
            else:  # Agosto Diciembre
                a = 'AD' + year

        # Nivel profesional, migración o Tec de Monterrey
        else:
            # Meses de cada periodo
            if mes >= 2 and mes <= 6:  # Febrero Junio
                a = 'FJ' + year
            elif mes == 7:  # Verano
                a = 'V' + year
            elif mes == 1:  # Invierno
                a = 'I' + year
            else:  # Agosto Diciembre
                a = 'AD' + year
        return a

    for cadena in df['Periodo/Ciclo Recepción'].unique():
        # Si no tiene formato de abreviatura (AD, FJ, V o I) y es una string
        if (type(cadena) is str) and not (('AD' in cadena.upper()) or ('FJ' in cadena.upper()) or ('V20' in cadena.upper()) or ('I20' in cadena.upper())):

            # Folios de los registros que pasaron la condición
            ind = df[df['Periodo/Ciclo Recepción'] == cadena].index
            for i in ind:
                df.loc[i, 'Periodo/Ciclo Recepción'] = periodo(i)

    # Formato de mayúscula a la columna
    df['Periodo/Ciclo Recepción'] = df['Periodo/Ciclo Recepción'].str.upper()

    # Fecha a valores nulos
    ind = df[df['Periodo/Ciclo Recepción'].isnull()].index
    for i in ind:
        df.loc[i, 'Periodo/Ciclo Recepción'] = periodo(i)
    # %%% 17 - Region
    # Diccionario key = campus, value = región correspondiente
    regiones_dict = regiones.set_index('Campus')['Región'].to_dict()

    # Diccionario key = folio de denuncia, value = campus
    den_dict = den.set_index('Folio')['Localización'].to_dict()

    # Mapeo donde por el folio de denuncia se obtiene el campus y por este se obtiene la región
    df['Región'] = df.index.map(den_dict).map(regiones_dict)
    df['Región'].replace('No valida', 'No válida', inplace=True)

    # %%% 18 - Sexo De La Persona Reportada
    # Valores inválidos
    df['Sexo De La Persona Reportada'].replace(
        'no_identificado', np.nan, inplace=True)

    # Unificar los términos por 'M' o 'H'
    val = {
        'mujeres': 'M',
        'hombres': 'H',
        'mujer': 'M',
        'hombre': 'H',
        'femenino': 'M',
        'masculino': 'H',
        'y': '',
        '_': ''
    }

    # Se cambian los valores del dict val para cuestiones de formato
    df['Sexo De La Persona Reportada'] = df['Sexo De La Persona Reportada'].apply(
        lambda x: modificar_palabra(x, val))

    # Darle formato de xH-zM, donde x y z son números
    sexo = ['M', 'H']

    def formato_sexo(s):
        # Parte que en el caso de ser una string, divide los términos de H o M por un -
        if isinstance(s, str):
            for combo in sexo:
                s = re.sub(fr'(?<={combo})(?=\d+)', '-', s)
        # Parte que en caso de ser una string y con el formato de H y M, ordena para que sea primero M
        if isinstance(s, str) and ('-' in s):
            # Divide por '-'
            terminos = s.split('-')
            # M primero que H
            if 'H' in terminos[1]:
                terminos[0], terminos[1] = terminos[1], terminos[0]
            # Unir los términos con '-'
            nuevo_elemento = '-'.join(terminos)
            return nuevo_elemento
        else:
            return s

    df['Sexo De La Persona Reportada'] = df['Sexo De La Persona Reportada'].apply(
        formato_sexo)

    # %%% Incongruencias con cantidad de personas y cantidad sexo
    # Índice de los registros que tienen 'H' o 'M' y solo hay una persona implicada
    ind_h1 = df[(df['Sexo De La Persona Reportada'] == 'H') & (
        df['E Cantidad De Personas Implicadas En La Falta'] == 1)].index
    ind_m1 = df[(df['Sexo De La Persona Reportada'] == 'M') & (
        df['E Cantidad De Personas Implicadas En La Falta'] == 1)].index

    # Se reemplaza por '1H' o '1M'
    df.loc[ind_h1, 'Sexo De La Persona Reportada'] = '1H'
    df.loc[ind_m1, 'Sexo De La Persona Reportada'] = '1M'

    index_M_H = df[(df['Sexo De La Persona Reportada'] == 'M') | (
        df['Sexo De La Persona Reportada'] == 'H')].index
    df.loc[index_M_H, 'Sexo De La Persona Reportada'] = df.loc[index_M_H,
                                                               'E Cantidad De Personas Implicadas En La Falta'].astype(str) + df.loc[index_M_H, 'Sexo De La Persona Reportada']

    # %%% 19 - Tipo De Cierre/Resolución
    # Diccionario de Descripción con Categorización donde no se tiene alguna observación (condiciones)
    cierre_dict = cierre[cierre['Observación'].isnull()].set_index('Descripción')[
        'Categorización'].to_dict()
    df['Tipo De Cierre/Resolución'].replace(cierre_dict, inplace=True)

    # Df que tiene las sanciones aplicables en caso de una persona o 2+, además de la descripción
    sanciones_multiples = cierre[cierre['Observación'].notnull(
    )]['Categorización'].str.split(pat=',', expand=True)
    sanciones_multiples.columns = ['Individual', 'Grupal']
    sanciones_multiples['Descripción'] = cierre[cierre['Observación'].notnull(
    )]['Descripción']

    # Diccionarios
    sanciones_multiples.to_dict()
    san_ind = sanciones_multiples.set_index(
        'Descripción')['Individual'].to_dict()
    san_mult = sanciones_multiples.set_index('Descripción')['Grupal'].to_dict()

    # Cantidad de personas nulo
    index_sancion_null = df[(df['E Cantidad De Personas Implicadas En La Falta'].isnull()) & (
        df['Tipo De Cierre/Resolución'].isin(sanciones_multiples['Descripción'].values))].index
    df.loc[index_sancion_null, 'Tipo De Cierre/Resolución'] = df.loc[index_sancion_null,
                                                                     'Tipo De Cierre/Resolución'].replace(san_ind)

    # Cantidad de personas 1 y es el caso de condición sanción múltiple
    index_sancion_ind = df[(df['E Cantidad De Personas Implicadas En La Falta'] == 1) & (
        df['Tipo De Cierre/Resolución'].isin(sanciones_multiples['Descripción'].values))].index
    df.loc[index_sancion_ind, 'Tipo De Cierre/Resolución'] = df.loc[index_sancion_ind,
                                                                    'Tipo De Cierre/Resolución'].replace(san_ind)

    # Cantidad de personas >=2 y es el caso de condición sanción múltiple
    index_sancion_mult = df[(df['E Cantidad De Personas Implicadas En La Falta'] >= 2) & (
        df['Tipo De Cierre/Resolución'].isin(sanciones_multiples['Descripción'].values))].index
    df.loc[index_sancion_mult, 'Tipo De Cierre/Resolución'] = df.loc[index_sancion_mult,
                                                                     'Tipo De Cierre/Resolución'].replace(san_mult)

    # %%% Género
    def contar_sexo(s):
        if isinstance(s, str) and ('-' in s):
            # Dividir la cadena en 'H' y 'M' y convertir a números
            partes = s.split('-')
            return int(partes[0].replace('H', '')) + int(partes[1].replace('M', ''))
        elif isinstance(s, str) and s[0].isdigit():
            return int(s[0])  # caso de 1M o 1H
        elif isinstance(s, str):
            return 0  # caso de M o H
        else:
            return s  # nulo

    # Mayor número de matrículas que conteo sexo
    mayor_conteo = df[df['E Matrícula Del Estudiante Reportado'].str.count(
        'a').astype('Int64') > df['Sexo De La Persona Reportada'].apply(contar_sexo)]
    # Matrículas
    mat_sexo = mayor_conteo['E Matrícula Del Estudiante Reportado']

    # Conteo sexo
    conteo_sexo = pd.DataFrame(
        mayor_conteo['Sexo De La Persona Reportada'].apply(contar_sexo))
    conteo_sexo.columns = ['Conteo_original']

    # Separa matrículas, obtención de sexo en la base de datos insc_al y conteo
    def separar_matriculas(s):
        su = s.upper()
        indice = mat_sexo[mat_sexo == s].index.values[0]
        # Más de una matrícula
        if '-' in s:
            partes = su.split('-')  # matrículas
            # h, m = 0,0
            h = list(insc_al.loc[list(
                set(partes) - (set(partes) - set(insc_al.index))), 'Genero']).count('Masculino')
            m = list(insc_al.loc[list(
                set(partes) - (set(partes) - set(insc_al.index))), 'Genero']).count('Femenino')
            # Se considera el sexo original
            # Indice del registro s
            if h == 0:
                string = str(m)+'M'  # si solo es M
            elif m == 0:
                string = str(h)+'H'  # si solo es N
            else:
                string = str(h)+'H-'+str(m)+'M'

            # conteo original de sexo
            conteo_original = conteo_sexo.loc[indice].values[0]
            # si es entero (para evitar nan)
            if conteo_original.dtype == 'int64':
                if h+m > conteo_original:  # si el conteo nuevo es mayor al original, se regresa el nuevo formato
                    if h + m == 0:
                        # conteo original
                        return mayor_conteo.loc[indice, 'Sexo De La Persona Reportada']
                    else:
                        return string

                # Conteo nuevo menor al original (en caso que no se hayan encontrado matrículas)
                else:
                    # conteo original
                    return mayor_conteo.loc[indice, 'Sexo De La Persona Reportada']
            else:  # es nan el conteo original
                if h == 0 and m == 0:
                    return np.nan  # conteo original
                else:
                    return string

        else:  # una sola matrícula
            if su in insc_al.index:
                return '1'+insc_al.loc[su, 'Genero']  # solo una matrícula
            # si la matrícula no se tiene registrada
            else:
                return mayor_conteo.loc[indice, 'Sexo De La Persona Reportada']

    # Cambiar df
    df.loc[mayor_conteo.index, 'Sexo De La Persona Reportada'] = mat_sexo.apply(
        separar_matriculas)

    # %%% 14 - Nombre Del Reportado
    # Matrículas
    matriculas = df['E Matrícula Del Estudiante Reportado']

    def obtener_nombres(s):
        if isinstance(s, str):
            su = s.upper()
            indice = matriculas[matriculas == s].index.values[0]
            # Más de una matrícula
            if '-' in su:
                partes = su.split('-')  # matrículas
                # Nombres de las matrículas que están registradas con la base de datos de alumnos
                nombres = '_'.join(insc_al.loc[list(
                    set(partes) - (set(partes) - set(insc_al.index))), 'Nombre completo'].values)
                if nombres == '':
                    return np.nan
                else:
                    return nombres

            else:  # una sola matrícula
                if su in insc_al.index:
                    # solo una matrícula
                    return insc_al.loc[su, 'Nombre completo']
                # si la matrícula no se tiene registrada
                else:
                    return np.nan  # nombre original
        else:
            return np.nan
    df['Nombre del Reportado Nuevo'] = mat_sexo.apply(obtener_nombres)

    # %%% Formato final

    # Aplicar formato a nombres
    df['Nombre Del Reportado'] = df['Nombre Del Reportado'].apply(
        formato_strings_cap)
    df['Nombre Y Id Del Que Reporta'] = df['Nombre Y Id Del Que Reporta'].apply(
        formato_strings_cap)

    # Aplicar formato de separación de espacios
    df['E Lugar/Actividad Donde Ocurrió La Falta'] = df['E Lugar/Actividad Donde Ocurrió La Falta'].apply(
        formato_strings)
    df['E Materia/Unidad De Formación Donde Ocurrió La Falta'] = df['E Materia/Unidad De Formación Donde Ocurrió La Falta'].apply(
        formato_strings)

    # Formato de nan y tipo de datos
    df.replace({pd.NA: np.nan, 'NaN': np.nan,
               'NA': np.nan, 'nan': np.nan}, inplace=True)
    df['E Cantidad De Personas Implicadas En La Falta'] = df['E Cantidad De Personas Implicadas En La Falta'].astype(
        'Int64')
    df['E Semestre (#)'] = df['E Semestre (#)'].astype('Int64')

    # Mostrar el DataFrame modificado
    st.write("Nuevo archivo de Etiquetas:")
    st.write(df)

    # %%% Limpieza de Denuncia
    # Registros donde la unidad de negocio no es válida (prepa, profesional o posgrado) y no se sabe la sublocalización => nan
    den.loc[(~den['Nivel'].isin(['Preparatoria', 'Profesional', 'Posgrado'])) & (
        den['Sublocalización'].isna()), 'Nivel'] = np.nan

    # Registros donde la unidad de negocio no es válida (prepa, profesional o posgrado) y se sabe la sublocalización => profesional
    den.loc[(~den['Nivel'].isin(['Preparatoria', 'Profesional', 'Posgrado'])) & (
        ~den['Sublocalización'].isna()), 'Nivel'] = 'Profesional'
    den.rename(columns={'Mes': 'Mes de recepción', 'Year': 'Año de recepción',
               'Nivel': 'Unidad de Negocio'}, inplace=True)
    # Merge de Denuncias y Etiquetas
    tot = pd.merge(df.reset_index(), den, on='Folio')
    tot.set_index('Folio', inplace=True)
    den.set_index('Folio', inplace=True)

    # Asignar cantidad a las columnas de hombres y mujeres
    tot['Cantidad Mujeres'] = tot['Sexo De La Persona Reportada'].str.extract(
        r'(\d+)M')
    tot['Cantidad Hombres'] = tot['Sexo De La Persona Reportada'].str.extract(
        r'(\d+)H-?(\d*)M?')[0]

    # %%% Descargar los DataFrames modificados como archivo XLSX
    def df_to_excel(df, den, tot):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Etiquetas limpia', index=True)
            den.to_excel(writer, sheet_name='Denuncias limpia', index=True)
            tot.to_excel(writer, sheet_name='Etiquetas-Denuncias', index=True)
        output.seek(0)
        return output

    st.download_button(
        label="Etiquetas Modificado XLSX",
        data=df_to_excel(df, den, tot),
        file_name=f"{df_up.name.replace('.xlsx','')}_modificado.xlsx")

else:  # Si no se tiene los archivos suficientes
    st.subheader(":red[¡Alerta!]")
    st.write('_:red[Favor de subir los archivos faltantes]_')
