import streamlit as st
import pandas as pd
import io
import numpy as np

# Título de la aplicación
st.title("Búsqueda de Matrículas en Reportes de IA")

# Sección para cargar el archivo
st.subheader("Matrículas", divider='blue')
a = st.file_uploader("Archivo xlsx de _'Matrículas'_ a buscar", type=["xlsx"])

st.subheader("Etiquetas y Denuncias", divider='blue')
df_up = st.file_uploader(
    "Archivo xlsx de la hoja combinada de _'Etiqueta y Denuncias limpio'_ del reporte de IA", type=["xlsx"])

st.subheader("Base de datos escolar", divider='blue')
df_up2 = st.file_uploader(
    "Archivo xlsx de _'Base de datos escolar'_", type=["xlsx"])

# Si hay archivos
if (df_up is not None) and (df_up2 is not None) and (a is not None):
    st.header("Filtrado", divider='orange')
    # Leer archivos cargados
    df = pd.read_excel(df_up)
    df_a = pd.read_excel(a)
    df_esc = pd.read_excel(df_up2)

    # Arreglo de matrículas y convertirlo en una string
    mat = ','.join(map(str, df_a['Matrículas'].to_numpy()))
    mat = mat.replace(' ', '').lower()  # formato
    mats = '|'.join(mat.split(','))

    # Filtrado de instancias que contienen la matrícula
    df_mats = df[df['E Matrícula Del Estudiante Reportado'].str.contains(
        mats, na=False)]

    # Columnas a agregar
    df_mats['E Cantidad De Personas Implicadas En La Falta'] = df_mats['E Cantidad De Personas Implicadas En La Falta'].astype(
        'Int32')

    # Cada caso de aparición de matriculas
    mat_conteo = mat.split(',')
    conteo, personas, folios, cierre, region = [], [
    ], [], [], []  # de la base de datos modificada
    name, nivel = [], []  # de la base de datos de escolar
    for i in mat_conteo:
        # ----
        # Conteo de la matrícula en folios
        conteo.append(
            df_mats['E Matrícula Del Estudiante Reportado'].str.count(i).sum())

        # Subset de los folios que contienen la matrícula i
        x = df_mats[df_mats['E Matrícula Del Estudiante Reportado'].str.contains(
            i, na=False)]
        personas.append(','.join(x['E Cantidad De Personas Implicadas En La Falta'].astype(
            str)))  # números de personas implicadas
        # región(es) de la matrícula
        region.append(','.join(x['Región'].astype(str)))
        # folios donde está la matrícula
        folios.append(','.join(x['Folio'].astype(str)))
        # tipos de cierres
        cierre.append(','.join(x['Tipo De Cierre/Resolución'].astype(str)))

        # ----
        # Registro DB-Escolar de la matrícula i
        esc = df_esc[df_esc['Matrícula'] == i]
        unidad = df_mats[df_mats['E Matrícula Del Estudiante Reportado'] == i]
        if esc['Nombre completo'].empty:  # Si no existe el registro
            name.append(np.nan)
        else:
            name.append(esc['Nombre completo'].values[0])  # nombre

        if unidad['Unidad de Negocio'].empty:  # Si no existe el registro
            nivel.append(np.nan)
        else:
            # nivel de estudios
            nivel.append(unidad['Unidad de Negocio'].values[0])

    df_a['Nombre completo'] = name
    df_a['Unidad de Negocio'] = nivel
    df_a['Región'] = region
    df_a['Folios relacionados'] = folios
    df_a['Cantidad de folios relacionados'] = conteo
    df_a['Cantidad Personas'] = personas
    df_a['Tipo De Cierre/Resolución'] = cierre

    # Matrículas no encontradas
    no_mat = mats.split('|')
    no_mat = [elemento for elemento in no_mat if not df_mats['E Matrícula Del Estudiante Reportado'].str.contains(
        elemento, regex=True).any()]

    # Función para escribir el excel
    def df_to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(
                writer, sheet_name='Información matrículas', index=False)
        output.seek(0)
        return output

    # Si no se encontró ninguna matrícula
    if df_mats.empty:
        st.write('***:red[No se encontraron registros con las matrículas]***')

    # Si todas las matrículas se encontraron
    elif not no_mat:
        st.subheader('Archivo filtrado', divider='grey')
        df_mats
        st.download_button(
            label="Descargar XLSX filtrado",
            data=df_to_excel(df_mats),
            file_name=f"{df_up.name.replace('.xlsx','')}_filtrado.xlsx")

        st.subheader('Resumen matrículas', divider='grey')
        df_a
        st.download_button(
            label="Descargar XLSX de conteo",
            data=df_to_excel(df_a),
            file_name=f"{a.name.replace('.xlsx','')}_conteo.xlsx")

    # Si hubo matrículas tanto no encontradas como encontradas
    else:
        st.subheader('Matrículas no encontradas', divider='grey')
        st.write(','.join(no_mat))

        st.subheader('Archivo filtrado', divider='grey')
        df_mats

        st.download_button(
            label="Descargar XLSX filtrado",
            data=df_to_excel(df_mats),
            file_name=f"{df_up.name.replace('.xlsx','')}_filtrado.xlsx")

        st.subheader('Resumen matrículas', divider='grey')
        df_a
        st.download_button(
            label="Descargar XLSX de conteo",
            data=df_to_excel(df_a),
            file_name=f"{a.name.replace('.xlsx','')}_conteo.xlsx")


else:  # Si no se tiene los archivos suficientes
    st.subheader(":red[¡Alerta!]")
    st.write('_:red[Favor de subir los archivos faltantes]_')
