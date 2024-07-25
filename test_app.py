import streamlit as st
import pandas as pd
from io import BytesIO
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

st.title('Drag and Drop File Upload')

# Crear la casilla de arrastrar y soltar
uploaded_file = st.file_uploader("Arrastra y suelta un archivo Excel aquí", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leer el archivo Excel
    df = pd.read_excel(uploaded_file) 
    st.subheader("Datos Originales")
    st.write(df)


    max_row_to_delete= list()
    for i, fila in enumerate(df.index):
        lista_valores_fila = list()
        for j, columna in enumerate(df.columns):
            contenido_celda = df.loc[fila, columna]
            lista_valores_fila.append(contenido_celda)
        if  all(pd.isna(elemento) for elemento in lista_valores_fila):
            max_row_to_delete.append(fila)
    df.drop(max_row_to_delete,inplace=True)

    for columna in df.columns:
        # Verifica si todos los valores en la columna son nulos
        if df[columna].isnull().all():
            # Elimina la columna si todos los valores son nulos
            df.drop(columna, axis=1, inplace=True)
    df.reset_index(drop=True,inplace=True)  # resetear índice de filas
    df.columns = range(df.shape[1])  # resetear índice de columnas
    titular = df.iloc[0,1]
    rut = df.iloc[1,1]
    cuenta = df.iloc[2,1]
    fecha_emision = df.iloc[4,3]
    variables_cartola = (titular,rut,cuenta,fecha_emision)

    titular,rut,cuenta,fecha_emision = variables_cartola
    headers_values = []
    expected_headers = ['Fecha', 'Descripción', 'Canal o Sucursal', 'Cargos (PESOS)', 'Abonos (PESOS)', 'Saldo (PESOS)'] ### cuales son los headers a dejar 
    deleted_rows_not_table = 0 ## almacena hasta donde se va a eliminar datos que no son propios de la tabla de datos
    for indice, fila in df.iterrows():
        valores_fila = fila.tolist()
        # Comparación insensible a mayúsculas y minúsculas y sin espacios adicionales
        valores_fila_strip = [valor.strip().lower() if isinstance(valor, str) else valor for valor in valores_fila]
        if valores_fila_strip == [header.lower() for header in expected_headers]:
            headers_values.append(valores_fila)
            deleted_rows_not_table = indice
            df.drop(list(range(deleted_rows_not_table)),inplace=True) ## elimina hasta el valor de indice que cohincide con el header esperado
            print(f'Eliminar filas hasta la fila: {deleted_rows_not_table}')
    nuevos_encabezados = df.iloc[0] 
    df.rename(columns=nuevos_encabezados,inplace = True)  # Establecer los nuevos encabezados de columna
    df.drop(df.index[0], inplace=True)
    df.reset_index(drop=True,inplace=True)
    df.drop(df.index[-1], inplace=True)  ## eliminar la ultima fila innecesaria de datos
    df['Cliente']= titular
    df['Rut'] = rut
    df['cuenta'] = cuenta
    df['fecha_emision'] = fecha_emision
    df['fecha_emision'] = pd.to_datetime(df['fecha_emision'])
    df['Año'] = df['fecha_emision'].dt.year #####formateo de
    anno = df['Año'].iloc[0]
    df['Fecha'] = df['Fecha'] + f'/{str(anno)}'
    df['Fecha'] = pd.to_datetime(df['Fecha'])
    df.drop(df.index[0], inplace=True)
    df.drop(df.index[-1], inplace=True)

    # Mostrar el DataFrame
    st.write("Aquí está el DataFrame:")
    st.dataframe(df)

    def to_excel_file(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    df_xlsx = to_excel_file(df)

    st.download_button(
        label="Descargar datos limpios en Excel",
        data=df_xlsx,
        file_name='datos_limpios.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.write("Por favor, sube un archivo Excel para continuar.")