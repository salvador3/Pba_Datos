import pandas as pd
from spellchecker import SpellChecker
import matplotlib.pyplot as plt

def transformar_datos(dataframe):
    
    # Eliminar los valores nulos o faltantes
    dataframe = dataframe.dropna()

    # Eliminar registros duplicados
    dataframe_sin_duplicados = dataframe.drop_duplicates()

    # Corregir la ortografía de los datos
    spell = SpellChecker()
    dataframe_corregido = dataframe_sin_duplicados.copy()

    for column in dataframe_corregido.columns:
        try:
            dataframe_corregido[column] = dataframe[column].apply(lambda x: " ".join(spell.correction(word) for word in x.split()))
        except:
            pass

    dataframe=dataframe_corregido.copy()

    # Obtener los nombres de las columnas
    nombres_columnas = dataframe.columns.tolist()

    # Eliminar espacios adicionales en los nombres de las columnas
    nombres_columnas_actualizados = [nombre.replace("\n", "") for nombre in nombres_columnas]

    # Renombrar las columnas en el DataFrame
    dataframe.rename(columns=dict(zip(nombres_columnas, nombres_columnas_actualizados)), inplace=True)

    # keys = dataframe.keys()
    # # Imprimir cada key
    # for key in keys:
    #     print(key)


    # Reemplazar los valores de la columna "Caracter IES"
    dataframe["Caracter IES"] = dataframe["Caracter IES"].replace({
        "Institución Universitaria/Escuela Tecnológica": "Institución Tecnológica"
    })

    dataframe["Programa Académico"] = dataframe["Programa Académico"].replace({
        "ADMINISTRACION  FINANCIERA": "ADMINISTRACIÓN  FINANCIERA"
    })

    dataframe["Programa Académico"] = dataframe["Programa Académico"].replace({
        "DOCTORADO EN CIENCIAS - FISICA": "DOCTORADO EN CIENCIAS - FÍSICA"
    })

    dataframe["Programa Académico"] = dataframe["Programa Académico"].replace({
        "ECONOMIA": "ECONOMÍA"
    })

    dataframe["Programa Académico"] = dataframe["Programa Académico"].replace({
        "DISEÑO DE ESPACIOS - ESCENARIO": "DISEÑO DE ESPACIOS Y ESCENARIOS"
    })

    dataframe["Departamento de domicilio de la IES"] = dataframe["Departamento de domicilio de la IES"].replace({
        "BOGOTÁ D.C": "BOGOTÁ, D.C"
    })
    
    # Estandarizar el formato de los datos a letras capitales 
    dataframe = dataframe.applymap(lambda x: str(x).upper())
    
    return dataframe

def graficas(dataframe):
    # Contar el número de registros por institución
    conteo_instituciones = dataframe['Institución de Educación Superior (IES)'].value_counts()

    # Crear el gráfico de barras
    fig, ax = plt.subplots(figsize=(10, 6))
    conteo_instituciones.plot(kind='bar', ax=ax)
    ax.set_xlabel('Institución')
    ax.set_ylabel('Cantidad')
    ax.set_title('Cantidad de Registros por Institución')
    plt.xticks(rotation=90)

    # Guardar el gráfico como PNG
    ruta_grafico = 'barras.png'
    plt.savefig(ruta_grafico, bbox_inches='tight')

    # --------------------------------------------

    # Contar el número de registros por departamento
    conteo_departamentos = dataframe['Departamento de domicilio de la IES'].value_counts()

    # Crear el gráfico de pastel
    fig, ax = plt.subplots(figsize=(8, 8))
    conteo_departamentos.plot(kind='pie', ax=ax, autopct='%1.1f%%')
    ax.set_ylabel('')
    ax.set_title('Porcentaje de registros por departamento')

    # Guardar el gráfico como PNG
    ruta_grafico = 'pastel.png'
    plt.savefig(ruta_grafico, bbox_inches='tight')

    # --------------------------------------------

    # Crear el gráfico de puntos
    fig, ax = plt.subplots(figsize=(10, 6))
    dataframe.plot.scatter(y='Área de Conocimiento', x='Sexo', ax=ax)
    ax.set_ylabel('Área de Conocimiento')
    ax.set_xlabel('Sexo')
    ax.set_title('Relación entre Área de Conocimiento y Sexo')
        
    # Guardar el gráfico como PNG
    ruta_grafico = 'puntos.png'
    plt.savefig(ruta_grafico, bbox_inches='tight')

    # --------------------------------------------
    # Seleccionar la columna para el histograma
    columna = 'ID Nivel de Formación'

    # Convertir los valores a numéricos
    dataframe[columna] = pd.to_numeric(dataframe[columna], errors='coerce')

    # Eliminar las filas con valores no numéricos (opcional)
    dataframe = dataframe.dropna(subset=[columna])

    # Crear el histograma
    fig, ax = plt.subplots(figsize=(10, 6))
    dataframe[columna].plot.hist(ax=ax, bins=10)

    ax.set_xlabel(columna)
    ax.set_ylabel('Frecuencia')
    ax.set_title('Histograma de {}'.format(columna))

    # Guardar el histograma en formato PNG
    ruta_grafico = 'histograma.png'
    plt.savefig(ruta_grafico, bbox_inches='tight')

def main():
    # Cargar el archivo XLSX en un DataFrame
    ruta_archivo = 'articles-391559_recurso.xlsx'
    # dataframe = pd.read_excel(ruta_archivo, nrows=400)
    dataframe = pd.read_excel(ruta_archivo)

    # Aplicar la transformación
    dataframe_transformado = transformar_datos(dataframe)
    
    # Guardar el DataFrame corregido en un archivo CSV
    ruta_salida = 'salida.csv'
    dataframe_transformado.to_csv(ruta_salida, index=False)
    
    print("Transformación completada. Los datos transformados se han guardado en", ruta_salida)

    graficas(dataframe_transformado)

    print("Grafica completada. Los datos graficados se han guardado en ", "histograma.png, grafico_barras.png", ', pastel.png y puntos.png')
if __name__ == "__main__":
    main()