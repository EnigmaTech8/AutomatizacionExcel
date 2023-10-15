import time
import pandas as pd

def main():

    start = time.time()

    # Lectura de datos del archivo csv.
    df = pd.read_csv("../Datasets/ofertas_de_empleo_sample.csv")

    # Tabla dinámica de total ofertas de empleo por compañía
    df_oe_company = df["company"].value_counts().rename_axis("COMPANY").reset_index(name="TOTAL EMPLEO")
    df_oe_company.sort_values("TOTAL EMPLEO", ascending=False, inplace=True)

    # Guardado de todos los datos en excel
    with pd.ExcelWriter("Reporte_Final.xlsx", engine="xlsxwriter") as writer:
        header_format = writer.book.add_format({"bg_color": "red"})
        df_to_table(df, writer, "Dataset", header_format, False)
        df_to_table(df_oe_company, writer, "OE_Company", header_format, False)

        final_plot(writer, writer.book.add_chart({"type": "pie"}))

    print("\nTiempo de ejecución: ", time.time() - start)

def df_to_table(df, writer, sheet, header_format, index):

    # Guardamos los datos en excel
    df.to_excel(writer, sheet_name=sheet, index=index)

    # Obtenemos la hoja
    worksheet = writer.sheets[sheet]

    # Creamos la lista con la cabecera de la tabla
    column_settings = [{'header': column_name, "header_format": header_format} for column_name in df.columns]

    # Obtenemos las dimensiones del dataframe o tabla
    (max_row, max_col) = df.shape

    # Añadimos la tabla con estilo
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings, "style": "Table style medium 2"})

    # Ampliamos un poco las columnas
    worksheet.set_column(0, max_col - 1, 15)

def final_plot(writer, chart):
    # Hoja en la que se insertará la gráfica
    sheet = "OE_Company"
    
    # Cambiamos el tamaño de la gráfica
    chart.set_size({"width": 550, "height": 300})

    # Añadimos el título a la gráfica
    chart.set_title({"name": "COMPAÑÍAS CON MÁS OFERTAS DE TRABAJO"})

    # Estilo de gráfica predefinido
    chart.set_style(42)

    # Añadimos una serie que contenga las diez primeras líneas (Top 10)
    chart.add_series(
        {
            "categories": f"={sheet}!$A$2:$A$11",
            "values": f"={sheet}!$B$2:$B11$",
            "data_labels":
            {
                "value": True,
                "category": True,
                "percentage": True,
                "leader_lines": True,
                "position": "outside_end",
                "font":
                {
                    "size": 10
                }
            }
        }
    )

    # Cambiamos el color de fondo de la gráfica
    chart.set_chartarea({'fill': {'color': '#5b5b5b'}})

    # Quitamos la leyenda de la gráfica
    chart.set_legend({"none": True})

    # Obtenemos la hoja e insertamos la gráfica
    worksheet = writer.sheets[sheet]
    worksheet.insert_chart("D2", chart)

if __name__ == '__main__':
    print("\nGenerando reporte...")
    main()
    print("\nReporte completado")