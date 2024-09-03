import os
from flask import Flask, render_template, request, redirect, flash, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'super_secret_key'

# Variable global para almacenar los datos de la tabla
table_data_EmisorReceptorTurisExcur = None


def create_upload_folder():
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_excel(file_path):
    global table_data_EmisorReceptorTurisExcur

    try:

        df = pd.read_excel(file_path)
             
    ############################################################################### Graf15 ###############################################################################
        filtro_operMar15 = df['Operación'].str.contains('Maritima', na=False)
        filtro_visitante_exc15 = df['Visitante'].str.contains('Excur', na=False)

        Gasto_tot_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]

        Gasto_paq_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_alo_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_tra_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_ali_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_cul_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_art_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_bie_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]

        Gasto_joy_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_dep_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_mer_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_med_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_otr_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]
        Gasto_mas_graf15_filtrado = df[filtro_operMar15 & filtro_visitante_exc15]

        multiplicacion_fexp_Gasto_tot_graf15_filtrado = (
        Gasto_tot_graf15_filtrado['Gasto_tot'] * Gasto_tot_graf15_filtrado['fexp']).sum()

        multiplicacion_fexp_Gasto_paq_graf15_filtrado = (
        Gasto_paq_graf15_filtrado['Gasto_paq'] * Gasto_paq_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_alo_graf15_filtrado = (
        Gasto_alo_graf15_filtrado['Gasto_alo'] * Gasto_alo_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_tra_graf15_filtrado = (
        Gasto_tra_graf15_filtrado['Gasto_tra'] * Gasto_tra_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_ali_graf15_filtrado = (
        Gasto_ali_graf15_filtrado['Gasto_ali'] * Gasto_ali_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_cul_graf15_filtrado = (
        Gasto_cul_graf15_filtrado['Gasto_cul'] * Gasto_cul_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_art_graf15_filtrado = (
        Gasto_art_graf15_filtrado['Gasto_art'] * Gasto_art_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_bie_graf15_filtrado = (
        Gasto_bie_graf15_filtrado['Gasto_bie'] * Gasto_bie_graf15_filtrado['fexp']).sum()

        multiplicacion_fexp_Gasto_joy_graf15_filtrado = (
        Gasto_joy_graf15_filtrado['Gasto_joy'] * Gasto_joy_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_dep_graf15_filtrado = (
        Gasto_dep_graf15_filtrado['Gasto_dep'] * Gasto_dep_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_mer_graf15_filtrado = (
        Gasto_mer_graf15_filtrado['Gasto_mer'] * Gasto_mer_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_med_graf15_filtrado = (
        Gasto_med_graf15_filtrado['Gasto_med'] * Gasto_med_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_mas_graf15_filtrado = (
        Gasto_mas_graf15_filtrado['Gasto_mas'] * Gasto_mas_graf15_filtrado['fexp']).sum()
        multiplicacion_fexp_Gasto_otr_graf15_filtrado = (
        Gasto_otr_graf15_filtrado['Gasto_otr'] * Gasto_otr_graf15_filtrado['fexp']).sum()

        Gasto_paq_graf15 = round(
    (multiplicacion_fexp_Gasto_paq_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_alo_graf15 = round(
    (multiplicacion_fexp_Gasto_alo_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_tra_graf15 = round(
    (multiplicacion_fexp_Gasto_tra_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_ali_graf15 = round(
    (multiplicacion_fexp_Gasto_ali_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_cul_graf15 = round(
    (multiplicacion_fexp_Gasto_cul_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_art_graf15 = round(
    (multiplicacion_fexp_Gasto_art_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_bie_graf15 = round(
    (multiplicacion_fexp_Gasto_bie_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_joy_graf15 = round(
    (multiplicacion_fexp_Gasto_joy_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_dep_graf15 = round(
    (multiplicacion_fexp_Gasto_dep_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_mer_graf15 = round(
    (multiplicacion_fexp_Gasto_mer_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_med_graf15 = round(
    (multiplicacion_fexp_Gasto_med_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)
        Gasto_mas_graf15 = round(
    (multiplicacion_fexp_Gasto_mas_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        Gasto_otr_graf15 = round(
    (multiplicacion_fexp_Gasto_otr_graf15_filtrado / multiplicacion_fexp_Gasto_tot_graf15_filtrado) * 100, 1)

        otros_juntos15 = sum([Gasto_dep_graf15, Gasto_med_graf15, Gasto_mer_graf15, Gasto_otr_graf15, Gasto_alo_graf15])

        Gasto_tra_graf15_formateada = "{:,.1f}".format(Gasto_tra_graf15)
        Gasto_ali_graf15_formateada = "{:,.1f}".format(Gasto_ali_graf15)
        Gasto_cul_graf15_formateada = "{:,.1f}".format(Gasto_cul_graf15)
        Gasto_art_graf15_formateada = "{:,.1f}".format(Gasto_art_graf15)
        Gasto_Joy_graf15_formateada = "{:,.1f}".format(Gasto_joy_graf15)
        Gasto_bie_graf15_formateada = "{:,.1f}".format(Gasto_bie_graf15)
        otros_juntos_graf15_formateada = "{:,.1f}".format(otros_juntos15)
        Gasto_Paquete_graf15_formateada = "{:,.1f}".format(Gasto_paq_graf15)

        table_data_Gráfico_15 = [
    ["Servicios culturales y recreativos:", f"{Gasto_cul_graf15_formateada} %"],
    ["Compra de artesanías, suvenires y/o regalos:", f"{Gasto_art_graf15_formateada} %"],
    ["Compra de joyas y accesorios de lujo:", f"{Gasto_Joy_graf15_formateada} %"],
    ["Alimentos y bebidas:", f"{Gasto_ali_graf15_formateada} %"],
    ["Otros*:", f"{otros_juntos_graf15_formateada} %"],
    ["Transporte terrestre, acuático y aéreo:", f"{Gasto_tra_graf15_formateada} %"],
    ["Bienes de uso personal:", f"{Gasto_bie_graf15_formateada} %"],
    ["Paquete turístico:", f"{Gasto_Paquete_graf15_formateada} %"],
]

# Datos para el gráfico 15
        labels15 = ['Servicios culturales y recreativos:', 'Compra de artesanías, suvenires y/o regalos:',
            'Compra de joyas y accesorios de lujo:', 'Alimentos y bebidas:',
            'Otros*:', 'Transporte terrestre, acuático y aéreo:', 'Bienes de uso personal:',
            'Paquete turístico:']
        values15 = [Gasto_cul_graf15, Gasto_art_graf15, Gasto_joy_graf15, Gasto_ali_graf15, otros_juntos15,
            Gasto_tra_graf15, Gasto_bie_graf15, Gasto_paq_graf15]

# Palabras finales
        palabras_finales = ["Otros*:", "Otros:", "Promedio total*:", "Total:"]

# Función para verificar si una palabra está en palabras_finales
        def es_palabra_final(palabra):
            return palabra in palabras_finales

# Crear una lista de tuplas con etiquetas y valores
        data15 = list(zip(labels15, values15))

# Ordenar la lista de tuplas: primero por si es palabra final (para ponerlas al final), luego por valor (de mayor a menor)
        data15_sorted = sorted(data15, key=lambda x: (es_palabra_final(x[0]), -x[1]), reverse=True)

# Extraer las etiquetas ordenadas y los valores ordenados
        labels15_sorted = [item[0] for item in data15_sorted]
        values15_sorted = [item[1] for item in data15_sorted]

# Colores para las barras (el mismo color para todas las barras)
        colors15 = ['#8b0000'] * len(labels15_sorted)

# Crear la figura y los ejes para el gráfico
        fig, ax15 = plt.subplots(figsize=(8, 6))
        bars = ax15.barh(labels15_sorted, values15_sorted, color=colors15, height=0.3)

# Establecer etiquetas y título
# ax15.set_xlabel('N° de viajes internacionales', color='#A6A6A6')
# ax15.set_title('Participación del gasto por rubro', color='#A6A6A6')

# Ajustar el formato del eje X para evitar formato científico
        ax15.ticklabel_format(style='plain', axis='x')

# Agregar los valores encima de las barras (horizontalmente)
        for bar, value in zip(bars, values15_sorted):
            ax15.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2, f'{value:.1f} %', va='center',
              ha='left')

# Eliminar bordes innecesarios
        ax15.spines['top'].set_visible(False)
        ax15.spines['right'].set_visible(False)
        ax15.spines['bottom'].set_visible(False)
        ax15.spines['left'].set_visible(False)

# Mostrar los labels ordenados al lado de las barras
        ax15.set_yticklabels(labels15_sorted)

# Ajustar los márgenes del eje y para reducir la distancia entre las barras y el eje x
        ax15.set_ylim(ax15.get_ylim()[0] - 0.01, ax15.get_ylim()[1])

# Guardar la figura en un archivo temporal
        plt.tight_layout()
        plt.savefig('static/bar_chart_15.png')
        plt.close()

############################################################################### Graf16 ###############################################################################
        filtro_Operacion_graftab16 = df['Operación'].str.contains('Maritima', na=False)
        filtro_tconocioR16 = (df['Conocio'] == 'Recomendaciones')
        filtro_tconocioM16 = (df['Conocio'] == 'Medios digitales')
        filtro_tconocioME16 = (df['Conocio'] == 'Medios especializ.')
        filtro_tconocioT16 = (df['Conocio'] == 'Televisión')

        filtro_total_conocio = (df['Conocio'] == 'Recomendaciones') | (df['Conocio'] == 'Medios digitales') | (
        df['Conocio'] == 'Medios especializ.') | (df['Conocio'] == 'Televisión')

        Recomendaciones_graf16_filtrado = df[filtro_tconocioR16 & filtro_Operacion_graftab16]
        Medio_dig_graf16_filtrado = df[filtro_tconocioM16 & filtro_Operacion_graftab16]
        Medios_esp_graf16_filtrado = df[filtro_tconocioME16 & filtro_Operacion_graftab16]
        Televsion_graf16_filtrado = df[filtro_tconocioT16 & filtro_Operacion_graftab16]

        todosConocido_graf16_filtrado = df[filtro_total_conocio & filtro_Operacion_graftab16]

        suma_fexp_Recomendaciones_graf16 = round(Recomendaciones_graf16_filtrado['fexp'].sum())
        suma_fexp_medios_dig_graf16 = round(Medio_dig_graf16_filtrado['fexp'].sum())
        suma_fexp_medios_esp_graf16 = round(Medios_esp_graf16_filtrado['fexp'].sum())
        suma_fexp_Tv_graf16 = round(Televsion_graf16_filtrado['fexp'].sum())

        suma_fexp_todo_Cono = round(todosConocido_graf16_filtrado['fexp'].sum())

        Recomendaciones_graf16 = round((suma_fexp_Recomendaciones_graf16 / suma_fexp_todo_Cono) * 100, 1)
        Medios_dig_graf16 = round((suma_fexp_medios_dig_graf16 / suma_fexp_todo_Cono) * 100, 1)
        Medios_esp_graf16 = round((suma_fexp_medios_esp_graf16 / suma_fexp_todo_Cono) * 100, 1)
        Tv_graf16 = round((suma_fexp_Tv_graf16 / suma_fexp_todo_Cono) * 100, 1)

        Recomendaciones_graf16_formateada = "{:,.1f}".format(Recomendaciones_graf16)
        Rmedios_dig_graf16_formateada = "{:,.1f}".format(Medios_dig_graf16)
        medios_esp_graf16_formateada = "{:,.1f}".format(Medios_esp_graf16)
        Tv_graf16_formateada = "{:,.1f}".format(Tv_graf16)

        table_data_Gráfico_16 = [
    [" Recomendaciones de familiares y amigos", f"{Recomendaciones_graf16_formateada} %"],
    [" Medios digitales", f"{Rmedios_dig_graf16_formateada} %"],
    [" Televisión", f"{Tv_graf16_formateada} %"],
    [" Medios especializados", f"{medios_esp_graf16_formateada} %"],

]

        # Datos del gráfico 16
        labels16 = ["Recomendaciones de familiares y amigos", "Medios digitales", "Televisión", "Medios especializados"]
        values16 = [Recomendaciones_graf16, Medios_dig_graf16, Tv_graf16, Medios_esp_graf16]

# Colores para las secciones del gráfico
        colors16 = ['#ff99cc', '#575756', '#f29b00', '#8a4e8c']

# Crear una lista de tuplas con etiquetas, valores y colores
        data16 = list(zip(labels16, values16, colors16))

# Ordenar la lista de tuplas por los valores en orden descendente
        data16_sorted = sorted(data16, key=lambda x: x[1], reverse=True)

# Extraer las etiquetas ordenadas, los valores ordenados y los colores ordenados
        labels16_sorted = [item[0] for item in data16_sorted]
        values16_sorted = [item[1] for item in data16_sorted]
        colors16_sorted = [item[2] for item in data16_sorted]

# Función para formatear los porcentajes con coma
        def format_percent(val):
            return f'{val:.1f}'.replace('.', ',') + '%'

# Crear el gráfico de donut
        fig, ax16 = plt.subplots(figsize=(10, 8))  # Aumentar el tamaño de la figura para mayor legibilidad
        wedges, texts, autotexts = ax16.pie(values16_sorted, colors=colors16_sorted,
                                    autopct=lambda p: format_percent(p),
                                    pctdistance=0.85, startangle=0,
                                    wedgeprops={'linewidth': 2, 'edgecolor': 'white'})

# Dibujar un círculo blanco en el centro para el efecto de donut
        centre_circle = plt.Circle((0, 0), 0.60, fc='white')
        fig.gca().add_artist(centre_circle)

# Establecer el aspecto igual para asegurar que el gráfico de donut sea un círculo
        ax16.axis('equal')

# Ajustar los porcentajes para que no se superpongan
        for text in autotexts:
            text.set_fontsize(9)
            text.set_color('#000000')
            text.set_fontweight('bold')  # Establecer el texto en negrita
            text.set_bbox(dict(facecolor='#d9d9d9', edgecolor='none', pad=1))  # Agregar el recuadro gris claro

# Colocar la leyenda en el lado derecho con etiquetas ordenadas
        ax16.legend(wedges, labels16_sorted, title="Categorías", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

# Guardar el gráfico
        plt.tight_layout()
        plt.savefig('static/donut_chart_16.png')
        plt.close()
        
        ############################################################################### Graf-Tabla 4 ###############################################################################

        filtro_visitanteT_graftab4 = df['Visitante'].str.contains('Turis', na=False)
        filtro_visitanteE_graftab4 = df['Visitante'].str.contains('Excur', na=False)
        filtro_Operacion_graftab4 = df['Operación'].str.contains('Maritima', na=False)

        Excur_graftab4_filtrado = df[filtro_Operacion_graftab4 & filtro_visitanteE_graftab4]
        Turis_graftab4_filtrado = df[filtro_Operacion_graftab4 & filtro_visitanteT_graftab4]
        Maritm_graftab4_filtrado = df[filtro_Operacion_graftab4]

        suma_fexp_Turis_graftab4 = Turis_graftab4_filtrado['fexp'].sum()
        suma_fexp_Excur_graftab4 = Excur_graftab4_filtrado['fexp'].sum()
        suma_fexp_TurExc_graftab4 = Maritm_graftab4_filtrado['fexp'].sum()

        multiplicacion_fexp_gastoT_Turis_grafTab4 = Maritm_graftab4_filtrado['Gasto_tot'] * Maritm_graftab4_filtrado[
            'fexp']
        division_fexp_gastoT_Turis_grafTab4 = multiplicacion_fexp_gastoT_Turis_grafTab4.sum() / suma_fexp_TurExc_graftab4

        Turis_graftab4 = round(suma_fexp_Turis_graftab4)
        Excur_graftab4 = round(suma_fexp_Excur_graftab4)
        TurExc_graftab4 = round(suma_fexp_TurExc_graftab4)
        DidivTurisExc_graftab4 = round(division_fexp_gastoT_Turis_grafTab4, 1)

        Turis_grafTab4_formateada = "{:,.0f}".format(Turis_graftab4).replace(",", ".")
        Excur_grafTab4_formateada = "{:,.0f}".format(Excur_graftab4).replace(",", ".")
        TurExc_grafTab4_formateada = "{:,.0f}".format(TurExc_graftab4).replace(",", ".")
        DidivTurisExc_graftab4_formateada = "{:,.1f}".format(DidivTurisExc_graftab4).replace(".", ",")

        table_data_Tabla_4 = [
    (TurExc_graftab4, Turis_graftab4, Excur_graftab4, DidivTurisExc_graftab4,
     TurExc_grafTab4_formateada, Turis_grafTab4_formateada, Excur_grafTab4_formateada,
     DidivTurisExc_graftab4_formateada)
]

############################################################################### Graf-Tabla 5 ###############################################################################

        filtro_RecomiendaSi_graftab5 = df['Recomienda'].str.contains('Sí', na=False)
        filtro_RecomiendaNo_graftab5 = df['Recomienda'].str.contains('No', na=False)
        filtro_Operacion_graftab5 = df['Operación'].str.contains('Maritima', na=False)

        SiRecom_graftab5_filtrado = df[filtro_RecomiendaSi_graftab5 & filtro_Operacion_graftab5]
        NoRecom_graftab5_filtrado = df[filtro_RecomiendaNo_graftab5 & filtro_Operacion_graftab5]

        suma_fexp_SiRecom_graftab5 = SiRecom_graftab5_filtrado['fexp'].sum()
        suma_fexp_NoRecom_graftab5 = NoRecom_graftab5_filtrado['fexp'].sum()
        suma_fexp_SiNoRecom_graftab5 = SiRecom_graftab5_filtrado['fexp'].sum() + NoRecom_graftab5_filtrado['fexp'].sum()

        Valor_fexp_RecomdSi_graftab5 = round(suma_fexp_SiRecom_graftab5)
        Valor_fexp_RecomdNo_graftab5 = round(suma_fexp_NoRecom_graftab5)
        porcentaje_fexp_RecomdSi_graftab5 = round((suma_fexp_SiRecom_graftab5 / suma_fexp_SiNoRecom_graftab5) * 100, 1)
        porcentaje_fexp_RecomdNo_graftab5 = round((suma_fexp_NoRecom_graftab5 / suma_fexp_SiNoRecom_graftab5) * 100, 1)

        RecomdSi_graftab5_grafTab5_formateada = "{:,.0f}".format(Valor_fexp_RecomdSi_graftab5).replace(",", ".")
        porcentaje_RecomdSi_graftab5_formateada = "{:,.1f}".format(porcentaje_fexp_RecomdSi_graftab5).replace(".", ",")
        RecomdNo_graftab5_formateada = "{:,.0f}".format(Valor_fexp_RecomdNo_graftab5).replace(",", ".")
        porcentaje_RecomdNo_graftab5 = "{:,.1f}".format(porcentaje_fexp_RecomdNo_graftab5).replace(".", ",")

        table_data_Tabla_5 = [
            (Valor_fexp_RecomdSi_graftab5, porcentaje_fexp_RecomdSi_graftab5, Valor_fexp_RecomdNo_graftab5,
             porcentaje_fexp_RecomdNo_graftab5,
             RecomdSi_graftab5_grafTab5_formateada, porcentaje_RecomdSi_graftab5_formateada,
             RecomdNo_graftab5_formateada, porcentaje_RecomdNo_graftab5)
]
 

        # Retorna todas las tablas
        return table_data_Gráfico_15, table_data_Gráfico_16, table_data_Tabla_5, table_data_Tabla_4

    except Exception as e:
        flash(f"Ocurrió un error al procesar el archivo: {str(e)}", 'error')
        return None


@app.route('/', methods=['GET', 'POST'])
def upload_excel():
    create_upload_folder()

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Llama a la función process_excel para obtener datos de las tablas
            result = process_excel(file_path)

            # Verifica si se obtuvieron datos de las 16 tablas
            if result and len(result) == 4:
                # Retorna los datos obtenidos
                return render_template('dashboard.html', graphs=['static/bar_chart_15.png',
                                                                 'static/donut_chart_16.png'],
                                       table_data_Gráfico_15=result[0], table_data_Gráfico_16=result[1],  table_data_Tabla_5=result[2],
                                       table_data_Tabla_4=result[3])
            else:
                # Muestra un mensaje de error si no se obtuvieron datos de todas las tablas
                flash('Error al procesar el archivo. Algunas tablas no se generaron correctamente.', 'error')

    # Retorna la plantilla de carga de archivos si no se procesó correctamente el archivo
    return render_template('upload.html')


@app.route('/download/<filename>')
def download(filename):
    global table_data_Gráfico_15, table_data_Gráfico_16, table_data_Tabla_5, table_data_Tabla_4

    # Generar la imagen de la tabla
    fig, axs = plt.subplots(4)  # Dos subplots para dos tablas
    axs[0].axis('tight')
    axs[0].axis('off')
    axs[0].table(cellText=table_data_Gráfico_15, colLabels=['Otra etiqueta', 'Otro dato'], loc='center')

    axs[1].axis('tight')
    axs[1].axis('off')
    axs[1].table(cellText=table_data_Gráfico_16, colLabels=['Otra etiqueta', 'Otro dato'], loc='center')

    axs[2].axis('tight')
    axs[2].axis('off')
    axs[2].table(cellText=table_data_Tabla_5, colLabels=['Otra etiqueta', 'Otro dato'], loc='center')

    axs[3].axis('tight')
    axs[3].axis('off')
    axs[3].table(cellText=table_data_Tabla_4, colLabels=['Otra etiqueta', 'Otro dato'], loc='center')

    plt.subplots_adjust(left=0.2, right=0.8, top=0.8, bottom=0.2)  # Ajusta los márgenes

    plt.tight_layout()
    plt.savefig('table_image.png')  # Guarda la imagen como PNG
    plt.close()

    # Descargar la imagen generada
    return send_file('table_image.png', as_attachment=True)


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)
