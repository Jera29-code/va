import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
import requests
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

# URL de descarga directa de Dropbox
dropbox_url = "https://www.dropbox.com/scl/fi/rh0zps7wltxyz8od323ej/Archivo_SIILNEVA_Recopiladas.xlsx?rlkey=u9eg3jrvbsghx8grczu4k7f9v&st=fe8fkkyl&dl=1"

# Función para descargar el archivo desde Dropbox y cargarlo en memoria
def download_excel_from_dropbox():
    try:
        # Realizar la solicitud para obtener el archivo
        response = requests.get(dropbox_url)
        
        if response.status_code == 200:
            # Cargar el archivo Excel desde los datos en memoria
            file_stream = BytesIO(response.content)
            data = pd.read_excel(file_stream)
            print("Archivo descargado y cargado exitosamente desde Dropbox.")
            return data
        else:
            print(f"Error al descargar el archivo: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error al descargar el archivo: {e}")
        return None

# Cargar datos desde Dropbox
data = download_excel_from_dropbox()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    folio = request.form['folio']
    if folio in data['Folio SIILNEVA'].values:
        row = data.loc[data['Folio SIILNEVA'] == folio].iloc[0]
        result = {
            'Id_Entidad': row['Id_Entidad'],
            'Entidad': row['Entidad'],
            'Id_Distrito Electoral Federal': row['Id_Distrito Electoral Federal'],
            'Cabecera D.E.F': row['Cabecera D.E.F'],
            'Número de Envio': row['Número de Envio']
        }
        return render_template('result.html', result=result, folio=folio)
    else:
        return render_template('index.html', error="Folio no encontrado.")

@app.route('/update', methods=['POST'])
def update():
    folio = request.form['folio']
    entregada = request.form['entregada']
    participacion = request.form['participacion']
    causal = request.form['causal']

    try:
        # Actualizar los datos en el DataFrame
        idx = data.loc[data['Folio SIILNEVA'] == folio].index[0]
        data.at[idx, 'Entregada (Si/No)'] = entregada
        data.at[idx, 'Manifestó su intención de participar (Si/No)'] = participacion
        data.at[idx, 'Causal'] = causal
    except Exception as e:
        print(f"Error al actualizar los datos: {e}")
        return render_template('index.html', error="Error al actualizar los datos.")
    
    return redirect(url_for('save'))

@app.route('/save')
def save():
    try:
        # Crear nombre de archivo con la fecha actual
        current_date = datetime.now().strftime("%d-%m-%y")  # Formato 29-01-25
        file_name = f"seguimiento_invitaciones_{current_date}.xlsx"

        # Generar la ruta temporal para guardar el archivo Excel en Dropbox
        # Aquí podrías devolver un enlace directo para descargarlo desde Dropbox
        # Si no es necesario almacenarlo localmente, puedes redirigir al usuario a una página con un enlace de descarga
        file_path = f"https://www.dropbox.com/scl/fi/rh0zps7wltxyz8od323ej/{file_name}?dl=1"

        # Guardar el archivo en Dropbox o en algún lugar temporal (como una carpeta pública)
        data.to_excel(file_name, index=False)
        
        # Subir a Dropbox sería un proceso adicional que no está incluido aquí.
        print(f"Archivo guardado como {file_name}. Puedes subirlo a Dropbox o mostrar el enlace de descarga.")

        return render_template('save.html', file_path=file_path)

    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
        return render_template('index.html', error=f"Error al guardar el archivo: {e}")

if __name__ == '__main__':
    # Descargar el archivo desde Dropbox
    data = download_excel_from_dropbox()
    app.run(debug=True)
