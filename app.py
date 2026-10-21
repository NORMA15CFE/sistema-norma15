from datetime import datetime
import os
import re
import zipfile
import io
import docx
from docx.text.paragraph import Paragraph
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell

import sqlite3
from xml.sax.saxutils import escape
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'clave_secreta_cfe_norma15'

UPLOAD_FOLDER = 'static/uploads_alumnos'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

USUARIOS_PERMITIDOS = {
    'DJ000': {'nombre': 'Oficinas Divisionales', 'rol': 'admin'},
    'DJ010': {'nombre': 'Poza Rica', 'rol': 'seccion'},
    'DJ020': {'nombre': 'Xalapa', 'rol': 'seccion'},
    'DJ030': {'nombre': 'Teziutlán', 'rol': 'seccion'},
    'DJ060': {'nombre': 'Veracruz', 'rol': 'seccion'},
    'DJO70': {'nombre': 'Papaloapan', 'rol': 'seccion'},
    'DJ100': {'nombre': 'Los Tuxtlas', 'rol': 'seccion'},
    'DJ110': {'nombre': 'Coatzacoalcos', 'rol': 'seccion'},
    'DJ130': {'nombre': 'Orizaba', 'rol': 'seccion'},
    'DJ140': {'nombre': 'Córdoba', 'rol': 'seccion'},
}

MESES_ESP = {
    'January': 'enero',
    'February': 'febrero',
    'March': 'marzo',
    'April': 'abril',
    'May': 'mayo',
    'June': 'junio',
    'July': 'julio',
    'August': 'agosto',
    'September': 'septiembre',
    'October': 'octubre',
    'November': 'noviembre',
    'December': 'diciembre',
}


def fecha_a_espanol(dt_obj):
    dia = dt_obj.strftime('%d')
    mes_en = dt_obj.strftime('%B')
    anio = dt_obj.strftime('%Y')
    mes_es = MESES_ESP.get(mes_en, mes_en)
    return f'{dia} de {mes_es} de {anio}'


def init_db():
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            co INTEGER PRIMARY KEY AUTOINCREMENT,
            seccion TEXT NOT NULL,
            nombre_alumno TEXT NOT NULL,
            tipo_apoyo TEXT NOT NULL,
            area TEXT NOT NULL,
            institucion TEXT NOT NULL,
            carrera TEXT NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_fin TEXT NOT NULL,
            estatus TEXT NOT NULL,
            horas INTEGER NOT NULL,
            es_familiar TEXT NOT NULL,
            nombre_familiar TEXT,
            parentesco_familiar TEXT,
            tipo_remuneracion TEXT NOT NULL,
            banco_clabe TEXT,
            numero_clabe TEXT,
            archivo_carta_presentacion TEXT,
            archivo_carta_aceptacion TEXT,
            archivo_constancia TEXT,
            archivo_buena_conducta TEXT,
            archivo_ine TEXT,
            archivo_reporte_mensual TEXT,
            archivo_recibo_pago TEXT,
            archivo_clabe TEXT,
            archivo_contrato TEXT,
            motivo_baja TEXT,
            archivo_acta_nacimiento TEXT,
            archivo_vigencia_imss TEXT,
            archivo_formato_solicitud TEXT
        )
    ''')

    for col_sql in ["ALTER TABLE registros ADD COLUMN motivo_baja TEXT", 
                    "ALTER TABLE registros ADD COLUMN archivo_acta_nacimiento TEXT", 
                    "ALTER TABLE registros ADD COLUMN archivo_vigencia_imss TEXT",
                    "ALTER TABLE registros ADD COLUMN archivo_formato_solicitud TEXT"]:
        try:
            cursor.execute(col_sql)
        except sqlite3.OperationalError:
            pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reportes_mensuales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alumno_co INTEGER,
            mes_anio TEXT NOT NULL,
            nombre_archivo TEXT NOT NULL,
            FOREIGN KEY (alumno_co) REFERENCES registros (co) ON DELETE CASCADE
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS recibos_alumno_mensuales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alumno_co INTEGER,
            mes_anio TEXT NOT NULL,
            nombre_archivo TEXT NOT NULL,
            FOREIGN KEY (alumno_co) REFERENCES registros (co) ON DELETE CASCADE
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS recibos_pago_mensuales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alumno_co INTEGER,
            mes_anio TEXT NOT NULL,
            nombre_archivo TEXT NOT NULL,
            FOREIGN KEY (alumno_co) REFERENCES registros (co) ON DELETE CASCADE
        )
    ''')

    conn.commit()
    conn.close()


@app.route('/')
def login():
    return render_template('login.html')


@app.route('/login', methods=['POST'])
def hacer_login():
    password = request.form.get('password')
    if password in USUARIOS_PERMITIDOS:
        session['usuario'] = USUARIOS_PERMITIDOS[password]['nombre']
        session['rol'] = USUARIOS_PERMITIDOS[password]['rol']
        return redirect(url_for('dashboard'))
    else:
        flash('Contraseña incorrecta. Intenta de nuevo.')
        return redirect(url_for('login'))


@app.route('/dashboard')
def dashboard():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    zona_filtro = request.args.get('zona_filtro')

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()

    if session['rol'] == 'admin':
        if zona_filtro and zona_filtro != 'TODOS':
            cursor.execute(
                'SELECT * FROM registros WHERE seccion = ? ORDER BY co ASC',
                (zona_filtro,),
            )
        else:
            cursor.execute('SELECT * FROM registros ORDER BY co ASC')
    else:
        cursor.execute(
            'SELECT * FROM registros WHERE seccion = ? ORDER BY co ASC',
            (session['usuario'],),
        )

    registros = cursor.fetchall()
    conn.close()
    return render_template(
        'dashboard.html',
        registros=registros,
        usuario=session['usuario'],
        rol=session['rol'],
        secciones_disponibles=list(USUARIOS_PERMITIDOS.values()),
        zona_filtro=zona_filtro,
    )


@app.route('/registrar', methods=['GET', 'POST'])
def registrar():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        seccion = session['usuario']
        nombre_alumno = request.form['nombre_alumno']
        tipo_apoyo = request.form['tipo_apoyo']
        area = request.form['area']
        institucion = request.form['institucion']
        carrera = request.form['carrera']
        fecha_inicio_str = request.form['fecha_inicio']
        fecha_fin_str = request.form['fecha_fin']
        horas = int(request.form['horas'])
        es_familiar = request.form.get('es_familiar', 'No')

        hoy = datetime.now().date()
        f_fin = datetime.strptime(fecha_fin_str, '%Y-%m-%d').date()
        estatus = 'Finalizado' if hoy > f_fin else 'Activo'

        if es_familiar == 'Sí':
            nombre_familiar = request.form.get('nombre_familiar', '')
            parentesco_familiar = request.form.get('parentesco_familiar', '')
            monto_apoyo = request.form.get('monto_apoyo', '2,580.86')
            tipo_remuneracion = f"Apoyo económico: ${monto_apoyo} pesos"
        else:
            nombre_familiar = 'N/A'
            parentesco_familiar = 'N/A'
            tipo_remuneracion = 'Título gratuito'

        conn = sqlite3.connect('datos.db')
        cursor = conn.cursor()
        cursor.execute(
            '''
                INSERT INTO registros (seccion, nombre_alumno, tipo_apoyo, area, institucion, carrera, fecha_inicio, fecha_fin, estatus, horas, es_familiar, nombre_familiar, parentesco_familiar, tipo_remuneracion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''',
            (
                seccion,
                nombre_alumno,
                tipo_apoyo,
                area,
                institucion,
                carrera,
                fecha_inicio_str,
                fecha_fin_str,
                estatus,
                horas,
                es_familiar,
                nombre_familiar,
                parentesco_familiar,
                tipo_remuneracion,
            ),
        )
        conn.commit()
        conn.close()

        flash('Estudiante registrado exitosamente.')
        return redirect(url_for('dashboard'))

    return render_template('registrar.html')


@app.route('/editar/<int:co>', methods=['GET', 'POST'])
def editar(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()

    if request.method == 'POST':
        nombre_alumno = request.form['nombre_alumno']
        tipo_apoyo = request.form['tipo_apoyo']
        area = request.form['area']
        institucion = request.form['institucion']
        carrera = request.form['carrera']
        fecha_inicio_str = request.form['fecha_inicio']
        fecha_fin_str = request.form['fecha_fin']
        horas = int(request.form['horas'])
        es_familiar = request.form.get('es_familiar', 'No')

        hoy = datetime.now().date()
        f_fin = datetime.strptime(fecha_fin_str, '%Y-%m-%d').date()
        estatus = 'Finalizado' if hoy > f_fin else 'Activo'

        if es_familiar == 'Sí':
            nombre_familiar = request.form.get('nombre_familiar', '')
            parentesco_familiar = request.form.get('parentesco_familiar', '')
            monto_apoyo = request.form.get('monto_apoyo', '2,580.86')
            tipo_remuneracion = f"Apoyo económico: ${monto_apoyo} pesos"
        else:
            nombre_familiar = 'N/A'
            parentesco_familiar = 'N/A'
            tipo_remuneracion = 'Título gratuito'

        if session['rol'] == 'admin':
            cursor.execute(
                '''
                    UPDATE registros SET nombre_alumno=?, tipo_apoyo=?, area=?, institucion=?, carrera=?, fecha_inicio=?, fecha_fin=?, estatus=?, horas=?, es_familiar=?, nombre_familiar=?, parentesco_familiar=?, tipo_remuneracion=?
                    WHERE co=?
                ''',
                (
                    nombre_alumno,
                    tipo_apoyo,
                    area,
                    institucion,
                    carrera,
                    fecha_inicio_str,
                    fecha_fin_str,
                    estatus,
                    horas,
                    es_familiar,
                    nombre_familiar,
                    parentesco_familiar,
                    tipo_remuneracion,
                    co,
                ),
            )
        else:
            cursor.execute(
                '''
                    UPDATE registros SET nombre_alumno=?, tipo_apoyo=?, area=?, institucion=?, carrera=?, fecha_inicio=?, fecha_fin=?, estatus=?, horas=?, es_familiar=?, nombre_familiar=?, parentesco_familiar=?, tipo_remuneracion=?
                    WHERE co=? AND seccion=?
                ''',
                (
                    nombre_alumno,
                    tipo_apoyo,
                    area,
                    institucion,
                    carrera,
                    fecha_inicio_str,
                    fecha_fin_str,
                    estatus,
                    horas,
                    es_familiar,
                    nombre_familiar,
                    parentesco_familiar,
                    tipo_remuneracion,
                    co,
                    session['usuario'],
                ),
            )

        conn.commit()
        conn.close()
        flash('Datos del alumno actualizados correctamente.')
        return redirect(url_for('dashboard'))

    if session['rol'] == 'admin':
        cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    else:
        cursor.execute(
            'SELECT * FROM registros WHERE co = ? AND seccion = ?',
            (co, session['usuario']),
        )

    alumno = cursor.fetchone()
    conn.close()

    if not alumno:
        flash('Registro no encontrado o sin permisos.')
        return redirect(url_for('dashboard'))

    return render_template('editar.html', alumno=alumno)


@app.route('/dar_de_baja/<int:co>', methods=['POST'])
def dar_de_baja(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    motivo = request.form.get('motivo_baja', '').strip()

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    if session['rol'] == 'admin':
        cursor.execute('UPDATE registros SET estatus = "Baja obligada", motivo_baja = ? WHERE co = ?', (motivo, co))
    else:
        cursor.execute('UPDATE registros SET estatus = "Baja obligada", motivo_baja = ? WHERE co = ? AND seccion = ?', (motivo, co, session['usuario']))
    
    conn.commit()
    conn.close()
    flash('Alumno dado de baja correctamente (Baja obligada).')
    return redirect(url_for('dashboard'))


@app.route('/subir_archivos/<int:co>', methods=['POST'])
def subir_archivos(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    banco = request.form.get('banco_clabe', '').strip()
    num_clabe = request.form.get('numero_clabe', '').strip()
    file_clabe = request.files.get('clabe')

    if (
        file_clabe
        and file_clabe.filename != ''
        or banco
        or len(num_clabe) > 0
    ):
        if not banco or len(num_clabe) != 18 or not num_clabe.isdigit():
            flash(
                'Error: Para la CLABE Interbancaria es obligatorio indicar el banco'
                ' y un número de exactamente 18 dígitos numéricos.'
            )
            return redirect(url_for('dashboard'))

    campos_archivos = {
        'carta_presentacion': 'archivo_carta_presentacion',
        'carta_aceptacion': 'archivo_carta_aceptacion',
        'constancia': 'archivo_constancia',
        'buena_conducta': 'archivo_buena_conducta',
        'ine': 'archivo_ine',
        'clabe': 'archivo_clabe',
        'contrato': 'archivo_contrato',
        'acta_nacimiento': 'archivo_acta_nacimiento',
        'vigencia_imss': 'archivo_vigencia_imss',
        'formato_solicitud': 'archivo_formato_solicitud',
    }

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()

    if banco and num_clabe:
        cursor.execute(
            'UPDATE registros SET banco_clabe = ?, numero_clabe = ? WHERE co = ?',
            (banco, num_clabe, co),
        )

    for input_name, db_col in campos_archivos.items():
        file = request.files.get(input_name)
        if file and file.filename != '':
            filename = secure_filename(f'alumno_{co}_{input_name}_{file.filename}')
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            cursor.execute(
                f'UPDATE registros SET {db_col} = ? WHERE co = ?', (filename, co)
            )
    conn.commit()
    conn.close()
    flash('Archivos y datos bancarios actualizados correctamente.')
    return redirect(url_for('dashboard'))


@app.route('/eliminar_archivo/<int:co>/<string:campo>')
def eliminar_archivo(co, campo):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    columnas_validas = {
        'carta_presentacion': 'archivo_carta_presentacion',
        'carta_aceptacion': 'archivo_carta_aceptacion',
        'constancia': 'archivo_constancia',
        'buena_conducta': 'archivo_buena_conducta',
        'ine': 'archivo_ine',
        'clabe': 'archivo_clabe',
        'contrato': 'archivo_contrato',
        'acta_nacimiento': 'archivo_acta_nacimiento',
        'vigencia_imss': 'archivo_vigencia_imss',
        'formato_solicitud': 'archivo_formato_solicitud',
    }

    if campo not in columnas_validas:
        flash('Campo de archivo no válido.')
        return redirect(url_for('dashboard'))

    db_col = columnas_validas[campo]

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute(f'SELECT {db_col} FROM registros WHERE co = ?', (co,))
    row = cursor.fetchone()

    if row and row[0]:
        filename = row[0]
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(filepath):
            os.remove(filepath)

        cursor.execute(f'UPDATE registros SET {db_col} = NULL WHERE co = ?', (co,))
        if db_col == 'archivo_clabe':
            cursor.execute(
                'UPDATE registros SET banco_clabe = NULL, numero_clabe = NULL WHERE'
                ' co = ?',
                (co,),
            )
        conn.commit()
        flash('Archivo eliminado correctamente.')

    conn.close()
    return redirect(url_for('dashboard'))


@app.route('/seleccionar_mes_orden/<int:co>')
def seleccionar_mes_orden(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()
    conn.close()

    if not r:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    f_inicio_str = r[7]
    f_fin_str = r[8]

    try:
        dt_actual = datetime.strptime(f_inicio_str, '%Y-%m-%d').replace(
            day=1
        )
        dt_final = datetime.strptime(f_fin_str, '%Y-%m-%d').replace(day=1)
    except:
        dt_actual = datetime(2026, 9, 1)
        dt_final = datetime(2027, 3, 1)

    meses_disponibles = []
    while dt_actual <= dt_final:
        mes_en = dt_actual.strftime('%B')
        mes_es = MESES_ESP.get(mes_en, mes_en).capitalize()
        anio = dt_actual.strftime('%Y')
        meses_disponibles.append({
            'valor': f'{mes_es} {anio}',
            'texto': f'{mes_es} {anio}',
        })
        if dt_actual.month == 12:
            dt_actual = dt_actual.replace(year=dt_actual.year + 1, month=1)
        else:
            dt_actual = dt_actual.replace(month=dt_actual.month + 1)

    return render_template(
        'seleccionar_mes.html', alumno=r, meses=meses_disponibles
    )


@app.route('/generar_orden_pago/<int:co>', methods=['POST'])
def generar_orden_pago(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mes_seleccionado = request.form.get(
        'mes_reportado', 'Septiembre 2026'
    )

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()

    if not r:
        conn.close()
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    nombre_alumno = r[2]
    tipo_apoyo = r[3].strip().lower()
    f_inicio_str = r[7]
    f_fin_str = r[8]
    zona = r[1]

    archivo_ine = r[21]
    archivo_clabe = r[24]
    archivo_acta = r[27] if len(r) > 27 else None
    archivo_vigencia_imss = r[28] if len(r) > 28 else None
    archivo_solicitud = r[29] if len(r) > 29 else None

    cursor.execute(
        'SELECT nombre_archivo FROM reportes_mensuales WHERE alumno_co = ? AND mes_anio = ?',
        (co, mes_seleccionado)
    )
    rep_row = cursor.fetchone()

    cursor.execute(
        'SELECT nombre_archivo FROM recibos_alumno_mensuales WHERE alumno_co = ? AND mes_anio = ?',
        (co, mes_seleccionado)
    )
    recibo_alum_row = cursor.fetchone()

    cursor.execute(
        'SELECT nombre_archivo FROM recibos_pago_mensuales WHERE alumno_co = ? AND mes_anio = ?',
        (co, mes_seleccionado)
    )
    recibo_emp_row = cursor.fetchone()
    conn.close()

    archivo_reporte_mes = rep_row[0] if rep_row else None
    archivo_recibo_alum = recibo_alum_row[0] if recibo_alum_row else None
    archivo_recibo_emp = recibo_emp_row[0] if recibo_emp_row else None

    try:
        dt_inicio = datetime.strptime(f_inicio_str, '%Y-%m-%d')
        dt_fin = datetime.strptime(f_fin_str, '%Y-%m-%d')
        f_inicio = dt_inicio.strftime('%d/%m/%Y')
        f_fin = dt_fin.strftime('%d/%m/%Y')
    except:
        f_inicio = '01/09/2026'
        f_fin = '01/03/2027'

    mes_reportado = mes_seleccionado
    hoy_str = fecha_a_espanol(datetime.now())

    x_practicas = ''
    x_servicio = ''
    x_residencias = ''

    if 'práctica' in tipo_apoyo or 'practica' in tipo_apoyo:
        x_practicas = 'X'
    elif 'servicio' in tipo_apoyo:
        x_servicio = 'X'
    elif (
        'residencia' in tipo_apoyo
        or 'estadía' in tipo_apoyo
        or 'estadia' in tipo_apoyo
    ):
        x_residencias = 'X'
    else:
        x_servicio = 'X'

    plantilla_path = 'RECIBO PAGO.docx'
    if not os.path.exists(plantilla_path):
        flash('Error: No se encontró el archivo "RECIBO PAGO.docx".')
        return redirect(url_for('dashboard'))

    doc = docx.Document(plantilla_path)

    reemplazos = {
        '[PERIODO]': f'Periodo: {f_inicio} al {f_fin}',
        '[MES_REPORTADO]': f'Mes reportado: {mes_reportado}',
        '[NOMBRE_ALUMNO]': f'C. {nombre_alumno}',
        '[FECHA_ZONA]': f'{zona}, Ver., a {hoy_str}.',
        '[X_PRACTICAS]': x_practicas,
        '[X_SERV]': x_servicio,
        '[X_SERVICIO]': x_servicio,
        '[X_SERVICIOS]': x_servicio,
        '[X_RES]': x_residencias,
        '[X_RESIDENCIAS]': x_residencias,
    }

    def procesar_parrafo_obj(parrafo):
        texto_p = parrafo.text
        hubo_cambio = False
        for clave in reemplazos:
            if clave in texto_p:
                texto_p = texto_p.replace(clave, reemplazos[clave])
                hubo_cambio = True

        if hubo_cambio:
            if parrafo.runs:
                estilo_fuente = parrafo.runs[0].font.name
                tamano_fuente = parrafo.runs[0].font.size
                subrayado = parrafo.runs[0].underline
                parrafo.runs[0].text = texto_p
                if estilo_fuente:
                    parrafo.runs[0].font.name = estilo_fuente
                if tamano_fuente:
                    parrafo.runs[0].font.size = tamano_fuente
                if subrayado is not None:
                    parrafo.runs[0].underline = subrayado
                for run in parrafo.runs[1:]:
                    run.text = ''
            else:
                parrafo.add_run(texto_p)

    for p in doc.paragraphs:
        procesar_parrafo_obj(p)

    def procesar_celdas(cells):
        for cell in cells:
            for p in cell.paragraphs:
                procesar_parrafo_obj(p)
            if cell.tables:
                for sub_table in cell.tables:
                    for sub_row in sub_table.rows:
                        procesar_celdas(sub_row.cells)

    for table in doc.tables:
        for row in table.rows:
            procesar_celdas(row.cells)

    def buscar_y_reemplazar_en_todos_los_nodos(element):
        for child in element:
            if (
                'txbxContent' in child.tag
                or 'textbox' in child.tag
                or 'shape' in child.tag
                or 'graphic' in child.tag
            ):
                for p_elem in child.iter(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
                ):
                    p_obj = Paragraph(p_elem, doc)
                    procesar_parrafo_obj(p_obj)
            buscar_y_reemplazar_en_todos_los_nodos(child)

    buscar_y_reemplazar_en_todos_los_nodos(doc.element.body)

    for section in doc.sections:
        for p in section.header.paragraphs:
            procesar_parrafo_obj(p)
        for p in section.footer.paragraphs:
            procesar_parrafo_obj(p)

    output_filename = f'Recibo_Pago_{nombre_alumno.replace(" ", "_")}.docx'
    doc.save(output_filename)

    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(output_filename, arcname=output_filename)
        
        if archivo_ine:
            path_ine = os.path.join(app.config['UPLOAD_FOLDER'], archivo_ine)
            if os.path.exists(path_ine):
                zf.write(path_ine, arcname=f"INE_{archivo_ine}")

        if archivo_acta:
            path_acta = os.path.join(app.config['UPLOAD_FOLDER'], archivo_acta)
            if os.path.exists(path_acta):
                zf.write(path_acta, arcname=f"ActaNacimiento_{archivo_acta}")

        if archivo_vigencia_imss:
            path_imss = os.path.join(app.config['UPLOAD_FOLDER'], archivo_vigencia_imss)
            if os.path.exists(path_imss):
                zf.write(path_imss, arcname=f"VigenciaIMSS_{archivo_vigencia_imss}")

        if archivo_solicitud:
            path_sol = os.path.join(app.config['UPLOAD_FOLDER'], archivo_solicitud)
            if os.path.exists(path_sol):
                zf.write(path_sol, arcname=f"FormatoSolicitud_{archivo_solicitud}")
                
        if archivo_reporte_mes:
            path_rep = os.path.join(app.config['UPLOAD_FOLDER'], archivo_reporte_mes)
            if os.path.exists(path_rep):
                zf.write(path_rep, arcname=f"ReporteMensual_{mes_seleccionado.replace(' ', '_')}_{archivo_reporte_mes}")

        if archivo_recibo_alum:
            path_rec_alum = os.path.join(app.config['UPLOAD_FOLDER'], archivo_recibo_alum)
            if os.path.exists(path_rec_alum):
                zf.write(path_rec_alum, arcname=f"ReciboPagoAlumno_{mes_seleccionado.replace(' ', '_')}_{archivo_recibo_alum}")

        if archivo_recibo_emp:
            path_rec_emp = os.path.join(app.config['UPLOAD_FOLDER'], archivo_recibo_emp)
            if os.path.exists(path_rec_emp):
                zf.write(path_rec_emp, arcname=f"OficioPagoEmpresa_{mes_seleccionado.replace(' ', '_')}_{archivo_recibo_emp}")
                
        if archivo_clabe:
            path_clabe = os.path.join(app.config['UPLOAD_FOLDER'], archivo_clabe)
            if os.path.exists(path_clabe):
                zf.write(path_clabe, arcname=f"CLABE_{archivo_clabe}")

    memory_file.seek(0)

    if os.path.exists(output_filename):
        os.remove(output_filename)

    return send_file(
        memory_file,
        as_attachment=True,
        download_name=f'Paquete_Pago_{nombre_alumno.replace(" ", "_")}.zip',
        mimetype='application/zip'
    )


@app.route('/generar_convenio_gratuito/<int:co>')
def generar_convenio_gratuito(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()
    conn.close()

    if not r:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    zona = r[1]
    nombre_alumno = r[2]
    tipo_apoyo = r[3]
    institucion = r[5]
    carrera = r[6]
    f_inicio_str = r[7]
    f_fin_str = r[8]

    try:
        dt_inicio = datetime.strptime(f_inicio_str, '%Y-%m-%d')
        dt_fin = datetime.strptime(f_fin_str, '%Y-%m-%d')
        f_inicio_formato = f"{dt_inicio.strftime('%d')} DE {MESES_ESP.get(dt_inicio.strftime('%B'), dt_inicio.strftime('%B')).upper()} DEL {dt_inicio.strftime('%Y')}"
        f_fin_formato = f"{dt_fin.strftime('%d')} DE {MESES_ESP.get(dt_fin.strftime('%B'), dt_fin.strftime('%B')).upper()} DEL {dt_fin.strftime('%Y')}"
    except:
        f_inicio_formato = "02 DE FEBRERO DEL 2026"
        f_fin_formato = "02 DE AGOSTO DEL 2026"

    dt_hoy = datetime.now()
    dia_hoy = dt_hoy.strftime('%d')
    mes_hoy = MESES_ESP.get(dt_hoy.strftime('%B'), dt_hoy.strftime('%B')).upper()
    anio_hoy = dt_hoy.strftime('%Y')
    f_actual_formato = f"{dia_hoy} DIAS DEL MES DE {mes_hoy} DEL {anio_hoy}"

    plantilla_convenio = 'CONVENIO A TITULO GRATUITO.docx'
    if not os.path.exists(plantilla_convenio):
        flash('Error: No se encontró el archivo "CONVENIO A TITULO GRATUITO.docx".')
        return redirect(url_for('dashboard'))

    doc = docx.Document(plantilla_convenio)

    reemplazos = {
        '[TIPO_DE_APOYO]': tipo_apoyo.upper(),
        '[TIPO- DE- APOYO]': tipo_apoyo.upper(),
        '[NOMBRE_ALUMNO]': nombre_alumno,
        '[CARRERA]': carrera,
        '[INSTITUCION_EDUCATIVA]': institucion,
        '[INSTITUCIÓN EDUCATIVA]': institucion,
        '[FECHA_INICIO]': f_inicio_formato,
        '[FECHA_FIN]': f_fin_formato,
        '[FECHA_ACTUAL]': f_actual_formato,
        'Xalapa, Veracruz': zona,
        'XALAPA, VERACRUZ': zona.upper(),
    }

    def procesar_parrafo_obj(parrafo):
        texto_p = parrafo.text
        hubo_cambio = False
        for clave in reemplazos:
            if clave in texto_p:
                texto_p = texto_p.replace(clave, reemplazos[clave])
                hubo_cambio = True
        if hubo_cambio:
            if parrafo.runs:
                estilo_fuente = parrafo.runs[0].font.name
                tamano_fuente = parrafo.runs[0].font.size
                subrayado = parrafo.runs[0].underline
                parrafo.runs[0].text = texto_p
                if estilo_fuente:
                    parrafo.runs[0].font.name = estilo_fuente
                if tamano_fuente:
                    parrafo.runs[0].font.size = tamano_fuente
                if subrayado is not None:
                    parrafo.runs[0].underline = subrayado
                for run in parrafo.runs[1:]:
                    run.text = ''
            else:
                parrafo.add_run(texto_p)

    for p in doc.paragraphs:
        procesar_parrafo_obj(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    procesar_parrafo_obj(p)

    output_filename = f'Convenio_Gratuito_{nombre_alumno.replace(" ", "_")}.docx'
    doc.save(output_filename)

    return send_file(
        output_filename,
        as_attachment=True,
        download_name=output_filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


# ==========================================
# GESTIÓN DE OFICIOS DE PAGO POR ZONA
# ==========================================
@app.route('/gestionar_oficios_zona', methods=['GET'])
def gestionar_oficios_zona():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    usuario_actual = session['usuario']
    rol_usuario = session['rol']
    
    if rol_usuario == 'admin':
        zona_seleccionada = request.args.get('zona_seleccionada', 'Oficinas Divisionales')
    else:
        zona_seleccionada = usuario_actual

    carpeta_zona = os.path.join('static', 'oficios_pago_zona', zona_seleccionada)
    os.makedirs(carpeta_zona, exist_ok=True)
    
    archivos_zona = []
    if os.path.exists(carpeta_zona):
        archivos_zona = os.listdir(carpeta_zona)

    todas_las_zonas = [sec['nombre'] for sec in USUARIOS_PERMITIDOS.values()]

    return render_template(
        'gestionar_oficios_zona.html',
        zona_seleccionada=zona_seleccionada,
        archivos_zona=archivos_zona,
        todas_las_zonas=todas_las_zonas,
        rol=rol_usuario,
        usuario=usuario_actual
    )


@app.route('/subir_oficio_zona', methods=['POST'])
def subir_oficio_zona():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    zona_destino = request.form.get('zona_destino')
    rol_usuario = session['rol']
    usuario_actual = session['usuario']

    if rol_usuario != 'admin' and zona_destino != usuario_actual:
        flash('No tienes permisos para subir archivos a otra zona.')
        return redirect(url_for('gestionar_oficios_zona'))

    file = request.files.get('archivo_oficio_zona')
    if file and file.filename != '':
        filename = secure_filename(file.filename)
        carpeta_zona = os.path.join('static', 'oficios_pago_zona', zona_destino)
        os.makedirs(carpeta_zona, exist_ok=True)
        file.save(os.path.join(carpeta_zona, filename))
        flash(f'Oficio de pago subido correctamente para la zona {zona_destino}.')

    return redirect(url_for('gestionar_oficios_zona', zona_seleccionada=zona_destino))


@app.route('/eliminar_oficio_zona/<string:zona>/<string:nombre_archivo>')
def eliminar_oficio_zona(zona, nombre_archivo):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    rol_usuario = session['rol']
    usuario_actual = session['usuario']

    if rol_usuario != 'admin' and zona != usuario_actual:
        flash('No tienes permisos para eliminar archivos de otra zona.')
        return redirect(url_for('gestionar_oficios_zona'))

    filepath = os.path.join('static', 'oficios_pago_zona', zona, nombre_archivo)
    if os.path.exists(filepath):
        os.remove(filepath)
        flash('Oficio de pago eliminado correctamente.')

    return redirect(url_for('gestionar_oficios_zona', zona_seleccionada=zona))


# ==========================================
# GENERADOR DEL EXCEL NORMA 15 (MÁXIMO 20 ALUMNOS POR HOJA, MEZCLANDO ZONAS)
# ==========================================
FILA_INICIO = 15          # primera fila de datos en la plantilla
FILAS_BASE = 20           # filas con borde que trae la plantilla (15 a 34 -> exactamente 20 alumnos)
ESTILO_DATOS = 9          # estilo de celda de datos de la plantilla
COLUMNAS = "CDEFGHIJKLMN" # C=No, D=AÑO, E=Mes, F=Coordinación, G=Centro, H=Tipo, I=Nombre, J=Inst, K=Carrera, L=Inicio, M=Fin, N=Monto

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _fecha_dmy(txt):
    try:
        return datetime.strptime(str(txt), "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "" if txt is None else str(txt)


def _celda(ref, valor):
    if valor is None or valor == "":
        return f'<c r="{ref}" s="{ESTILO_DATOS}"/>'
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}" s="{ESTILO_DATOS}"><v>{valor}</v></c>'
    return (f'<c r="{ref}" s="{ESTILO_DATOS}" t="inlineStr">'
            f'<is><t xml:space="preserve">{escape(str(valor))}</t></is></c>')


def _fila_datos(num_fila, valores):
    celdas = "".join(_celda(f"{col}{num_fila}", v) for col, v in zip(COLUMNAS, valores))
    return f'<row r="{num_fila}">{celdas}</row>'


def _desplazar_fila(xml_fila, delta):
    def nueva(m):
        return m.group(1) + str(int(m.group(2)) + delta)
    xml_fila = re.sub(r'(<row r=")(\d+)', nueva, xml_fila, count=1)
    xml_fila = re.sub(r'(<c r="[A-Z]+)(\d+)', nueva, xml_fila)
    return xml_fila


def _construir_hoja(sheet_tpl, filas_valores, primera_hoja):
    n = len(filas_valores)
    total_filas = max(n, FILAS_BASE)
    extra = total_filas - FILAS_BASE
    ultima_fila_datos = FILA_INICIO + FILAS_BASE - 1

    m = re.search(r"<sheetData>(.*)</sheetData>", sheet_tpl, re.S)
    filas = re.findall(r'<row\b[^>]*?(?:/>|>.*?</row>)', m.group(1), re.S)
    por_num = {}
    for f in filas:
        match_r = re.search(r'r="(\d+)"', f)
        if match_r:
            por_num[int(match_r.group(1))] = f

    salida = []
    for r in sorted(por_num):
        if r < FILA_INICIO:
            salida.append(por_num[r])
            
    for i in range(total_filas):
        vals = filas_valores[i] if i < n else [None] * len(COLUMNAS)
        salida.append(_fila_datos(FILA_INICIO + i, vals))
        
    for r in sorted(por_num):
        if r > ultima_fila_datos:
            salida.append(_desplazar_fila(por_num[r], extra) if extra else por_num[r])

    xml = sheet_tpl[: m.start(1)] + "".join(salida) + sheet_tpl[m.end(1):]

    if extra:
        def mover_merge(mm):
            def f(c):
                col, row = c.group(1), int(c.group(2))
                return f"{col}{row + extra if row > ultima_fila_datos else row}"
            return re.sub(r"([A-Z]+)(\d+)", f, mm.group(0))
        xml = re.sub(r'<mergeCell ref="[^"]+"', mover_merge, xml)
        xml = re.sub(r'<dimension ref="C1:O(\d+)"',
                     lambda d: f'<dimension ref="C1:O{int(d.group(1)) + extra}"', xml)

    if not primera_hoja:
        xml = xml.replace(' tabSelected="1"', "")
    return xml


def generar_norma15(registros, plantilla_path, output_path, anio=None, mes_es=None):
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    hoy = datetime.now()
    anio = str(anio or hoy.year)
    mes_es = mes_es or meses[hoy.month - 1]

    if not registros:
        registros = []

    # Dividir todos los registros en bloques exactos de máximo 20 alumnos por hoja
    tamano_bloque = 20
    bloques = [registros[i:i + tamano_bloque] for i in range(0, len(registros), tamano_bloque)]
    if not bloques:
        bloques = [[]]

    zin = zipfile.ZipFile(plantilla_path)
    arch = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    sheet_tpl = arch["xl/worksheets/sheet1.xml"].decode("utf-8")
    drawing_tpl = arch["xl/drawings/drawing1.xml"]
    drawing_rels_tpl = arch["xl/drawings/_rels/drawing1.xml.rels"]
    comments_tpl = arch["xl/comments1.xml"]
    vml_tpl = arch["xl/drawings/vmlDrawing1.vml"].decode("utf-8")

    for k in ["xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels",
              "xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels",
              "xl/comments1.xml", "xl/drawings/vmlDrawing1.vml"]:
        arch.pop(k, None)

    hojas_meta = []
    for idx, bloque in enumerate(bloques, start=1):
        nombre = f"Informe {idx}"
        hojas_meta.append((nombre, idx))

        filas_valores = []
        for i, r in enumerate(bloque, start=1):
            filas_valores.append([
                i, anio, mes_es, "Distribución", r[1], r[3], r[2], r[5], r[6],
                _fecha_dmy(r[7]), _fecha_dmy(r[8]), r[14],
            ])

        arch[f"xl/worksheets/sheet{idx}.xml"] = _construir_hoja(
            sheet_tpl, filas_valores, idx == 1).encode("utf-8")

        arch[f"xl/worksheets/_rels/sheet{idx}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<Relationships xmlns="{NS_REL}">'
            f'<Relationship Id="rId3" Type="{REL_BASE}/comments" Target="../comments{idx}.xml"/>'
            f'<Relationship Id="rId2" Type="{REL_BASE}/vmlDrawing" Target="../drawings/vmlDrawing{idx}.vml"/>'
            f'<Relationship Id="rId1" Type="{REL_BASE}/drawing" Target="../drawings/drawing{idx}.xml"/>'
            '</Relationships>').encode("utf-8")

        arch[f"xl/drawings/drawing{idx}.xml"] = drawing_tpl
        arch[f"xl/drawings/_rels/drawing{idx}.xml.rels"] = drawing_rels_tpl
        arch[f"xl/comments{idx}.xml"] = comments_tpl

        vml = re.sub(r'(idmap v:ext="edit" data=")\d+(")', rf'\g<1>{idx}\2', vml_tpl)
        vml = re.sub(r"_x0000_s(\d+)",
                     lambda m: f"_x0000_s{int(m.group(1)) + (idx - 1) * 1024}", vml)
        arch[f"xl/drawings/vmlDrawing{idx}.vml"] = vml.encode("utf-8")

    wb = arch["xl/workbook.xml"].decode("utf-8")
    sheets_xml = "<sheets>" + "".join(
        f'<sheet name="{escape(n, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rIdS{i}"/>'
        for n, i in hojas_meta) + "</sheets>"
    wb = re.sub(r"<sheets>.*?</sheets>", sheets_xml, wb, flags=re.S)
    arch["xl/workbook.xml"] = wb.encode("utf-8")

    rels = arch["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship Id="rId1"[^>]*worksheets/sheet1\.xml"/>', "", rels)
    nuevas = "".join(
        f'<Relationship Id="rIdS{i}" Type="{REL_BASE}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for _, i in hojas_meta)
    rels = rels.replace("</Relationships>", nuevas + "</Relationships>")
    arch["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    ct = arch["[Content_Types].xml"].decode("utf-8")
    for pat in [r'<Override PartName="/xl/worksheets/sheet1\.xml"[^>]*/>',
                r'<Override PartName="/xl/drawings/drawing1\.xml"[^>]*/>',
                r'<Override PartName="/xl/comments1\.xml"[^>]*/>']:
        ct = re.sub(pat, "", ct)
    extra_ct = ""
    for _, i in hojas_meta:
        extra_ct += (
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            f'<Override PartName="/xl/drawings/drawing{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            f'<Override PartName="/xl/comments{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>')
    ct = ct.replace("</Types>", extra_ct + "</Types>")
    arch["[Content_Types].xml"] = ct.encode("utf-8")

    app = arch["docProps/app.xml"].decode("utf-8")
    titulos = "".join(f"<vt:lpstr>{escape(n)}</vt:lpstr>" for n, _ in hojas_meta)
    app = re.sub(r"<vt:i4>\d+</vt:i4>", f"<vt:i4>{len(hojas_meta)}</vt:i4>", app, count=1)
    app = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>',
                 f'<TitlesOfParts><vt:vector size="{len(hojas_meta)}" baseType="lpstr">{titulos}</vt:vector></TitlesOfParts>',
                 app, flags=re.S)
    arch["docProps/app.xml"] = app.encode("utf-8")

    orden = ["[Content_Types].xml"] + [k for k in arch if k != "[Content_Types].xml"]
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for k in orden:
            zout.writestr(k, arch[k])
    return output_path


# ==========================================
# DESCARGA DE EXCEL NORMA 15
# ==========================================
@app.route('/descargar_excel')
def descargar_excel():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros ORDER BY seccion ASC, co ASC')
    registros = cursor.fetchall()
    conn.close()

    base_dir = os.path.abspath(os.path.dirname(__file__))
    plantilla_excel = os.path.join(base_dir, 'plantillareporte_norma15.xlsx')
    output_filename = 'Informe_Norma15_Por_Zonas.xlsx'
    output_path = os.path.join(base_dir, output_filename)

    try:
        if not os.path.exists(plantilla_excel):
            flash('Error: No se encontró la plantilla "plantillareporte_norma15.xlsx".')
            return redirect(url_for('dashboard'))

        generar_norma15(registros, plantilla_excel, output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        flash(f'Error al generar el reporte con Excel: {str(e)}')
        return redirect(url_for('dashboard'))


# ==========================================
# DESCARGA DE EXCEL TODOS / HISTORIAL
# ==========================================
@app.route('/descargar_excel_todos')
def descargar_excel_todos():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros ORDER BY co ASC')
    registros = cursor.fetchall()
    conn.close()

    base_dir = os.path.abspath(os.path.dirname(__file__))
    plantilla_excel = os.path.join(base_dir, 'plantillareporte_todos.xlsx')
    output_filename = 'Informe_Historial_Todos_Alumnos.xlsx'
    output_path = os.path.join(base_dir, output_filename)

    try:
        if not os.path.exists(plantilla_excel):
            flash('Error: No se encontró la plantilla "plantillareporte_todos.xlsx".')
            return redirect(url_for('dashboard'))

        wb = openpyxl.load_workbook(plantilla_excel)
        ws = wb.active

        start_row = 4
        row_idx = start_row
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for idx, r in enumerate(registros, start=1):
            valores_fila = [
                idx, r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[14]
            ]

            for c_offset, val in enumerate(valores_fila):
                col_num = 1 + c_offset
                cell = ws.cell(row=row_idx, column=col_num)
                
                if isinstance(cell, MergedCell):
                    for range_merged in ws.merged_cells.ranges:
                        if cell.coordinate in range_merged:
                            cell = ws.cell(row=range_merged.min_row, column=range_merged.min_col)
                            break

                cell.value = val
                cell.font = Font(name='Arial', size=9, color='000000')
                cell.border = thin_border

                if col_num in [1, 7, 8, 9, 10, 11]:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            row_idx += 1

        if os.path.exists(output_path):
            os.remove(output_path)

        wb.save(output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        flash(f'Error al generar el Excel de Historial: {str(e)}')
        return redirect(url_for('dashboard'))


# ==========================================
# VENTANA 1: GESTIÓN DE REPORTES MENSUALES (ALUMNO)
# ==========================================
@app.route('/gestionar_reportes/<int:co>', methods=['GET'])
def gestionar_reportes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()

    if not r:
        conn.close()
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    f_inicio_str = r[7]
    f_fin_str = r[8]

    try:
        dt_actual = datetime.strptime(f_inicio_str, '%Y-%m-%d').replace(day=1)
        dt_final = datetime.strptime(f_fin_str, '%Y-%m-%d').replace(day=1)
    except:
        dt_actual = datetime(2026, 10, 1)
        dt_final = datetime(2027, 5, 1)

    meses_disponibles = []
    while dt_actual <= dt_final:
        mes_en = dt_actual.strftime('%B')
        mes_es = MESES_ESP.get(mes_en, mes_en).capitalize()
        anio = dt_actual.strftime('%Y')
        meses_disponibles.append(f'{mes_es} {anio}')
        if dt_actual.month == 12:
            dt_actual = dt_actual.replace(year=dt_actual.year + 1, month=1)
        else:
            dt_actual = dt_actual.replace(month=dt_actual.month + 1)

    cursor.execute(
        'SELECT id, mes_anio, nombre_archivo FROM reportes_mensuales WHERE alumno_co = ?',
        (co,),
    )
    reportes_subidos = cursor.fetchall()
    conn.close()

    return render_template(
        'gestionar_reportes.html',
        alumno=r,
        meses_disponibles=meses_disponibles,
        reportes_subidos=reportes_subidos,
    )


@app.route('/subir_reporte_mes/<int:co>', methods=['POST'])
def subir_reporte_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_reporte')

    if file and file.filename != '' and mes_anio:
        filename = secure_filename(
            f'alumno_{co}_reporte_{mes_anio.replace(" ", "_")}_{file.filename}'
        )
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        conn = sqlite3.connect('datos.db')
        cursor = conn.cursor()
        cursor.execute(
            'DELETE FROM reportes_mensuales WHERE alumno_co = ? AND mes_anio = ?',
            (co, mes_anio),
        )
        cursor.execute(
            'INSERT INTO reportes_mensuales (alumno_co, mes_anio, nombre_archivo) VALUES (?, ?, ?)',
            (co, mes_anio, filename),
        )
        conn.commit()
        conn.close()
        flash(f'Reporte mensual de {mes_anio} subido correctamente.')

    return redirect(url_for('gestionar_reportes', co=co))


@app.route('/eliminar_reporte_mes/<int:reporte_id>/<int:co>')
def eliminar_reporte_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute(
        'SELECT nombre_archivo FROM reportes_mensuales WHERE id = ?',
        (reporte_id,),
    )
    row = cursor.fetchone()
    if row:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], row[0])
        if os.path.exists(filepath):
            os.remove(filepath)
        cursor.execute(
            'DELETE FROM reportes_mensuales WHERE id = ?', (reporte_id,)
        )
        conn.commit()
    conn.close()
    flash('Reporte mensual eliminado.')
    return redirect(url_for('gestionar_reportes', co=co))


# ==========================================
# VENTANA 2: GESTIÓN DE RECIBOS DE PAGO DEL ALUMNO
# ==========================================
@app.route('/gestionar_recibos_alumno/<int:co>', methods=['GET'])
def gestionar_recibos_alumno(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()

    if not r:
        conn.close()
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    f_inicio_str = r[7]
    f_fin_str = r[8]

    try:
        dt_actual = datetime.strptime(f_inicio_str, '%Y-%m-%d').replace(day=1)
        dt_final = datetime.strptime(f_fin_str, '%Y-%m-%d').replace(day=1)
    except:
        dt_actual = datetime(2026, 10, 1)
        dt_final = datetime(2027, 5, 1)

    meses_disponibles = []
    while dt_actual <= dt_final:
        mes_en = dt_actual.strftime('%B')
        mes_es = MESES_ESP.get(mes_en, mes_en).capitalize()
        anio = dt_actual.strftime('%Y')
        meses_disponibles.append(f'{mes_es} {anio}')
        if dt_actual.month == 12:
            dt_actual = dt_actual.replace(year=dt_actual.year + 1, month=1)
        else:
            dt_actual = dt_actual.replace(month=dt_actual.month + 1)

    cursor.execute(
        'SELECT id, mes_anio, nombre_archivo FROM recibos_alumno_mensuales WHERE alumno_co = ?',
        (co,),
    )
    recibos_subidos = cursor.fetchall()
    conn.close()

    return render_template(
        'gestionar_recibos_alumno.html',
        alumno=r,
        meses_disponibles=meses_disponibles,
        recibos_subidos=recibos_subidos,
    )


@app.route('/subir_recibo_alumno_mes/<int:co>', methods=['POST'])
def subir_recibo_alumno_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_recibo_alumno')

    if file and file.filename != '' and mes_anio:
        filename = secure_filename(
            f'alumno_{co}_reciboalumno_{mes_anio.replace(" ", "_")}_{file.filename}'
        )
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        conn = sqlite3.connect('datos.db')
        cursor = conn.cursor()
        cursor.execute(
            'DELETE FROM recibos_alumno_mensuales WHERE alumno_co = ? AND mes_anio = ?',
            (co, mes_anio),
        )
        cursor.execute(
            'INSERT INTO recibos_alumno_mensuales (alumno_co, mes_anio, nombre_archivo) VALUES (?, ?, ?)',
            (co, mes_anio, filename),
        )
        conn.commit()
        conn.close()
        flash(f'Recibo de pago del alumno de {mes_anio} subido correctamente.')

    return redirect(url_for('gestionar_recibos_alumno', co=co))


@app.route('/eliminar_recibo_alumno_mes/<int:reporte_id>/<int:co>')
def eliminar_recibo_alumno_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute(
        'SELECT nombre_archivo FROM recibos_alumno_mensuales WHERE id = ?',
        (reporte_id,),
    )
    row = cursor.fetchone()
    if row:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], row[0])
        if os.path.exists(filepath):
            os.remove(filepath)
        cursor.execute(
            'DELETE FROM recibos_alumno_mensuales WHERE id = ?', (reporte_id,)
        )
        conn.commit()
    conn.close()
    flash('Recibo de pago del alumno eliminado.')
    return redirect(url_for('gestionar_recibos_alumno', co=co))


# ==========================================
# VENTANA 3: GESTIÓN DE OFICIOS DE PAGO (EMPRESA)
# ==========================================
@app.route('/gestionar_recibos_pago/<int:co>', methods=['GET'])
def gestionar_recibos_pago(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM registros WHERE co = ?', (co,))
    r = cursor.fetchone()

    if not r:
        conn.close()
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    f_inicio_str = r[7]
    f_fin_str = r[8]

    try:
        dt_actual = datetime.strptime(f_inicio_str, '%Y-%m-%d').replace(day=1)
        dt_final = datetime.strptime(f_fin_str, '%Y-%m-%d').replace(day=1)
    except:
        dt_actual = datetime(2026, 10, 1)
        dt_final = datetime(2027, 5, 1)

    meses_disponibles = []
    while dt_actual <= dt_final:
        mes_en = dt_actual.strftime('%B')
        mes_es = MESES_ESP.get(mes_en, mes_en).capitalize()
        anio = dt_actual.strftime('%Y')
        meses_disponibles.append(f'{mes_es} {anio}')
        if dt_actual.month == 12:
            dt_actual = dt_actual.replace(year=dt_actual.year + 1, month=1)
        else:
            dt_actual = dt_actual.replace(month=dt_actual.month + 1)

    cursor.execute(
        'SELECT id, mes_anio, nombre_archivo FROM recibos_pago_mensuales WHERE alumno_co = ?',
        (co,),
    )
    recibos_subidos = cursor.fetchall()
    conn.close()

    return render_template(
        'gestionar_recibos_pago.html',
        alumno=r,
        meses_disponibles=meses_disponibles,
        recibos_subidos=recibos_subidos,
    )


@app.route('/subir_recibo_pago_mes/<int:co>', methods=['POST'])
def subir_recibo_pago_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_recibo')

    if file and file.filename != '' and mes_anio:
        filename = secure_filename(
            f'alumno_{co}_oficiopago_{mes_anio.replace(" ", "_")}_{file.filename}'
        )
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        conn = sqlite3.connect('datos.db')
        cursor = conn.cursor()
        cursor.execute(
            'DELETE FROM recibos_pago_mensuales WHERE alumno_co = ? AND mes_anio = ?',
            (co, mes_anio),
        )
        cursor.execute(
            'INSERT INTO recibos_pago_mensuales (alumno_co, mes_anio, nombre_archivo) VALUES (?, ?, ?)',
            (co, mes_anio, filename),
        )
        conn.commit()
        conn.close()
        flash(f'Oficio de pago de la empresa de {mes_anio} subido correctamente.')

    return redirect(url_for('gestionar_recibos_pago', co=co))


@app.route('/eliminar_recibo_pago_mes/<int:reporte_id>/<int:co>')
def eliminar_recibo_pago_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    conn = sqlite3.connect('datos.db')
    cursor = conn.cursor()
    cursor.execute(
        'SELECT nombre_archivo FROM recibos_pago_mensuales WHERE id = ?',
        (reporte_id,),
    )
    row = cursor.fetchone()
    if row:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], row[0])
        if os.path.exists(filepath):
            os.remove(filepath)
        cursor.execute(
            'DELETE FROM recibos_pago_mensuales WHERE id = ?', (reporte_id,)
        )
        conn.commit()
    conn.close()
    flash('Oficio de pago eliminado.')
    return redirect(url_for('gestionar_recibos_pago', co=co))


if __name__ == '__main__':
    init_db()
    app.run(debug=True)
    