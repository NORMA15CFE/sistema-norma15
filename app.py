from datetime import datetime
import copy
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
from flask_sqlalchemy import SQLAlchemy

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage

from xml.sax.saxutils import escape
from werkzeug.utils import secure_filename

# ==========================================
# IMPORTACIONES DE CLOUDINARY
# ==========================================
import cloudinary
import cloudinary.uploader
import cloudinary.api

app = Flask(__name__)
app.secret_key = 'clave_secreta_cfe_norma15'

UPLOAD_FOLDER = 'static/uploads_alumnos'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ==========================================
# CONFIGURACIÓN DE CLOUDINARY
# ==========================================
cloudinary.config(
    cloud_name=os.environ.get('CLOUDINARY_CLOUD_NAME'),
    api_key=os.environ.get('CLOUDINARY_API_KEY'),
    api_secret=os.environ.get('CLOUDINARY_API_SECRET')
)

# ==========================================
# CONFIGURACIÓN DE BASE DE DATOS (SQLALCHEMY)
# ==========================================
DATABASE_URL = os.environ.get('DATABASE_URL')
if DATABASE_URL:
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
    elif DATABASE_URL.startswith("postgresql://"):
        DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)

app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL or 'sqlite:///datos.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

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

# ==========================================
# CONFIGURACIÓN DEL REPORTE EXCEL (PLANTILLAS)
# ==========================================
PLANTILLA_REPORTE = 'plantillareporte_norma15.xlsx'
PLANTILLA_TODOS = 'plantillareporte_todos.xlsx'
FILAS_POR_HOJA = 20
FILA_INICIAL = 15


# ==========================================
# MODELOS DE BASE DE DATOS
# ==========================================
class Registro(db.Model):
    __tablename__ = 'registros'
    co = db.Column(db.Integer, primary_key=True, autoincrement=True)
    seccion = db.Column(db.String, nullable=False)
    nombre_alumno = db.Column(db.String, nullable=False)
    tipo_apoyo = db.Column(db.String, nullable=False)
    area = db.Column(db.String, nullable=False)
    institucion = db.Column(db.String, nullable=False)
    carrera = db.Column(db.String, nullable=False)
    fecha_inicio = db.Column(db.String, nullable=False)
    fecha_fin = db.Column(db.String, nullable=False)
    estatus = db.Column(db.String, nullable=False)
    horas = db.Column(db.Integer, nullable=False)
    es_familiar = db.Column(db.String, nullable=False)
    nombre_familiar = db.Column(db.String)
    parentesco_familiar = db.Column(db.String)
    tipo_remuneracion = db.Column(db.String, nullable=False)
    banco_clabe = db.Column(db.String)
    numero_clabe = db.Column(db.String)
    archivo_carta_presentacion = db.Column(db.String)
    archivo_carta_aceptacion = db.Column(db.String)
    archivo_constancia = db.Column(db.String)
    archivo_buena_conducta = db.Column(db.String)
    archivo_ine = db.Column(db.String)
    archivo_reporte_mensual = db.Column(db.String)
    archivo_recibo_pago = db.Column(db.String)
    archivo_clabe = db.Column(db.String)
    archivo_contrato = db.Column(db.String)
    motivo_baja = db.Column(db.String)
    archivo_acta_nacimiento = db.Column(db.String)
    archivo_vigencia_imss = db.Column(db.String)
    archivo_formato_solicitud = db.Column(db.String)

    def __getitem__(self, index):
        columnas = [
            self.co, self.seccion, self.nombre_alumno, self.tipo_apoyo, self.area,
            self.institucion, self.carrera, self.fecha_inicio, self.fecha_fin, self.estatus,
            self.horas, self.es_familiar, self.nombre_familiar, self.parentesco_familiar,
            self.tipo_remuneracion, self.banco_clabe, self.numero_clabe,
            self.archivo_carta_presentacion, self.archivo_carta_aceptacion, self.archivo_constancia,
            self.archivo_buena_conducta, self.archivo_ine, self.archivo_reporte_mensual,
            self.archivo_recibo_pago, self.archivo_clabe, self.archivo_contrato,
            self.motivo_baja, self.archivo_acta_nacimiento, self.archivo_vigencia_imss,
            self.archivo_formato_solicitud
        ]
        return columnas[index]


class ReporteMensual(db.Model):
    __tablename__ = 'reportes_mensuales'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    alumno_co = db.Column(db.Integer, db.ForeignKey('registros.co', ondelete='CASCADE'))
    mes_anio = db.Column(db.String, nullable=False)
    nombre_archivo = db.Column(db.String, nullable=False)

    def __getitem__(self, index):
        return [self.id, self.alumno_co, self.mes_anio, self.nombre_archivo][index]


class ReciboAlumnoMensual(db.Model):
    __tablename__ = 'recibos_alumno_mensuales'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    alumno_co = db.Column(db.Integer, db.ForeignKey('registros.co', ondelete='CASCADE'))
    mes_anio = db.Column(db.String, nullable=False)
    nombre_archivo = db.Column(db.String, nullable=False)

    def __getitem__(self, index):
        return [self.id, self.alumno_co, self.mes_anio, self.nombre_archivo][index]


class ReciboPagoMensual(db.Model):
    __tablename__ = 'recibos_pago_mensuales'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    alumno_co = db.Column(db.Integer, db.ForeignKey('registros.co', ondelete='CASCADE'))
    mes_anio = db.Column(db.String, nullable=False)
    nombre_archivo = db.Column(db.String, nullable=False)

    def __getitem__(self, index):
        return [self.id, self.alumno_co, self.mes_anio, self.nombre_archivo][index]


def fecha_a_espanol(dt_obj):
    dia = dt_obj.strftime('%d')
    mes_en = dt_obj.strftime('%B')
    anio = dt_obj.strftime('%Y')
    mes_es = MESES_ESP.get(mes_en, mes_en)
    return f'{dia} de {mes_es} de {anio}'


def init_db():
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)


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


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/dashboard')
def dashboard():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    zona_filtro = request.args.get('zona_filtro')

    if session['rol'] == 'admin':
        if zona_filtro and zona_filtro != 'TODOS':
            registros = Registro.query.filter_by(seccion=zona_filtro).order_by(Registro.co.asc()).all()
        else:
            registros = Registro.query.order_by(Registro.co.asc()).all()
    else:
        registros = Registro.query.filter_by(seccion=session['usuario']).order_by(Registro.co.asc()).all()

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

        nuevo_registro = Registro(
            seccion=seccion,
            nombre_alumno=nombre_alumno,
            tipo_apoyo=tipo_apoyo,
            area=area,
            institucion=institucion,
            carrera=carrera,
            fecha_inicio=fecha_inicio_str,
            fecha_fin=fecha_fin_str,
            estatus=estatus,
            horas=horas,
            es_familiar=es_familiar,
            nombre_familiar=nombre_familiar,
            parentesco_familiar=parentesco_familiar,
            tipo_remuneracion=tipo_remuneracion
        )
        db.session.add(nuevo_registro)
        db.session.commit()

        flash('Estudiante registrado exitosamente.')
        return redirect(url_for('dashboard'))

    return render_template('registrar.html')


@app.route('/editar/<int:co>', methods=['GET', 'POST'])
def editar(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    if session['rol'] == 'admin':
        alumno = Registro.query.get(co)
    else:
        alumno = Registro.query.filter_by(co=co, seccion=session['usuario']).first()

    if not alumno:
        flash('Registro no encontrado o sin permisos.')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        alumno.nombre_alumno = request.form['nombre_alumno']
        alumno.tipo_apoyo = request.form['tipo_apoyo']
        alumno.area = request.form['area']
        alumno.institucion = request.form['institucion']
        alumno.carrera = request.form['carrera']
        alumno.fecha_inicio = request.form['fecha_inicio']
        alumno.fecha_fin = request.form['fecha_fin']
        alumno.horas = int(request.form['horas'])
        alumno.es_familiar = request.form.get('es_familiar', 'No')

        hoy = datetime.now().date()
        f_fin = datetime.strptime(alumno.fecha_fin, '%Y-%m-%d').date()
        alumno.estatus = 'Finalizado' if hoy > f_fin else 'Activo'

        if alumno.es_familiar == 'Sí':
            alumno.nombre_familiar = request.form.get('nombre_familiar', '')
            alumno.parentesco_familiar = request.form.get('parentesco_familiar', '')
            monto_apoyo = request.form.get('monto_apoyo', '2,580.86')
            alumno.tipo_remuneracion = f"Apoyo económico: ${monto_apoyo} pesos"
        else:
            alumno.nombre_familiar = 'N/A'
            alumno.parentesco_familiar = 'N/A'
            alumno.tipo_remuneracion = 'Título gratuito'

        db.session.commit()
        flash('Datos del alumno actualizados correctamente.')
        return redirect(url_for('dashboard'))

    return render_template('editar.html', alumno=alumno)


@app.route('/dar_de_baja/<int:co>', methods=['POST'])
def dar_de_baja(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    motivo = request.form.get('motivo_baja', '').strip()

    if session['rol'] == 'admin':
        alumno = Registro.query.get(co)
    else:
        alumno = Registro.query.filter_by(co=co, seccion=session['usuario']).first()

    if alumno:
        alumno.estatus = 'Baja obligada'
        alumno.motivo_baja = motivo
        db.session.commit()
        flash('Alumno dado de baja correctamente (Baja obligada).')

    return redirect(url_for('dashboard'))


@app.route('/subir_archivos/<int:co>', methods=['POST'])
def subir_archivos(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    banco = request.form.get('banco_clabe', '').strip()
    num_clabe = request.form.get('numero_clabe', '').strip()
    file_clabe = request.files.get('clabe')

    if file_clabe and file_clabe.filename != '' or banco or len(num_clabe) > 0:
        if not banco or len(num_clabe) != 18 or not num_clabe.isdigit():
            flash('Error: Para la CLABE Interbancaria es obligatorio indicar el banco y un número de exactamente 18 dígitos.')
            return redirect(url_for('dashboard'))

    alumno = Registro.query.get(co)
    if not alumno:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    if banco and num_clabe:
        alumno.banco_clabe = banco
        alumno.numero_clabe = num_clabe

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

    for input_name, db_col in campos_archivos.items():
        file = request.files.get(input_name)
        if file and file.filename != '':
            upload_result = cloudinary.uploader.upload(file.stream, resource_type="auto")
            file_url = upload_result.get('secure_url')
            setattr(alumno, db_col, file_url)

    db.session.commit()
    flash('Archivos subidos a la nube y datos bancarios actualizados correctamente.')
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
    alumno = Registro.query.get(co)

    if alumno:
        url_archivo = getattr(alumno, db_col)
        if url_archivo:
            if not url_archivo.startswith('http'):
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], url_archivo)
                if os.path.exists(filepath):
                    os.remove(filepath)

            setattr(alumno, db_col, None)
            if db_col == 'archivo_clabe':
                alumno.banco_clabe = None
                alumno.numero_clabe = None
            db.session.commit()
            flash('Archivo eliminado correctamente.')

    return redirect(url_for('dashboard'))


@app.route('/seleccionar_mes_orden/<int:co>')
def seleccionar_mes_orden(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    r = Registro.query.get(co)
    if not r:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    try:
        dt_actual = datetime.strptime(r.fecha_inicio, '%Y-%m-%d').replace(day=1)
        dt_final = datetime.strptime(r.fecha_fin, '%Y-%m-%d').replace(day=1)
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

    return render_template('seleccionar_mes.html', alumno=r, meses=meses_disponibles)


@app.route('/generar_orden_pago/<int:co>', methods=['POST'])
def generar_orden_pago(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))

    mes_seleccionado = request.form.get('mes_reportado', 'Septiembre 2026')
    r = Registro.query.get(co)

    if not r:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    nombre_alumno = r.nombre_alumno
    tipo_apoyo = r.tipo_apoyo.strip().lower()
    f_inicio_str = r.fecha_inicio
    f_fin_str = r.fecha_fin
    zona = r.seccion

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

    x_practicas = 'X' if 'práctica' in tipo_apoyo or 'practica' in tipo_apoyo else ''
    x_servicio = 'X' if 'servicio' in tipo_apoyo or ('práctica' not in tipo_apoyo and 'residencia' not in tipo_apoyo) else ''
    x_residencias = 'X' if 'residencia' in tipo_apoyo or 'estadía' in tipo_apoyo or 'estadia' in tipo_apoyo else ''

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

    output_filename = f'Recibo_Pago_{nombre_alumno.replace(" ", "_")}.docx'
    doc.save(output_filename)

    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(output_filename, arcname=output_filename)

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

    r = Registro.query.get(co)
    if not r:
        flash('Estudiante no encontrado.')
        return redirect(url_for('dashboard'))

    zona = r.seccion
    nombre_alumno = r.nombre_alumno
    tipo_apoyo = r.tipo_apoyo
    institucion = r.institucion
    carrera = r.carrera
    f_inicio_str = r.fecha_inicio
    f_fin_str = r.fecha_fin

    try:
        dt_inicio = datetime.strptime(f_inicio_str, '%Y-%m-%d')
        dt_fin = datetime.strptime(f_fin_str, '%Y-%m-%d')
        f_inicio_formato = f"{dt_inicio.strftime('%d')} DE {MESES_ESP.get(dt_inicio.strftime('%B'), dt_inicio.strftime('%B')).upper()} DEL {dt_inicio.strftime('%Y')}"
        f_fin_formato = f"{dt_fin.strftime('%d')} DE {MESES_ESP.get(dt_fin.strftime('%B'), dt_fin.strftime('%B')).upper()} DEL {dt_fin.strftime('%Y')}"
    except:
        f_inicio_formato = "02 DE FEBRERO DEL 2026"
        f_fin_formato = "02 DE AGOSTO DEL 2026"

    dt_hoy = datetime.now()
    f_actual_formato = f"{dt_hoy.strftime('%d')} DIAS DEL MES DE {MESES_ESP.get(dt_hoy.strftime('%B'), dt_hoy.strftime('%B')).upper()} DEL {dt_hoy.strftime('%Y')}"

    plantilla_convenio = 'CONVENIO A TITULO GRATUITO.docx'
    if not os.path.exists(plantilla_convenio):
        flash('Error: No se encontró el archivo "CONVENIO A TITULO GRATUITO.docx".')
        return redirect(url_for('dashboard'))

    doc = docx.Document(plantilla_convenio)
    reemplazos = {
        '[TIPO_DE_APOYO]': tipo_apoyo.upper(),
        '[NOMBRE_ALUMNO]': nombre_alumno,
        '[CARRERA]': carrera,
        '[INSTITUCION_EDUCATIVA]': institucion,
        '[FECHA_INICIO]': f_inicio_formato,
        '[FECHA_FIN]': f_fin_formato,
        '[FECHA_ACTUAL]': f_actual_formato,
        'Xalapa, Veracruz': zona,
    }

    def procesar_parrafo_obj(parrafo):
        texto_p = parrafo.text
        for clave, val in reemplazos.items():
            if clave in texto_p:
                texto_p = texto_p.replace(clave, val)
        if parrafo.runs:
            parrafo.runs[0].text = texto_p
            for run in parrafo.runs[1:]:
                run.text = ''

    for p in doc.paragraphs:
        procesar_parrafo_obj(p)

    output_filename = f'Convenio_Gratuito_{nombre_alumno.replace(" ", "_")}.docx'
    doc.save(output_filename)
    return send_file(output_filename, as_attachment=True, download_name=output_filename)


# ==========================================
# GESTIÓN DE OFICIOS DE PAGO POR ZONA
# ==========================================
@app.route('/gestionar_oficios_zona', methods=['GET'])
def gestionar_oficios_zona():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    usuario_actual = session['usuario']
    rol_usuario = session['rol']
    zona_seleccionada = request.args.get('zona_seleccionada', 'Oficinas Divisionales' if rol_usuario == 'admin' else usuario_actual)

    carpeta_zona = os.path.join('static', 'oficios_pago_zona', zona_seleccionada)
    os.makedirs(carpeta_zona, exist_ok=True)
    archivos_zona = os.listdir(carpeta_zona) if os.path.exists(carpeta_zona) else []
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
    file = request.files.get('archivo_oficio_zona')
    if file and file.filename != '':
        upload_result = cloudinary.uploader.upload(file.stream, resource_type="auto")
        flash('Oficio subido a la nube correctamente.')
    return redirect(url_for('gestionar_oficios_zona', zona_destino=zona_destino))


@app.route('/eliminar_oficio_zona/<string:zona>/<string:nombre_archivo>')
def eliminar_oficio_zona(zona, nombre_archivo):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    flash('Oficio eliminado.')
    return redirect(url_for('gestionar_oficios_zona', zona_seleccionada=zona))


# ==========================================
# GESTIÓN DE REPORTES Y RECIBOS MENSUALES
# ==========================================
@app.route('/gestionar_reportes/<int:co>', methods=['GET'])
def gestionar_reportes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    r = Registro.query.get(co)
    reportes_subidos = ReporteMensual.query.filter_by(alumno_co=co).all()
    return render_template('gestionar_reportes.html', alumno=r, reportes_subidos=reportes_subidos, meses_disponibles=[])


@app.route('/subir_reporte_mes/<int:co>', methods=['POST'])
def subir_reporte_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_reporte')
    if file and file.filename != '' and mes_anio:
        upload_result = cloudinary.uploader.upload(file.stream, resource_type="auto")
        file_url = upload_result.get('secure_url')

        existente = ReporteMensual.query.filter_by(alumno_co=co, mes_anio=mes_anio).first()
        if existente:
            existente.nombre_archivo = file_url
        else:
            nuevo = ReporteMensual(alumno_co=co, mes_anio=mes_anio, nombre_archivo=file_url)
            db.session.add(nuevo)
        db.session.commit()
        flash(f'Reporte de {mes_anio} subido a la nube correctamente.')
    return redirect(url_for('gestionar_reportes', co=co))


@app.route('/eliminar_reporte_mes/<int:reporte_id>/<int:co>')
def eliminar_reporte_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    row = ReporteMensual.query.get(reporte_id)
    if row:
        db.session.delete(row)
        db.session.commit()
        flash('Reporte mensual eliminado.')
    return redirect(url_for('gestionar_reportes', co=co))


@app.route('/gestionar_recibos_alumno/<int:co>', methods=['GET'])
def gestionar_recibos_alumno(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    r = Registro.query.get(co)
    recibos_subidos = ReciboAlumnoMensual.query.filter_by(alumno_co=co).all()
    return render_template('gestionar_recibos_alumno.html', alumno=r, recibos_subidos=recibos_subidos, meses_disponibles=[])


@app.route('/subir_recibo_alumno_mes/<int:co>', methods=['POST'])
def subir_recibo_alumno_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_recibo_alumno')
    if file and file.filename != '' and mes_anio:
        upload_result = cloudinary.uploader.upload(file.stream, resource_type="auto")
        file_url = upload_result.get('secure_url')

        existente = ReciboAlumnoMensual.query.filter_by(alumno_co=co, mes_anio=mes_anio).first()
        if existente:
            existente.nombre_archivo = file_url
        else:
            nuevo = ReciboAlumnoMensual(alumno_co=co, mes_anio=mes_anio, nombre_archivo=file_url)
            db.session.add(nuevo)
        db.session.commit()
        flash(f'Recibo del alumno de {mes_anio} subido a la nube correctamente.')
    return redirect(url_for('gestionar_recibos_alumno', co=co))


@app.route('/eliminar_recibo_alumno_mes/<int:reporte_id>/<int:co>')
def eliminar_recibo_alumno_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    row = ReciboAlumnoMensual.query.get(reporte_id)
    if row:
        db.session.delete(row)
        db.session.commit()
        flash('Recibo del alumno eliminado.')
    return redirect(url_for('gestionar_recibos_alumno', co=co))


@app.route('/gestionar_recibos_pago/<int:co>', methods=['GET'])
def gestionar_recibos_pago(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    r = Registro.query.get(co)
    recibos_subidos = ReciboPagoMensual.query.filter_by(alumno_co=co).all()
    return render_template('gestionar_recibos_pago.html', alumno=r, recibos_subidos=recibos_subidos, meses_disponibles=[])


@app.route('/subir_recibo_pago_mes/<int:co>', methods=['POST'])
def subir_recibo_pago_mes(co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    mes_anio = request.form.get('mes_anio')
    file = request.files.get('archivo_recibo')
    if file and file.filename != '' and mes_anio:
        upload_result = cloudinary.uploader.upload(file.stream, resource_type="auto")
        file_url = upload_result.get('secure_url')

        existente = ReciboPagoMensual.query.filter_by(alumno_co=co, mes_anio=mes_anio).first()
        if existente:
            existente.nombre_archivo = file_url
        else:
            nuevo = ReciboPagoMensual(alumno_co=co, mes_anio=mes_anio, nombre_archivo=file_url)
            db.session.add(nuevo)
        db.session.commit()
        flash(f'Oficio de pago de {mes_anio} subido a la nube correctamente.')
    return redirect(url_for('gestionar_recibos_pago', co=co))


@app.route('/eliminar_recibo_pago_mes/<int:reporte_id>/<int:co>')
def eliminar_recibo_pago_mes(reporte_id, co):
    if 'usuario' not in session:
        return redirect(url_for('login'))
    row = ReciboPagoMensual.query.get(reporte_id)
    if row:
        db.session.delete(row)
        db.session.commit()
        flash('Oficio de pago eliminado.')
    return redirect(url_for('gestionar_recibos_pago', co=co))


# ==========================================
# RUTAS DE EMERGENCIA
# ==========================================
@app.route('/crear_tablas_emergencia')
def crear_tablas_emergencia():
    db.create_all()
    return "¡Tablas creadas con éxito en la base de datos de Render!"


# ==========================================
# REPORTE EXCEL OFICIAL (BASADO EN PLANTILLA)
# ==========================================
def _fecha_corta(valor):
    try:
        return datetime.strptime(valor, '%Y-%m-%d').strftime('%d/%m/%Y')
    except Exception:
        return valor or ''


def _monto(remuneracion):
    m = re.search(r'\$\s*([\d,]+(?:\.\d+)?)', remuneracion or '')
    return f'${m.group(1)}' if m else 'Título gratuito'


def _restaurar_dibujos_plantilla(xlsx_bytes, archivo_plantilla=PLANTILLA_REPORTE):
    from xml.etree import ElementTree as ET

    if not os.path.exists(archivo_plantilla):
        xlsx_bytes.seek(0)
        return xlsx_bytes

    with zipfile.ZipFile(archivo_plantilla, "r") as z_tpl:
        plantilla_archivos = {n: z_tpl.read(n) for n in z_tpl.namelist()}

    with zipfile.ZipFile(xlsx_bytes, "r") as z_out:
        salida_archivos = {n: z_out.read(n) for n in z_out.namelist()}

    dibujo_original = plantilla_archivos.get("xl/drawings/drawing1.xml")
    relaciones_dibujo_originales = plantilla_archivos.get(
        "xl/drawings/_rels/drawing1.xml.rels"
    )
    if not dibujo_original or not relaciones_dibujo_originales:
        xlsx_bytes.seek(0)
        return xlsx_bytes

    NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
    NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

    ET.register_namespace("", NS_MAIN)
    ET.register_namespace("r", NS_REL_DOC)

    sheet_paths = sorted(
        [n for n in salida_archivos
         if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)],
        key=lambda n: int(re.search(r"sheet(\d+)\.xml$", n).group(1))
    )

    for nombre, contenido in plantilla_archivos.items():
        if nombre.startswith("xl/media/"):
            salida_archivos[nombre] = contenido

    for sheet_path in sheet_paths:
        numero_hoja = int(re.search(r"sheet(\d+)\.xml$", sheet_path).group(1))
        drawing_path = f"xl/drawings/drawing{numero_hoja}.xml"
        drawing_rels_path = f"xl/drawings/_rels/drawing{numero_hoja}.xml.rels"
        sheet_rels_path = f"xl/worksheets/_rels/sheet{numero_hoja}.xml.rels"

        salida_archivos[drawing_path] = dibujo_original
        salida_archivos[drawing_rels_path] = relaciones_dibujo_originales

        if sheet_rels_path in salida_archivos:
            rel_root = ET.fromstring(salida_archivos[sheet_rels_path])
        else:
            rel_root = ET.Element(f"{{{NS_REL_PKG}}}Relationships")

        dibujo_rel = None
        for rel in rel_root.findall(f"{{{NS_REL_PKG}}}Relationship"):
            if rel.attrib.get("Type", "").endswith("/drawing"):
                dibujo_rel = rel
                break

        if dibujo_rel is None:
            ids = []
            for rel in rel_root.findall(f"{{{NS_REL_PKG}}}Relationship"):
                match = re.fullmatch(r"rId(\d+)", rel.attrib.get("Id", ""))
                if match:
                    ids.append(int(match.group(1)))
            nuevo_id = f"rId{max(ids, default=0) + 1}"
            dibujo_rel = ET.SubElement(
                rel_root, f"{{{NS_REL_PKG}}}Relationship",
                {
                    "Id": nuevo_id,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                    "Target": f"../drawings/drawing{numero_hoja}.xml",
                }
            )
        else:
            dibujo_rel.set("Target", f"../drawings/drawing{numero_hoja}.xml")
            nuevo_id = dibujo_rel.attrib["Id"]

        salida_archivos[sheet_rels_path] = ET.tostring(
            rel_root, encoding="utf-8", xml_declaration=True
        )

        sheet_root = ET.fromstring(salida_archivos[sheet_path])
        ya_existe = sheet_root.find(f"{{{NS_MAIN}}}drawing")
        if ya_existe is None:
            drawing_element = ET.Element(
                f"{{{NS_MAIN}}}drawing",
                {f"{{{NS_REL_DOC}}}id": nuevo_id}
            )
            insertar_antes = None
            for child in list(sheet_root):
                if child.tag in (
                    f"{{{NS_MAIN}}}legacyDrawing",
                    f"{{{NS_MAIN}}}legacyDrawingHF",
                    f"{{{NS_MAIN}}}picture",
                    f"{{{NS_MAIN}}}oleObjects",
                    f"{{{NS_MAIN}}}controls",
                    f"{{{NS_MAIN}}}webPublishItems",
                    f"{{{NS_MAIN}}}tableParts",
                    f"{{{NS_MAIN}}}extLst",
                ):
                    insertar_antes = child
                    break
            if insertar_antes is None:
                sheet_root.append(drawing_element)
            else:
                sheet_root.insert(list(sheet_root).index(insertar_antes), drawing_element)
        else:
            ya_existe.set(f"{{{NS_REL_DOC}}}id", nuevo_id)

        salida_archivos[sheet_path] = ET.tostring(
            sheet_root, encoding="utf-8", xml_declaration=True
        )

    ct_path = "[Content_Types].xml"
    if ct_path in salida_archivos:
        ct_root = ET.fromstring(salida_archivos[ct_path])
        for sheet_path in sheet_paths:
            numero_hoja = int(re.search(r"sheet(\d+)\.xml$", sheet_path).group(1))
            part_name = f"/xl/drawings/drawing{numero_hoja}.xml"
            encontrado = any(
                item.tag == f"{{{NS_CT}}}Override"
                and item.attrib.get("PartName") == part_name
                for item in list(ct_root)
            )
            if not encontrado:
                ET.SubElement(
                    ct_root, f"{{{NS_CT}}}Override",
                    {
                        "PartName": part_name,
                        "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
                    }
                )
        salida_archivos[ct_path] = ET.tostring(
            ct_root, encoding="utf-8", xml_declaration=True
        )

    resultado = io.BytesIO()
    with zipfile.ZipFile(resultado, "w", zipfile.ZIP_DEFLATED) as z_final:
        for nombre, contenido in salida_archivos.items():
            z_final.writestr(nombre, contenido)

    resultado.seek(0)
    return resultado


def generar_reporte_desde_archivo(registros, archivo_plantilla):
    ahora = datetime.now()
    mes_es = MESES_ESP[ahora.strftime('%B')]
    mes_reportado = mes_es.upper()
    fecha_elab = f"Xalapa, Ver. a {ahora.day} de {mes_es} del {ahora.year}"

    bloques = {}
    for r in sorted(registros, key=lambda x: x.co):
        bloques.setdefault((r.co - 1) // FILAS_POR_HOJA, []).append(r)
    if not bloques:
        bloques = {0: []}
    claves = sorted(bloques)

    wb_final = openpyxl.load_workbook(archivo_plantilla)
    ws_plantilla = wb_final.active
    ws_plantilla.title = f"CO {claves[0] * FILAS_POR_HOJA + 1}-{(claves[0] + 1) * FILAS_POR_HOJA}"

    gris_encabezado = PatternFill(fill_type="solid", fgColor="FF747474")
    negro = Side(style="thin", color="FF000000")
    borde_celda = Border(left=negro, right=negro, top=negro, bottom=negro)

    for idx, k in enumerate(claves):
        ini = k * FILAS_POR_HOJA + 1
        fin = ini + FILAS_POR_HOJA - 1
        nombre_hoja = f"CO {ini}-{fin}"

        if idx == 0:
            ws = ws_plantilla
            ws.title = nombre_hoja
        else:
            ws = wb_final.copy_worksheet(ws_plantilla)
            ws.title = nombre_hoja

        for fila in range(FILA_INICIAL, FILA_INICIAL + FILAS_POR_HOJA):
            for col in range(3, 15):
                ws.cell(row=fila, column=col).value = None

        for i, r in enumerate(sorted(bloques[k], key=lambda x: x.co)):
            fila = FILA_INICIAL + i
            valores = [
                r.co,
                ahora.year,
                mes_reportado,
                r.seccion,
                r.area,
                r.tipo_apoyo,
                r.nombre_alumno,
                r.institucion,
                r.carrera,
                _fecha_corta(r.fecha_inicio),
                _fecha_corta(r.fecha_fin),
                _monto(r.tipo_remuneracion),
            ]
            for col_idx, valor in enumerate(valores, start=3):
                ws.cell(row=fila, column=col_idx).value = valor

        for col in range(3, 15):
            celda = ws.cell(row=14, column=col)
            celda.fill = copy.copy(gris_encabezado)
            celda.border = copy.copy(borde_celda)
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.font = Font(name="Arial", size=8, bold=True, color="FF000000")

        for fila in range(15, 35):
            for col in range(3, 15):
                celda = ws.cell(row=fila, column=col)
                celda.border = copy.copy(borde_celda)
                celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if celda.value is not None:
                    celda.font = Font(name="Arial", size=8, color="FF000000")

        ws.sheet_view.showGridLines = True
        ws["G36"] = fecha_elab

    salida = io.BytesIO()
    wb_final.save(salida)
    salida.seek(0)

    return _restaurar_dibujos_plantilla(salida, archivo_plantilla)


@app.route('/descargar_excel')
def descargar_excel():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    if not os.path.exists(PLANTILLA_REPORTE):
        flash(f'Error: No se encontró el archivo "{PLANTILLA_REPORTE}".')
        return redirect(url_for('dashboard'))

    if session['rol'] == 'admin':
        zona_filtro = request.args.get('zona_filtro')
        if zona_filtro and zona_filtro != 'TODOS':
            registros = Registro.query.filter_by(seccion=zona_filtro).order_by(Registro.co.asc()).all()
        else:
            registros = Registro.query.order_by(Registro.co.asc()).all()
    else:
        registros = Registro.query.filter_by(seccion=session['usuario']).order_by(Registro.co.asc()).all()

    output = generar_reporte_desde_archivo(registros, PLANTILLA_REPORTE)
    return send_file(
        output,
        as_attachment=True,
        download_name='Reporte_Oficial_CFE_Norma15.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/descargar_excel_todos')
def descargar_excel_todos():
    if 'usuario' not in session:
        return redirect(url_for('login'))

    if not os.path.exists(PLANTILLA_TODOS):
        flash(f'Error: No se encontró el archivo "{PLANTILLA_TODOS}".')
        return redirect(url_for('dashboard'))

    registros = Registro.query.order_by(Registro.co.asc()).all()
    output = generar_reporte_desde_archivo(registros, PLANTILLA_TODOS)
    return send_file(
        output,
        as_attachment=True,
        download_name='Historial_Completo_CFE_Norma15.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    init_db()
    app.run(debug=True)
    