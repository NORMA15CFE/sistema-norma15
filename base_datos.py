import sqlite3
import os

def crear_base_datos():
    conexion = sqlite3.connect('datos.db')
    cursor = conexion.cursor()
    
    # Tabla principal de registros con la estructura correcta de CFE Norma 15
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
            motivo_baja TEXT
        )
    ''')

    # Tablas auxiliares para reportes y recibos mensuales
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

    conexion.commit()
    conexion.close()
    print("¡Base de datos y tablas de CFE Norma 15 sincronizadas exitosamente!")

if __name__ == '__main__':
    crear_base_datos()
    