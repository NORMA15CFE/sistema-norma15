import sqlite3
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from app import db, Registro, app

# Conexión directa con SQLite local por medio de sqlite3 puro
sqlite_path = 'datos.db'
conn_sqlite = sqlite3.connect(sqlite_path)
conn_sqlite.row_factory = sqlite3.Row
cursor_sqlite = conn_sqlite.cursor()

# Conexión a tu base de datos remota de Render (PostgreSQL)
postgres_uri = "postgresql://sistema_norma15_db_user:E7VgAmQxDdUeHHSl1SKY368ASgUhOoxw@dpg-dbd2on8473hc73alnbvg-a.oregon-postgres.render.com/sistema_norma15_db"
if postgres_uri.startswith("postgres://"):
    postgres_uri = postgres_uri.replace("postgres://", "postgresql+psycopg2://", 1)
elif postgres_uri.startswith("postgresql://"):
    postgres_uri = postgres_uri.replace("postgresql://", "postgresql+psycopg2://", 1)

engine_pg = create_engine(postgres_uri)
SessionPg = sessionmaker(bind=engine_pg)
session_pg = SessionPg()

with app.app_context():
    try:
        cursor_sqlite.execute("SELECT * FROM registros")
        filas = cursor_sqlite.fetchall()
        print(f"Se encontraron {len(filas)} registros en tu base de datos local.")
        
        for fila in filas:
            # Convertimos la fila de sqlite en un diccionario usando los nombres de las columnas
            row_dict = dict(fila)
            co_id = row_dict.get('co')
            
            # Verificamos si ya existe en PostgreSQL
            existe = session_pg.query(Registro).filter_by(co=co_id).first()
            if not existe:
                nuevo_r = Registro(
                    co=co_id,
                    seccion=row_dict.get('seccion', 'Xalapa'),
                    nombre_alumno=row_dict.get('nombre_alumno', 'Sin nombre'),
                    tipo_apoyo=row_dict.get('tipo_apoyo', 'Servicio Social'),
                    area=row_dict.get('area', 'General'),
                    institucion=row_dict.get('institucion', 'UV'),
                    carrera=row_dict.get('carrera', 'General'),
                    fecha_inicio=row_dict.get('fecha_inicio', '2026-01-01'),
                    fecha_fin=row_dict.get('fecha_fin', '2026-06-01'),
                    estatus=row_dict.get('estatus', 'Activo'),
                    horas=row_dict.get('horas', 480),
                    es_familiar=row_dict.get('es_familiar', 'No'),
                    nombre_familiar=row_dict.get('nombre_familiar', 'N/A'),
                    parentesco_familiar=row_dict.get('parentesco_familiar', 'N/A'),
                    tipo_remuneracion=row_dict.get('tipo_remuneracion', 'Título gratuito'),
                    banco_clabe=row_dict.get('banco_clabe'),
                    numero_clabe=row_dict.get('numero_clabe'),
                    motivo_baja=row_dict.get('motivo_baja')
                )
                session_pg.add(nuevo_r)
                print(f"Migrando alumno: {row_dict.get('nombre_alumno')}")
                
        session_pg.commit()
        print("¡Migración completada con éxito!")
    except Exception as e:
        print("Error durante la migración:", e)

conn_sqlite.close()
session_pg.close()
