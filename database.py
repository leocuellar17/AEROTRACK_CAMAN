# -*- coding: utf-8 -*-
"""
Módulo de Base de Datos - AeroTrack CAMAN
Maneja la conexión a SQLite, inicialización de tablas e integridad referencial.
"""

import sqlite3

# Constantes del Sistema
DB_NAME = "aerotrack_caman.db"
EXCEL_FILE = "requerimientos_compras.xlsx"

def obtener_conexion():
    """
    Patrón Factory para obtener la conexión a la base de datos.
    Garantiza la integridad referencial habilitando las llaves foráneas.
    """
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def inicializar_bd():
    """
    Inicializa la base de datos SQLite con la estructura requerida.
    Garantiza la integridad referencial habilitando las llaves foráneas.
    REGLA DE SEGURIDAD: NO se inserta ningún dato de prueba para respetar el 100% de la información real.
    """
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Crear Tabla de Aeronaves
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Aeronaves (
        id_avion TEXT PRIMARY KEY,
        modelo TEXT NOT NULL,
        horas_totales_planeador REAL NOT NULL
    );
    """)
    
    # Crear Tabla de Componentes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Componentes (
        id_componente TEXT PRIMARY KEY,
        nombre TEXT NOT NULL,
        id_avion TEXT NOT NULL,
        horas_actuales REAL NOT NULL,
        limite_horas REAL NOT NULL,
        FOREIGN KEY (id_avion) REFERENCES Aeronaves(id_avion) ON DELETE CASCADE
    );
    """)
    
    # Crear Tabla de Mantenimiento
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Mantenimiento (
        id_registro INTEGER PRIMARY KEY AUTOINCREMENT,
        id_componente TEXT NOT NULL,
        tipo_inspeccion TEXT NOT NULL,
        fecha TEXT NOT NULL,
        tecnico_encargado TEXT NOT NULL,
        FOREIGN KEY (id_componente) REFERENCES Componentes(id_componente) ON DELETE CASCADE
    );
    """)
    
    conn.commit()
    conn.close()
