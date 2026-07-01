# AeroTrack-CAMAN ✈️

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![SQLite](https://img.shields.io/badge/database-SQLite3-green)
![ReportLab](https://img.shields.io/badge/PDF-ReportLab-red)
![OpenPyXL](https://img.shields.io/badge/Excel-OpenPyXL-brightgreen)
![Status](https://img.shields.io/badge/status-Production_Ready-success)

**Sistema Local de Control de Ciclos de Vuelo y Alertas de Mantenimiento**  
*Comando Aéreo de Mantenimiento (CAMAN) - Fuerza Aeroespacial Colombiana*

---

## 📖 Resumen Ejecutivo

**AeroTrack-CAMAN** es una solución de software de escritorio de arquitectura robusta, diseñada específicamente para auditar, gestionar y controlar el ciclo de vida de componentes aeronáuticos críticos. 

El sistema reemplaza los controles manuales mediante un "Logic Engine" que propaga automáticamente el desgaste de vuelo desde el planeador de la aeronave hacia todos sus repuestos instalados. Al llegar a umbrales de seguridad predefinidos, el sistema desencadena alertas visuales y documentales, ordenando el reemplazo preventivo antes de que ocurra una falla técnica en operación.

---

## 🚀 Características Principales

### 1. Motor de Control de Ciclos (Logic Engine)
El registro de un solo vuelo para una aeronave actualiza automáticamente y en cascada el desgaste (horas actuales) de todos y cada uno de los componentes vinculados a dicha matrícula. Este proceso se realiza dentro de transacciones SQL atómicas para evitar corrupción de datos.

### 2. Sistema de Alertas de Seguridad Tricolor
El sistema evalúa el estado del componente basado en la fórmula `(horas_actuales / limite_horas) * 100`:
- 🟢 **Normal (< 80%)**: Operación segura.
- 🟡 **Advertencia (80% - 89%)**: El componente entra en etapa de observación logística.
- 🔴 **Crítico (>= 90%)**: El componente ha llegado al límite de seguridad operacional. Se emite una alerta inmediata, gatillando la generación de documentos oficiales y bloqueando tácitamente su vuelo.

### 3. Automatización Documental (ReportLab y OpenPyXL)
El sistema elimina el papeleo manual:
- **Órdenes de Trabajo PDF**: Al alcanzar un estado Crítico, se genera un formato PDF oficial con el membrete institucional, detallando el nivel de desgaste y ordenando la inspección/reemplazo del componente.
- **Módulo de Compras (Excel)**: El área logística cuenta con una consolidación automatizada en Excel (`requerimientos_compras.xlsx`) con estilos profesionales que agrupa la necesidad de componentes a adquirir.

---

## 🏗️ Arquitectura Modular del Sistema

El monolito inicial ha sido refactorizado bajo principios **SOLID** y **DRY (Don't Repeat Yourself)**, implementando una clara **Separación de Conceptos (Separation of Concerns)**.

### 📂 `main.py` (Capa de Presentación y Lógica de Negocio)
Punto de entrada de la aplicación. Maneja la Interfaz Gráfica de Usuario (GUI) construida con la librería estándar `tkinter` y `ttk`.
- Centraliza la paleta de colores institucionales (`UI_COLORS`).
- Orquesta las interacciones del usuario mediante ventanas de diálogo para registrar vuelos, instalar repuestos y reemplazarlos.
- Renderiza el dashboard (KPIs) y la tabla de monitoreo en tiempo real.

### 📂 `database.py` (Capa de Acceso a Datos)
Fábrica de persistencia y abstracción de la base de datos `aerotrack_caman.db`.
- Implementa el patrón Factory mediante `obtener_conexion()`, garantizando que toda consulta aplique `PRAGMA foreign_keys = ON`.
- Define el modelo de datos relacional y las restricciones `ON DELETE CASCADE`.

### 📂 `reports.py` (Capa de Automatización)
Módulo especializado en la interacción ofimática.
- Encapsula la lógica de **ReportLab** para el dibujo vectorial y estructurado de la Orden de Trabajo PDF.
- Encapsula la lógica de **OpenPyXL** para el formateo celular, auto-ajuste de columnas y colores del reporte logístico.

---

## 🗄️ Esquema de Base de Datos (SQLite)

La base de datos local `aerotrack_caman.db` consta de 3 tablas relacionales fuertemente tipadas:

1. **Aeronaves**: 
   - `id_avion` (PK), `modelo`, `horas_totales_planeador`.
2. **Componentes**: 
   - `id_componente` (PK), `nombre`, `id_avion` (FK), `horas_actuales`, `limite_horas`.
3. **Mantenimiento**: 
   - `id_registro` (PK Auto), `id_componente` (FK), `tipo_inspeccion`, `fecha`, `tecnico_encargado`.
   - *Actúa como el Log Histórico Inmutable de todas las acciones operativas.*

---

## 🛠️ Instalación y Despliegue

### Requisitos Previos
- **Python 3.8** o superior instalado en el equipo (`python --version`).
- Sistema operativo Windows (Recomendado para la GUI de Tkinter), aunque es multiplataforma.

### Pasos de Instalación

1. **Clonar/Extraer el Repositorio** en la máquina local.
2. **Crear un Entorno Virtual (Opcional pero recomendado)**:
   ```bash
   python -m venv venv
   # En Windows:
   venv\Scripts\activate
   ```
3. **Instalar Dependencias de Generación de Reportes**:
   El sistema base funciona sin internet y sin dependencias, pero para generar Excel y PDF necesitas:
   ```bash
   pip install reportlab openpyxl
   ```
4. **Ejecutar la Aplicación**:
   ```bash
   python main.py
   ```

---

## 📘 Guía de Uso del Usuario

### Panel de KPIs (Dashboard)
Al iniciar, visualizarás en tiempo real:
- **Aeronaves Registradas**: Total de aviones en la flota.
- **Componentes Totales**: Repuestos bajo monitoreo.
- **Componentes Críticos**: Repuestos que exceden el 90%. (Este panel se torna color **Rojo** si el valor es mayor a 0).

### Botones de Acción (Acciones Rápidas)

- ✈️ **Registrar Vuelo**: Suma horas de uso a un avión. Inmediatamente el sistema auditará si algún componente de este avión sobrepasó el límite.
- 🔧 **Instalar Componente**: Permite matricular un componente completamente nuevo (0 horas) asociándolo a un avión de la base de datos.
- ♻️ **Reemplazar Componente**: Realiza un *Overhaul*. Se selecciona el ID de un componente averiado y se regresan sus horas a `0.0`. Guarda un registro en el log de mantenimiento indicando el técnico encargado.
- 📥 **Exportar a Compras**: Analiza toda la base de datos, extrae todos los repuestos con desgaste >= 90% y sobreescribe el archivo `requerimientos_compras.xlsx` en el directorio actual.
- 📄 **Generar PDF OT**: Requiere que selecciones un registro en la tabla visual principal. Si el desgaste es Crítico, emite la Orden de Trabajo manual en PDF.
- 🔄 **Actualizar Datos**: Refresca la tabla central desde la base de datos (Útil si se trabaja en red compartida).

---

## 🛡️ Políticas de Seguridad e Integridad de Datos

Este sistema está diseñado para entornos de ingeniería aeroespacial. Por lo tanto, cumple con las siguientes directrices:
- **No-Destrucción**: La aplicación carece de botones de borrado manual. Los historiales no pueden eliminarse desde la interfaz para asegurar auditorías limpias.
- **Sandboxing de Errores**: Si el archivo Excel/PDF está abierto en otro programa, la aplicación atrapará el `PermissionError` alertando al usuario en lugar de cerrar el sistema (Crash prevention).
- **Consistencia SQL**: El uso de `WITH conn:` en las actualizaciones críticas asegura transacciones completas (Commit) o nulas (Rollback), imposibilitando los estados parciales corruptos ante apagones.

---
*Desarrollado para la optimización de los procesos de mantenimiento preventivo y el aseguramiento del alistamiento operativo.*
