# -*- coding: utf-8 -*-
"""
Módulo Principal - AeroTrack CAMAN
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime

from database import obtener_conexion, inicializar_bd, DB_NAME, EXCEL_FILE
from reports import generar_orden_trabajo, exportar_requerimiento_compras, exportar_todos_los_criticos

UI_COLORS = {
    "primary": "#0F2A4A",
    "secondary": "#1E5A99",
    "bg": "#F4F6F9",
    "white": "#FFFFFF",
    "text": "#333333",
    "gray": "#666666",
    "critical": "#8B0000",
    "critical_bg": "#FFEBEB",
    "critical_bar": "#A83232",
    "warning": "#806000",
    "warning_bg": "#FFF2CC",
    "success": "#2E7D32",
    "success_hover": "#388E3C",
    "replace": "#BF360C",
    "replace_hover": "#E64A19",
    "pdf": "#7B1FA2",
    "pdf_hover": "#9C27B0",
    "refresh": "#607D8B",
    "refresh_hover": "#78909C"
}

class AeroTrackApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AeroTrack-CAMAN - Control de Ciclos y Mantenimiento")
        self.root.geometry("1050x660")
        self.root.configure(bg=UI_COLORS["bg"])
        
        # Asegurar inicialización de la BD
        inicializar_bd()
        
        # Variables de control
        self.kpi_aeronaves = tk.StringVar(value="0")
        self.kpi_componentes = tk.StringVar(value="0")
        self.kpi_criticos = tk.StringVar(value="0")
        
        # Configurar Estilos ttk (Tema clam para soporte óptimo de tags en Treeview)
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.style.configure("TFrame", background=UI_COLORS["bg"])
        self.style.configure("TLabel", background=UI_COLORS["bg"], foreground=UI_COLORS["text"], font=("Segoe UI", 10))
        
        # Estilo del Treeview
        self.style.configure("Treeview", 
                            background=UI_COLORS["white"], 
                            foreground=UI_COLORS["text"], 
                            fieldbackground=UI_COLORS["white"], 
                            rowheight=26, 
                            font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", 
                            background=UI_COLORS["primary"], 
                            foreground=UI_COLORS["white"], 
                            font=("Segoe UI", 10, "bold"), 
                            relief="flat")
        
        # Corrección del mapa de estilos en Windows para permitir tags de colores
        def fixed_map(option):
            return [elm for elm in self.style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        self.style.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))
        self.style.map("Treeview.Heading", background=[('active', '#1E5A99')])
        
        # Crear componentes visuales
        self.crear_interfaz()
        
        # Cargar datos iniciales
        self.cargar_datos_tabla()
        self.actualizar_kpis()
        
    def crear_interfaz(self):
        # 1. Header Banner Institucional
        header = tk.Frame(self.root, bg=UI_COLORS["primary"], height=75)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        lbl_titulo = tk.Label(header, text="AeroTrack - CAMAN ✈️", font=("Segoe UI", 16, "bold"), fg=UI_COLORS["white"], bg=UI_COLORS["primary"])
        lbl_titulo.pack(anchor="w", padx=20, pady=(10, 0))
        
        lbl_sub = tk.Label(header, text="Control de Horas de Vuelo y Alertas de Componentes Críticos — FUERZA AÉREA COLOMBIANA", 
                           font=("Segoe UI", 9, "italic"), fg="#A9C7DF", bg=UI_COLORS["primary"])
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 10))
        
        # Contenedor Principal
        main_frame = tk.Frame(self.root, bg=UI_COLORS["bg"], padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)
        
        # 2. Panel de Indicadores (KPIs)
        kpi_frame = tk.Frame(main_frame, bg=UI_COLORS["bg"])
        kpi_frame.pack(fill="x", pady=(0, 15))
        
        # Configurar anchos y rejilla de KPIs
        kpi_frame.columnconfigure(0, weight=1)
        kpi_frame.columnconfigure(1, weight=1)
        kpi_frame.columnconfigure(2, weight=1)
        
        self.kpi_aeronaves_card = self.crear_kpi_card(kpi_frame, "Aeronaves Registradas", self.kpi_aeronaves, UI_COLORS["primary"])
        self.kpi_aeronaves_card.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        self.kpi_componentes_card = self.crear_kpi_card(kpi_frame, "Componentes Totales", self.kpi_componentes, UI_COLORS["secondary"])
        self.kpi_componentes_card.grid(row=0, column=1, padx=5, sticky="ew")
        
        # KPI Críticos
        self.kpi_criticos_card = tk.Frame(kpi_frame, bg=UI_COLORS["white"], bd=1, relief="solid", highlightthickness=0)
        self.kpi_criticos_card.grid(row=0, column=2, padx=(10, 0), sticky="ew")
        
        self.kpi_criticos_bar = tk.Frame(self.kpi_criticos_card, bg=UI_COLORS["success"], width=5)
        self.kpi_criticos_bar.pack(side="left", fill="y")
        
        crit_content = tk.Frame(self.kpi_criticos_card, bg=UI_COLORS["white"], padx=15, pady=10)
        crit_content.pack(side="left", fill="both", expand=True)
        
        lbl_crit_title = tk.Label(crit_content, text="Componentes Críticos (>= 90%)", font=("Segoe UI", 9, "bold"), fg=UI_COLORS["gray"], bg=UI_COLORS["white"], anchor="w")
        lbl_crit_title.pack(fill="x")
        
        self.lbl_crit_val = tk.Label(crit_content, textvariable=self.kpi_criticos, font=("Segoe UI", 18, "bold"), fg=UI_COLORS["success"], bg=UI_COLORS["white"], anchor="w")
        self.lbl_crit_val.pack(fill="x")
        
        # 3. Layout de Tabla y Botones Acciones
        content_layout = tk.Frame(main_frame, bg=UI_COLORS["bg"])
        content_layout.pack(fill="both", expand=True)
        
        content_layout.columnconfigure(0, weight=4)
        content_layout.columnconfigure(1, weight=1)
        content_layout.rowconfigure(0, weight=1)
        
        # Contenedor Tabla
        table_container = tk.Frame(content_layout, bg=UI_COLORS["white"], bd=1, relief="solid")
        table_container.grid(row=0, column=0, sticky="nsew", padx=(0, 15))
        
        # Título del Listado
        tbl_header = tk.Frame(table_container, bg=UI_COLORS["primary"], pady=6)
        tbl_header.pack(fill="x")
        lbl_tbl_title = tk.Label(tbl_header, text="  Monitoreo del Ciclo de Vida de Componentes", font=("Segoe UI", 10, "bold"), fg=UI_COLORS["white"], bg=UI_COLORS["primary"])
        lbl_tbl_title.pack(side="left")
        
        # Tabla Treeview
        self.tree = ttk.Treeview(table_container, columns=("id_comp", "nombre", "id_avion", "horas_act", "limite", "desgaste"), show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Definir Encabezados de Columnas
        self.tree.heading("id_comp", text="ID Componente", anchor="center")
        self.tree.heading("nombre", text="Nombre del Componente", anchor="w")
        self.tree.heading("id_avion", text="ID Avión", anchor="center")
        self.tree.heading("horas_act", text="Horas Actuales", anchor="center")
        self.tree.heading("limite", text="Límite Horas", anchor="center")
        self.tree.heading("desgaste", text="Desgaste (%)", anchor="center")
        
        self.tree.column("id_comp", width=120, anchor="center")
        self.tree.column("nombre", width=260, anchor="w")
        self.tree.column("id_avion", width=100, anchor="center")
        self.tree.column("horas_act", width=110, anchor="center")
        self.tree.column("limite", width=110, anchor="center")
        self.tree.column("desgaste", width=115, anchor="center")
        
        # Configurar colores para estados de advertencia
        self.tree.tag_configure('critical', background='#FFEBEB', foreground='#8B0000') # >= 90%
        self.tree.tag_configure('warning', background='#FFF2CC', foreground='#806000')  # 80-89%
        self.tree.tag_configure('normal', background='#FFFFFF', foreground='#333333')
        
        # Contenedor Botones de Acción
        actions_panel = tk.LabelFrame(content_layout, text=" Acciones Rápidas ", bg=UI_COLORS["bg"], font=("Segoe UI", 10, "bold"), fg="#0f2a4a", labelanchor="n", padx=10, pady=15)
        actions_panel.grid(row=0, column=1, sticky="nsew")
        
        btn_vuelo = self.crear_boton_moderno(actions_panel, "✈️ Registrar Vuelo", self.abrir_registrar_vuelo, bg=UI_COLORS["secondary"], hover_bg="#2C75C4")
        btn_vuelo.pack(fill="x", pady=8)
      
        btn_instalar = self.crear_boton_moderno(actions_panel, "🔧 Instalar Componente", self.abrir_instalar_componente, bg=UI_COLORS["primary"], hover_bg=UI_COLORS["secondary"])
        btn_instalar.pack(fill="x", pady=8)
      
        btn_reemplazar = self.crear_boton_moderno(actions_panel, "♻️ Reemplazar Componente", self.abrir_reemplazar_componente, bg=UI_COLORS["replace"], hover_bg=UI_COLORS["replace_hover"])
        btn_reemplazar.pack(fill="x", pady=8)
        
        btn_exportar = self.crear_boton_moderno(actions_panel, "📥 Exportar a Compras", self.exportar_compras_click, bg=UI_COLORS["success"], hover_bg=UI_COLORS["success_hover"])
        btn_exportar.pack(fill="x", pady=8)
        
        btn_pdf = self.crear_boton_moderno(actions_panel, "📄 Generar PDF OT", self.generar_pdf_seleccionado, bg=UI_COLORS["pdf"], hover_bg=UI_COLORS["pdf_hover"])
        btn_pdf.pack(fill="x", pady=8)
        
        btn_refrescar = self.crear_boton_moderno(actions_panel, "🔄 Actualizar Datos", self.refrescar_todo, bg=UI_COLORS["refresh"], hover_bg=UI_COLORS["refresh_hover"])
        btn_refrescar.pack(fill="x", pady=8)
        
        # Separador decorativo
        sep = ttk.Separator(actions_panel, orient="horizontal")
        sep.pack(fill="x", pady=15)
        
        # Leyenda explicativa en el panel lateral
        leyenda_frame = tk.Frame(actions_panel, bg=UI_COLORS["bg"])
        leyenda_frame.pack(fill="both", expand=True)
        
        lbl_leyenda_title = tk.Label(leyenda_frame, text="Leyenda de Alertas:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg="#555555")
        lbl_leyenda_title.pack(anchor="w", pady=(0, 5))
        
        item_crit = tk.Frame(leyenda_frame, bg=UI_COLORS["critical_bg"], bd=1, relief="solid")
        item_crit.pack(fill="x", pady=2)
        tk.Label(item_crit, text="Crítico (>= 90% desgaste)", font=("Segoe UI", 8), bg=UI_COLORS["critical_bg"], fg=UI_COLORS["critical"]).pack(padx=5, pady=2)
        
        item_warn = tk.Frame(leyenda_frame, bg=UI_COLORS["warning_bg"], bd=1, relief="solid")
        item_warn.pack(fill="x", pady=2)
        tk.Label(item_warn, text="Advertencia (80-89% desgaste)", font=("Segoe UI", 8), bg=UI_COLORS["warning_bg"], fg=UI_COLORS["warning"]).pack(padx=5, pady=2)
        
        item_ok = tk.Frame(leyenda_frame, bg=UI_COLORS["white"], bd=1, relief="solid")
        item_ok.pack(fill="x", pady=2)
        tk.Label(item_ok, text="Normal (< 80% desgaste)", font=("Segoe UI", 8), bg=UI_COLORS["white"], fg=UI_COLORS["text"]).pack(padx=5, pady=2)
        
        # 4. Status Bar inferior
        self.status_bar = tk.Label(self.root, text=" Listo | Base de datos conectada correctamente.", bd=1, relief="sunken", anchor="w", bg="#E0E4EC", font=("Segoe UI", 8), fg="#555555")
        self.status_bar.pack(side="bottom", fill="x")
        
    def crear_kpi_card(self, parent, title, value_var, color):
        card = tk.Frame(parent, bg=UI_COLORS["white"], bd=1, relief="solid", highlightthickness=0)
        
        bar = tk.Frame(card, bg=color, width=5)
        bar.pack(side="left", fill="y")
        
        content = tk.Frame(card, bg=UI_COLORS["white"], padx=15, pady=10)
        content.pack(side="left", fill="both", expand=True)
        
        lbl_title = tk.Label(content, text=title, font=("Segoe UI", 9, "bold"), fg=UI_COLORS["gray"], bg=UI_COLORS["white"], anchor="w")
        lbl_title.pack(fill="x")
        
        lbl_val = tk.Label(content, textvariable=value_var, font=("Segoe UI", 18, "bold"), fg=UI_COLORS["text"], bg=UI_COLORS["white"], anchor="w")
        lbl_val.pack(fill="x")
        
        return card
        
    def crear_boton_moderno(self, parent, text, command, bg=UI_COLORS["secondary"], fg=UI_COLORS["white"], hover_bg="#2C75C4"):
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=12,
            pady=8,
            activebackground=hover_bg,
            activeforeground=UI_COLORS["white"]
        )
        def on_enter(e):
            btn.config(background=hover_bg)
        def on_leave(e):
            btn.config(background=bg)
            
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn
        
    def centrar_ventana(self, ventana, ancho, alto):
        self.root.update_idletasks()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        
        x = rx + (rw - ancho) // 2
        y = ry + (rh - alto) // 2
        ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
        
    def mostrar_estado(self, texto):
        self.status_bar.config(text=f" {texto}")
        
    def cargar_datos_tabla(self):
        # Limpiar filas existentes
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            cursor.execute("SELECT id_componente, nombre, id_avion, horas_actuales, limite_horas FROM Componentes;")
            componentes = cursor.fetchall()
            conn.close()
            
            for comp in componentes:
                id_comp, nombre, id_avion, horas_act, limite = comp
                desgaste = (horas_act / limite) * 100 if limite > 0 else 0.0
                desgaste_str = f"{desgaste:.1f}%"
                
                # Asignar tag según el desgaste
                if desgaste >= 90.0:
                    tag = 'critical'
                elif desgaste >= 80.0:
                    tag = 'warning'
                else:
                    tag = 'normal'
                    
                self.tree.insert("", "end", values=(id_comp, nombre, id_avion, f"{horas_act:.1f}", f"{limite:.1f}", desgaste_str), tags=(tag,))
                
            self.mostrar_estado(f"Datos de componentes cargados. Total: {len(componentes)} registros.")
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"Error al cargar datos desde sqlite3: {str(e)}")
            
    def actualizar_kpis(self):
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            # Cantidad aeronaves
            cursor.execute("SELECT COUNT(*) FROM Aeronaves;")
            count_aviones = cursor.fetchone()[0]
            
            # Cantidad componentes
            cursor.execute("SELECT COUNT(*) FROM Componentes;")
            count_comp = cursor.fetchone()[0]
            
            # Componentes críticos (wear >= 90%)
            cursor.execute("SELECT COUNT(*) FROM Componentes WHERE horas_actuales >= (limite_horas * 0.9);")
            count_criticos = cursor.fetchone()[0]
            
            conn.close()
            
            self.kpi_aeronaves.set(str(count_aviones))
            self.kpi_componentes.set(str(count_comp))
            self.kpi_criticos.set(str(count_criticos))
            
            # Colorear dinámicamente el KPI de críticos si es mayor a cero
            if count_criticos > 0:
                self.lbl_crit_val.config(fg=UI_COLORS["critical"])
                self.kpi_criticos_bar.config(bg=UI_COLORS["critical_bar"])
            else:
                self.lbl_crit_val.config(fg=UI_COLORS["success"])
                self.kpi_criticos_bar.config(bg=UI_COLORS["success"])
                
        except sqlite3.Error as e:
            self.mostrar_estado(f"Error al calcular KPIs: {str(e)}")
            
    def refrescar_todo(self):
        self.cargar_datos_tabla()
        self.actualizar_kpis()
        self.mostrar_estado("Datos refrescados con éxito.")
        
    def _obtener_lista_aeronaves(self):
        aeronaves_list = []
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("SELECT id_avion FROM Aeronaves ORDER BY id_avion;")
            aeronaves_list = [r[0] for r in cursor.fetchall()]
            conn.close()
        except sqlite3.Error as e:
            self.mostrar_estado(f"Error cargando aeronaves: {str(e)}")
        return aeronaves_list

    def abrir_registrar_vuelo(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Registrar Ciclo de Vuelo - AeroTrack")
        dialog.configure(bg=UI_COLORS["bg"])
        dialog.resizable(False, False)
        self.centrar_ventana(dialog, 420, 240)
        dialog.grab_set()
        
        # Encabezado del diálogo
        hdr = tk.Frame(dialog, bg=UI_COLORS["primary"], pady=8)
        hdr.pack(fill="x")
        lbl_hdr = tk.Label(hdr, text="Registrar Vuelo de Aeronave ✈️", font=("Segoe UI", 11, "bold"), fg="white", bg=UI_COLORS["primary"])
        lbl_hdr.pack()
        
        # Formulario
        form = tk.Frame(dialog, bg=UI_COLORS["bg"], padx=20, pady=20)
        form.pack(fill="both", expand=True)
        
        tk.Label(form, text="Aeronave ID:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=0, column=0, sticky="w", pady=6)
        
        # Cargar lista de aeronaves de la base de datos
        aeronaves_list = self._obtener_lista_aeronaves()
            
        combo_avion = ttk.Combobox(form, values=aeronaves_list, font=("Segoe UI", 10), state="readonly")
        combo_avion.grid(row=0, column=1, sticky="ew", pady=6, padx=(10, 0))
        if aeronaves_list:
            combo_avion.current(0)
            
        tk.Label(form, text="Horas de Vuelo:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=1, column=0, sticky="w", pady=6)
        entry_horas = ttk.Entry(form, font=("Segoe UI", 10))
        entry_horas.grid(row=1, column=1, sticky="ew", pady=6, padx=(10, 0))
        entry_horas.focus()
        
        form.columnconfigure(1, weight=1)
        
        # Contenedor Botones
        btn_frame = tk.Frame(dialog, bg=UI_COLORS["bg"], pady=10)
        btn_frame.pack(fill="x")
        
        def confirmar():
            id_avion = combo_avion.get()
            horas_str = entry_horas.get().strip()
            self.ejecutar_registro_vuelo(id_avion, horas_str, dialog)
            
        btn_cancel = self.crear_boton_moderno(btn_frame, "Cancelar", dialog.destroy, bg="#888888", hover_bg="#999999")
        btn_cancel.pack(side="right", padx=15)
        
        btn_ok = self.crear_boton_moderno(btn_frame, "Registrar", confirmar, bg=UI_COLORS["secondary"], hover_bg="#2C75C4")
        btn_ok.pack(side="right")
        
    def ejecutar_registro_vuelo(self, id_avion, horas_str, dialog):
        if not id_avion:
            messagebox.showerror("Error de Entrada", "Debe seleccionar un avión existente.")
            return
            
        try:
            horas_vuelo = float(horas_str)
            if horas_vuelo <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Error de Entrada", "Las horas de vuelo deben ser un número decimal mayor que cero.")
            return
            
        conn = obtener_conexion()
        conn.execute("PRAGMA foreign_keys = ON;")
        
        try:
            # Transacción única por lote
            with conn:
                cursor = conn.cursor()
                
                # Obtener detalles aeronave
                cursor.execute("SELECT modelo, horas_totales_planeador FROM Aeronaves WHERE id_avion = ?;", (id_avion,))
                aeronave = cursor.fetchone()
                if not aeronave:
                    messagebox.showerror("Error", "La aeronave seleccionada no existe en la base de datos.")
                    return
                    
                modelo_avion = aeronave[0]
                horas_antiguas_planeador = aeronave[1]
                nuevas_horas_planeador = horas_antiguas_planeador + horas_vuelo
                
                # 1. Actualizar Planeador en Aeronaves
                cursor.execute("""
                    UPDATE Aeronaves 
                    SET horas_totales_planeador = horas_totales_planeador + ? 
                    WHERE id_avion = ?;
                """, (horas_vuelo, id_avion))
                
                # 2. Actualizar Horas de los Componentes
                cursor.execute("""
                    UPDATE Componentes 
                    SET horas_actuales = horas_actuales + ? 
                    WHERE id_avion = ?;
                """, (horas_vuelo, id_avion))
                
                # 3. Consultar componentes del avión para verificar límites >= 90%
                cursor.execute("""
                    SELECT id_componente, nombre, horas_actuales, limite_horas 
                    FROM Componentes 
                    WHERE id_avion = ?;
                """, (id_avion,))
                
                componentes = cursor.fetchall()
                alertas_gatilladas = []
                errores_documentos = []
                
                fecha_ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for comp in componentes:
                    id_comp, nombre_comp, horas_act, lim_horas = comp
                    porcentaje_uso = (horas_act / lim_horas) * 100 if lim_horas > 0 else 0.0
                    
                    if horas_act >= (lim_horas * 0.9):
                        # Insertar auditoría de Mantenimiento Preventivo en la base de datos
                        cursor.execute("""
                            INSERT INTO Mantenimiento (id_componente, tipo_inspeccion, fecha, tecnico_encargado) 
                            VALUES (?, ?, ?, ?);
                        """, (id_comp, "Alerta Preventiva Automatizada (>= 90%)", fecha_ahora, "AeroTrack Logic Engine"))
                        
                        datos_alerta = {
                            "id_componente": id_comp,
                            "nombre": nombre_comp,
                            "id_avion": id_avion,
                            "modelo_avion": modelo_avion,
                            "horas_totales_planeador": nuevas_horas_planeador,
                            "horas_actuales": horas_act,
                            "limite_horas": lim_horas
                        }
                        
                        # Generar el PDF
                        pdf_exito, pdf_err = generar_orden_trabajo(datos_alerta)
                        if not pdf_exito:
                            errores_documentos.append(f"PDF [{id_comp}]: {pdf_err}")
                            
                        # Registrar en Excel de Compras
                        excel_exito, excel_err = exportar_requerimiento_compras(datos_alerta)
                        if not excel_exito:
                            errores_documentos.append(f"Excel [{id_comp}]: {excel_err}")
                            
                        alertas_gatilladas.append((id_comp, nombre_comp, porcentaje_uso))
            
            # Cierre y actualización GUI
            dialog.destroy()
            self.cargar_datos_tabla()
            self.actualizar_kpis()
            
            # Mostrar reportes al usuario
            if alertas_gatilladas:
                msg_base = "¡ALERTA DE MANTENIMIENTO PREVENTIVO CRÍTICO!\n\n"
                msg_base += "Los siguientes componentes han superado el límite del 90%:\n"
                for cid, cnom, cporc in alertas_gatilladas:
                    msg_base += f"• {cid} - {cnom} ({cporc:.1f}%)\n"
                
                msg_base += "\n- Se han generado los PDFs de Orden de Trabajo (OT_[ID].pdf).\n"
                msg_base += f"- Se han registrado los requerimientos en '{EXCEL_FILE}'.\n"
                
                if errores_documentos:
                    msg_base += "\nADVERTENCIAS ADICIONALES (Documentos con errores de bloqueo):\n"
                    for error_doc in errores_documentos:
                        msg_base += f"• {error_doc}\n"
                        
                messagebox.showwarning("Componentes Críticos Detectados", msg_base)
            else:
                messagebox.showinfo("Registro Exitoso", f"El ciclo de vuelo de +{horas_vuelo}h se registró correctamente para {id_avion}.")
                
        except sqlite3.Error as e:
            messagebox.showerror("Error en Transacción", f"No se pudo completar el registro del vuelo: {str(e)}")
            
    def abrir_instalar_componente(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Instalar Nuevo Componente - AeroTrack")
        dialog.configure(bg=UI_COLORS["bg"])
        dialog.resizable(False, False)
        self.centrar_ventana(dialog, 460, 360)
        dialog.grab_set()
        
        # Encabezado
        hdr = tk.Frame(dialog, bg=UI_COLORS["primary"], pady=8)
        hdr.pack(fill="x")
        lbl_hdr = tk.Label(hdr, text="Formulario de Instalación de Repuesto 🔧", font=("Segoe UI", 11, "bold"), fg="white", bg=UI_COLORS["primary"])
        lbl_hdr.pack()
        
        # Formulario
        form = tk.Frame(dialog, bg=UI_COLORS["bg"], padx=20, pady=15)
        form.pack(fill="both", expand=True)
        
        # ID Componente
        tk.Label(form, text="ID Componente:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=0, column=0, sticky="w", pady=5)
        entry_id = ttk.Entry(form, font=("Segoe UI", 10))
        entry_id.grid(row=0, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Nombre Componente
        tk.Label(form, text="Nombre Componente:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=1, column=0, sticky="w", pady=5)
        entry_nombre = ttk.Entry(form, font=("Segoe UI", 10))
        entry_nombre.grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Aeronave ID
        tk.Label(form, text="Asignado a Avión:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=2, column=0, sticky="w", pady=5)
        
        aeronaves_list = self._obtener_lista_aeronaves()
            
        combo_avion = ttk.Combobox(form, values=aeronaves_list, font=("Segoe UI", 10), state="readonly")
        combo_avion.grid(row=2, column=1, sticky="ew", pady=5, padx=(10, 0))
        if aeronaves_list:
            combo_avion.current(0)
            
        # Horas Actuales
        tk.Label(form, text="Horas Iniciales:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=3, column=0, sticky="w", pady=5)
        entry_horas = ttk.Entry(form, font=("Segoe UI", 10))
        entry_horas.insert(0, "0.0")
        entry_horas.grid(row=3, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Límite de Horas
        tk.Label(form, text="Límite Horas Operativas:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=4, column=0, sticky="w", pady=5)
        entry_limite = ttk.Entry(form, font=("Segoe UI", 10))
        entry_limite.insert(0, "1000.0")
        entry_limite.grid(row=4, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        form.columnconfigure(1, weight=1)
        
        # Botones
        btn_frame = tk.Frame(dialog, bg=UI_COLORS["bg"], pady=10)
        btn_frame.pack(fill="x")
        
        def confirmar():
            id_comp = entry_id.get().strip()
            nombre = entry_nombre.get().strip()
            id_avion = combo_avion.get()
            horas_str = entry_horas.get().strip()
            limite_str = entry_limite.get().strip()
            self.ejecutar_instalacion_componente(id_comp, nombre, id_avion, horas_str, limite_str, dialog)
            
        btn_cancel = self.crear_boton_moderno(btn_frame, "Cancelar", dialog.destroy, bg="#888888", hover_bg="#999999")
        btn_cancel.pack(side="right", padx=15)
        
        btn_ok = self.crear_boton_moderno(btn_frame, "Instalar", confirmar, bg=UI_COLORS["secondary"], hover_bg="#2C75C4")
        btn_ok.pack(side="right")
        
    def ejecutar_instalacion_componente(self, id_comp, nombre, id_avion, horas_str, limite_str, dialog):
        if not id_comp or not nombre or not id_avion:
            messagebox.showerror("Error de Entrada", "Todos los campos de texto son requeridos.")
            return
            
        try:
            horas_act = float(horas_str)
            limite_horas = float(limite_str)
            if horas_act < 0 or limite_horas <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Error de Entrada", "Las horas iniciales deben ser >= 0 y el límite operativo debe ser > 0.")
            return
            
        conn = obtener_conexion()
        conn.execute("PRAGMA foreign_keys = ON;")
        cursor = conn.cursor()
        
        try:
            # Comprobar duplicado
            cursor.execute("SELECT COUNT(*) FROM Componentes WHERE id_componente = ?;", (id_comp,))
            if cursor.fetchone()[0] > 0:
                messagebox.showerror("Error de Integridad", f"El ID de componente '{id_comp}' ya se encuentra registrado.")
                conn.close()
                return
                
            # Guardar el componente en la BD
            cursor.execute("""
                INSERT INTO Componentes (id_componente, nombre, id_avion, horas_actuales, limite_horas)
                VALUES (?, ?, ?, ?, ?);
            """, (id_comp, nombre, id_avion, horas_act, limite_horas))
            
            # Registrar instalación en el log de mantenimiento histórico
            fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("""
                INSERT INTO Mantenimiento (id_componente, tipo_inspeccion, fecha, tecnico_encargado)
                VALUES (?, ?, ?, ?);
            """, (id_comp, "Instalación de Nuevo Componente", fecha_hoy, "Tgo. de Mantenimiento CAMAN"))
            
            conn.commit()
            conn.close()
            
            dialog.destroy()
            self.cargar_datos_tabla()
            self.actualizar_kpis()
            messagebox.showinfo("Registro Exitoso", f"El componente '{id_comp}' se ha asignado e instalado en la aeronave '{id_avion}'.")
            
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo guardar el componente en la base de datos: {str(e)}")

    def abrir_reemplazar_componente(self):
        """Abre la ventana emergente para reemplazar un componente existente (reset a 0 horas)."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Reemplazar Componente - AeroTrack")
        dialog.configure(bg=UI_COLORS["bg"])
        dialog.resizable(False, False)
        self.centrar_ventana(dialog, 480, 340)
        dialog.grab_set()

        # Encabezado
        hdr = tk.Frame(dialog, bg=UI_COLORS["replace"], pady=8)
        hdr.pack(fill="x")
        lbl_hdr = tk.Label(hdr, text="Reemplazo de Componente / Overhaul ♻️", font=("Segoe UI", 11, "bold"), fg="white", bg=UI_COLORS["replace"])
        lbl_hdr.pack()

        # Formulario
        form = tk.Frame(dialog, bg=UI_COLORS["bg"], padx=20, pady=15)
        form.pack(fill="both", expand=True)

        # Cargar lista de IDs de componentes desde la BD
        componentes_ids = []
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("SELECT id_componente FROM Componentes ORDER BY id_componente;")
            componentes_ids = [r[0] for r in cursor.fetchall()]
            conn.close()
        except sqlite3.Error as e:
            self.mostrar_estado(f"Error cargando componentes: {str(e)}")

        # ID Componente (Combobox)
        tk.Label(form, text="ID Componente a Reemplazar:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=0, column=0, sticky="w", pady=6)
        combo_componente = ttk.Combobox(form, values=componentes_ids, font=("Segoe UI", 10), state="readonly")
        combo_componente.grid(row=0, column=1, sticky="ew", pady=6, padx=(10, 0))

        # Pre-seleccionar el componente si el usuario ya lo tenía seleccionado en el Treeview principal
        seleccion_previa = self.tree.selection()
        if seleccion_previa:
            item = self.tree.item(seleccion_previa[0])
            id_preseleccionado = str(item["values"][0])
            if id_preseleccionado in componentes_ids:
                combo_componente.set(id_preseleccionado)
        elif componentes_ids:
            combo_componente.current(0)

        # Nuevo Límite de Horas (opcional)
        tk.Label(form, text="Nuevo Límite de Horas:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=1, column=0, sticky="w", pady=6)
        entry_nuevo_limite = ttk.Entry(form, font=("Segoe UI", 10))
        entry_nuevo_limite.grid(row=1, column=1, sticky="ew", pady=6, padx=(10, 0))

        lbl_limite_hint = tk.Label(form, text="(Dejar vacío para mantener el límite actual)", font=("Segoe UI", 8, "italic"), bg=UI_COLORS["bg"], fg="#888888")
        lbl_limite_hint.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(0, 4))

        # Técnico Encargado
        tk.Label(form, text="Técnico Encargado:", font=("Segoe UI", 9, "bold"), bg=UI_COLORS["bg"], fg=UI_COLORS["text"]).grid(row=3, column=0, sticky="w", pady=6)
        entry_tecnico = ttk.Entry(form, font=("Segoe UI", 10))
        entry_tecnico.insert(0, "Tgo. de Mantenimiento CAMAN")
        entry_tecnico.grid(row=3, column=1, sticky="ew", pady=6, padx=(10, 0))

        form.columnconfigure(1, weight=1)

        # Botones
        btn_frame = tk.Frame(dialog, bg=UI_COLORS["bg"], pady=10)
        btn_frame.pack(fill="x")

        def confirmar():
            id_comp = combo_componente.get()
            nuevo_limite_str = entry_nuevo_limite.get().strip()
            tecnico = entry_tecnico.get().strip()
            self.ejecutar_reemplazo_componente(id_comp, nuevo_limite_str, tecnico, dialog)

        btn_cancel = self.crear_boton_moderno(btn_frame, "Cancelar", dialog.destroy, bg="#888888", hover_bg="#999999")
        btn_cancel.pack(side="right", padx=15)

        btn_ok = self.crear_boton_moderno(btn_frame, "Confirmar Reemplazo", confirmar, bg=UI_COLORS["replace"], hover_bg=UI_COLORS["replace_hover"])
        btn_ok.pack(side="right")

    def ejecutar_reemplazo_componente(self, id_comp, nuevo_limite_str, tecnico, dialog):
        """
        Ejecuta la lógica de reemplazo de un componente:
        - UPDATE horas_actuales = 0.0 (y opcionalmente limite_horas)
        - INSERT en Mantenimiento registrando la acción de reemplazo.
        REGLA DE SEGURIDAD: No se borran datos existentes. Operación envuelta en try-except.
        """
        if not id_comp:
            messagebox.showerror("Error de Entrada", "Debe seleccionar un componente válido para reemplazar.")
            return

        if not tecnico:
            messagebox.showerror("Error de Entrada", "El nombre del Técnico Encargado es obligatorio.")
            return

        # Validar nuevo límite si fue proporcionado
        nuevo_limite = None
        if nuevo_limite_str:
            try:
                nuevo_limite = float(nuevo_limite_str)
                if nuevo_limite <= 0:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Error de Entrada", "El nuevo límite de horas debe ser un número decimal mayor que cero.")
                return

        conn = obtener_conexion()
        conn.execute("PRAGMA foreign_keys = ON;")

        try:
            with conn:
                cursor = conn.cursor()

                # Verificar que el componente existe
                cursor.execute("SELECT id_componente, nombre, id_avion FROM Componentes WHERE id_componente = ?;", (id_comp,))
                comp_row = cursor.fetchone()
                if not comp_row:
                    messagebox.showerror("Error", f"El componente '{id_comp}' no existe en la base de datos.")
                    return

                nombre_comp = comp_row[1]

                # 1. UPDATE: Reset horas_actuales a 0.0 y (opcionalmente) actualizar limite_horas
                if nuevo_limite is not None:
                    cursor.execute("""
                        UPDATE Componentes
                        SET horas_actuales = 0.0, limite_horas = ?
                        WHERE id_componente = ?;
                    """, (nuevo_limite, id_comp))
                else:
                    cursor.execute("""
                        UPDATE Componentes
                        SET horas_actuales = 0.0
                        WHERE id_componente = ?;
                    """, (id_comp,))

                # 2. INSERT: Registrar acción en la tabla Mantenimiento
                fecha_ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    INSERT INTO Mantenimiento (id_componente, tipo_inspeccion, fecha, tecnico_encargado)
                    VALUES (?, ?, ?, ?);
                """, (id_comp, "Reemplazo de Componente / Overhaul", fecha_ahora, tecnico))

            # Transacción exitosa - actualizar GUI
            dialog.destroy()
            self.cargar_datos_tabla()
            self.actualizar_kpis()

            msg_exito = (
                f"El componente '{id_comp}' ({nombre_comp}) ha sido reemplazado exitosamente.\n\n"
                f"• Horas actuales reiniciadas a 0.0\n"
            )
            if nuevo_limite is not None:
                msg_exito += f"• Nuevo límite de horas: {nuevo_limite:.1f}\n"
            msg_exito += f"• Técnico: {tecnico}\n"
            msg_exito += f"• Fecha: {fecha_ahora}\n"
            msg_exito += f"\nRegistro guardado en el historial de Mantenimiento."

            messagebox.showinfo("Reemplazo Exitoso", msg_exito)
            self.mostrar_estado(f"Componente '{id_comp}' reemplazado con éxito. Horas reiniciadas a 0.0.")

        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo completar el reemplazo del componente: {str(e)}")
            
    def exportar_compras_click(self):
        # Desencadenar reporte de Excel
        cant, error = exportar_todos_los_criticos()
        
        if error:
            messagebox.showerror("Error de Exportación", error)
        elif cant == 0:
            messagebox.showinfo("Exportar a Compras", "No existen componentes en estado crítico (con desgaste >= 90%) en la flota para exportar.")
        else:
            messagebox.showinfo("Exportación Exitosa", f"Se ha generado el archivo '{EXCEL_FILE}' con {cant} requerimiento(s) de compra en estado crítico.")

    def generar_pdf_seleccionado(self):
        """Genera el PDF de Orden de Trabajo para el componente seleccionado en la tabla."""
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin Selección", "Seleccione un componente de la tabla antes de generar el PDF.")
            return
            
        item = self.tree.item(seleccion[0])
        valores = item["values"]
        id_comp = str(valores[0])
        desgaste_str = str(valores[5]).replace("%", "")
        
        try:
            desgaste = float(desgaste_str)
        except ValueError:
            messagebox.showerror("Error", "No se pudo leer el porcentaje de desgaste del componente seleccionado.")
            return
            
        if desgaste < 90.0:
            messagebox.showinfo("Componente No Crítico",
                                f"El componente '{id_comp}' tiene un desgaste de {desgaste:.1f}%.\n"
                                f"Solo se genera la Orden de Trabajo para componentes con desgaste >= 90%.")
            return
            
        # Consultar datos completos del componente y su aeronave
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT c.id_componente, c.nombre, c.id_avion, c.horas_actuales, c.limite_horas,
                       a.modelo, a.horas_totales_planeador
                FROM Componentes c
                JOIN Aeronaves a ON c.id_avion = a.id_avion
                WHERE c.id_componente = ?;
            """, (id_comp,))
            row = cursor.fetchone()
            conn.close()
            
            if not row:
                messagebox.showerror("Error", f"No se encontró el componente '{id_comp}' en la base de datos.")
                return
                
            datos = {
                "id_componente": row[0],
                "nombre": row[1],
                "id_avion": row[2],
                "horas_actuales": row[3],
                "limite_horas": row[4],
                "modelo_avion": row[5],
                "horas_totales_planeador": row[6]
            }
            
            exito, error = generar_orden_trabajo(datos)
            
            if exito:
                pdf_nombre = f"OT_{id_comp}.pdf"
                messagebox.showinfo("PDF Generado",
                                    f"La Orden de Trabajo se ha generado exitosamente:\n\n"
                                    f"📄 {pdf_nombre}\n\n"
                                    f"El archivo se encuentra en la carpeta del proyecto.")
                self.mostrar_estado(f"PDF '{pdf_nombre}' generado con éxito.")
            else:
                messagebox.showerror("Error al Generar PDF", error)
                
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo consultar el componente: {str(e)}")


if __name__ == "__main__":
    # Arrancar la aplicación visual
    root = tk.Tk()
    app = AeroTrackApp(root)
    root.mainloop()
