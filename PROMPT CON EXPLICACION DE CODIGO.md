Actua como experto en flet, highs y en python en general. Te voy a dar el codigo que he hecho para automatizar el planning de staff de eventos de mi organizacion. Funciona dándole (en la primera pestaña) una lista de personas, una lista de tareas, seleccionando las horas en las que se va a suceder el evento, asignando los pesos de la función objetivo, determinando el tiempo limite de ejecución. En la segunda pestaña, se seleccionan la disponibilidad de las personas, los requerimientos de tareas (cantidad de tareas por hora) y las habilitaciones de cada persona en cada tarea (si la pueden realizar o no). En la tercera pestaña se seleccionan las obligatoriedades, es decir, se selecciona si cierta persona debe trabajar en cierta tarea a cierta hora.
Una vez seleccionado todo eso (toda esta info se guarda en un .json), se le da al boton de ejecutar y usando un modelo de programacion lineal con la libreria PuLP, que usa el motor CBC, se obtiene un resultado óptimo o pseudo optimo (si el time limit se alcanza antes).

Quiero que me digas qué otras funcionalidades se le podrian añadir para mejorar el programa:


```
import flet as ft # Librería para el UI
from pulp import * # Librería para programación lineal, motor CBC por defecto
import json # Librería de python para leer .json
import os
import threading  # Para ejecutar el cálculo en segundo plano sin congelar la UI
import openpyxl   # Para generar el reporte en Excel
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import platform   # Para detectar el sistema operativo (abrir el excel automáticamente)
import subprocess
import highspy # Librería que permite ejecutar el algoritmo HiGHS una vez generado el archivo .mps con PuLP

# Archivo donde se guardará la persistencia de datos (JSON)
DATA_FILE = "staffing_data.json"

# =============================================================================
# FUNCIONES DE DATOS Y MODELO MATEMÁTICO
# =============================================================================

def load_data():
    # Carga los datos desde el archivo JSON si existe. Retorna None si no.
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def save_data(data):
    # Guarda el diccionario de datos actual en el archivo JSON.
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def solve_model(data):

    # Tomamos los datos guardados en el .json de la ejecución previa
    people = data['people']
    tasks = data['tasks']
    hours = data['hours']
    D = data['D']
    Q = data['Q']
    R = data['R']
    F = data['F']
     
    alpha = float(data['alpha'])
    beta = float(data['beta'])
    gamma = float(data['gamma'])
    epsilon = float(data['epsilon'])
    timelimit = int(data['timelimit'])
    solver_type = data.get('solver', 'highs') # 'cbc' o 'highs'
    
    # RESOLVEMOS EL MODELO
    print(f"--- INICIANDO CONSTRUCCIÓN DEL MODELO (Solver: {solver_type.upper()}) ---")
    model = LpProblem("Staffing", LpMinimize)

    # 1. CAPA DE SEGURIDAD
    # Convertimos los nombres de personas y de tareas en etiquetas genéricas (sin tildes ni símbolos raros)
    # De esta manera, no habrá problemas para leer el archivo .mps
    safe_people = [f"P{i}" for i in range(len(people))]
    safe_tasks =  [f"T{i}" for i in range(len(tasks))]
    
    # 2. VARIABLES DE DECISIÓN (con etiquetas genéricas)
    X_safe = LpVariable.dicts("X", (safe_people, safe_tasks, hours), cat='Binary')
    W_safe = LpVariable.dicts("W", safe_people, lowBound=0, cat='Integer')
    W_max = LpVariable("W_max", lowBound=0)
    W_min = LpVariable("W_min", lowBound=0)
    
    hours_minus_last = hours[:-1]
    Y_safe = LpVariable.dicts("Y", (safe_people, safe_tasks, hours_minus_last), cat='Binary')
    hours_minus_first = hours[1:]
    S_safe = LpVariable.dicts("S", (safe_people, hours_minus_first), cat='Binary')
    U_safe = LpVariable.dicts("U", (safe_people, safe_tasks, hours), cat='Binary')

    # 3. RECONSTRUCCIÓN DE DICCIONARIOS (Mapeo seguro con etiquetas genéricas -> Mapeo real)
    X = {people[i]: {tasks[j]: {h: X_safe[safe_people[i]][safe_tasks[j]][h] for h in hours} for j in range(len(tasks))} for i in range(len(people))}
    W = {people[i]: W_safe[safe_people[i]] for i in range(len(people))}
    Y = {people[i]: {tasks[j]: {h: Y_safe[safe_people[i]][safe_tasks[j]][h] for h in hours_minus_last} for j in range(len(tasks))} for i in range(len(people))}
    S = {people[i]: {h: S_safe[safe_people[i]][h] for h in hours_minus_first} for i in range(len(people))}
    U = {people[i]: {tasks[j]: {h: U_safe[safe_people[i]][safe_tasks[j]][h] for h in hours} for j in range(len(tasks))} for i in range(len(people))}
    
    # 4. FUNCIÓN OBJETIVO
    model += (
        alpha * (W_max - W_min) +
        beta * lpSum(Y[i][t][h] for i in people for t in tasks for h in hours_minus_last) +
        gamma * lpSum(S[i][h] for i in people for h in hours_minus_first) +
        epsilon * lpSum(U[i][t][h] for i in people for t in tasks for h in hours)
    )
    
    # 5. RESTRICCIONES

    # Una persona no debe hacer más de una tarea en una hora dada
    for i in people:
        for h in hours:
            model += lpSum(X[i][t][h] for t in tasks) <= D[i][h]

    # Todas las tareas de la matriz de requerimientos R deben ser satisfechas
    for t in tasks:
        for h in hours:
            model += lpSum(X[i][t][h] for i in people) == R[t][h]

    # Una persona no debe realizar más tareas a lo largo del día de lo que la matriz de disponibilidad Q dice
    for i in people:
        for t in tasks:
            for h in hours:
                model += X[i][t][h] <= Q[i][t]

    # Nadie deberá hacer más horas que el máximo ni menos horas que el mínimo establecido por el modelo
    for i in people:
        model += W[i] == lpSum(X[i][t][h] for t in tasks for h in hours)
        model += W_max >= W[i]
        model += W_min <= W[i]

    # (Restricción soft) En la medida de lo posible, se intentará que las personas no hagan dos tareas iguales en horas consecutivas
    # Es decir, se evitará la monotonía
    for i in people:
        for t in tasks:
            for h in hours_minus_last:
                h_next = hours[hours.index(h) + 1]
                model += Y[i][t][h] >= X[i][t][h] + X[i][t][h_next] - 1

    # (Restricción soft) En la medida de lo posible, se intentará que no haya descansos intermedios entre tarea
    # Es decir, se intentará que la gente trabaje todas sus horas de continuo
    for i in people:
        for h in hours_minus_first:
            h_prev = hours[hours.index(h) - 1]
            T_ih = lpSum(X[i][t][h] for t in tasks)
            T_ih_prev = lpSum(X[i][t][h_prev] for t in tasks)
            model += S[i][h] >= T_ih - T_ih_prev

    # (Restricción soft) En la medida de lo posible, se obligará a las personas a respetar la matriz de obligatoriedad F
    # Es decir, que si indicamos que la persona i debe trabajar en la tarea t en la hora h, deberá cumplirse
    for i in people:
        for t in tasks:
            for h in hours:
                model += U[i][t][h] >= F[i][t][h] - X[i][t][h]
    
    # =========================================================
    # LÓGICA DE SELECCIÓN DE MOTOR
    # =========================================================
    
    # --- OPCIÓN A: CBC (por defecto en PuLP) ---
    if solver_type == 'cbc':
        print(f"Ejecutando CBC (PuLP default) - TimeLimit: {timelimit}s...")
        try:
            # CBC rellena automáticamente las variables X_safe, W_safe, etc.
            # Como X, W apuntan a ellas, no hace falta inyección manual.
            model.solve(PULP_CBC_CMD(msg=1, timeLimit=timelimit))
        except Exception as e:
            print(f"Error CBC: {e}")
            model.status = LpStatusInfeasible

    # --- OPCIÓN B: HIGHS (MPS -> HIGHSPY) ---
    else:
        # Busca si se ha creado un archivo .mps
        mps_file = "temp_staffing_model.mps"
        if os.path.exists(mps_file): os.remove(mps_file)
        
        print(f"Exportando modelo a {mps_file}...")
        model.writeMPS(mps_file)
        print(f"Ejecutando Highs (native highspy)...")
        
        try:
            h = highspy.Highs()
            h.setOptionValue("time_limit", float(timelimit))
            h.setOptionValue("output_flag", True) 
            h.setOptionValue("presolve", "on")
            
            h.readModel(mps_file)
            h.run()
            
            status_h = h.getModelStatus()
            info = h.getInfo()
            
            print(f"Highs Code: {status_h}")
            has_feasible_sol = (info.primal_solution_status == 2)
            
            if status_h == highspy.HighsModelStatus.kOptimal:
                model.status = LpStatusOptimal
            elif status_h == highspy.HighsModelStatus.kTimeLimit:
                if has_feasible_sol:
                    model.status = LpStatusOptimal
                else:
                    model.status = LpStatusNotSolved
            elif status_h == highspy.HighsModelStatus.kInfeasible:
                model.status = LpStatusInfeasible
            else:
                model.status = LpStatusUndefined

            # Inyección Manual para Highs
            if has_feasible_sol:
                solution = h.getSolution()
                col_vals = solution.col_value
                num_cols = h.getNumCol()
                val_map = {}
                for k in range(num_cols):
                    ret = h.getColName(k)
                    if isinstance(ret, tuple): _, col_name = ret 
                    else: col_name = ret
                    val_map[col_name] = round(col_vals[k])
                
                for v in model.variables():
                    if v.name in val_map:
                        v.varValue = val_map[v.name]
                    else:
                        v.varValue = 0 
            
        except Exception as e:
            print(f"ERROR CRÍTICO HIGHS: {e}")
            model.status = LpStatusInfeasible
        finally:
            if os.path.exists(mps_file):
                try: os.remove(mps_file)
                except: pass

    return model, X, W, W_max, W_min

# =============================================================================
# APLICACIÓN PRINCIPAL (FLET)
# =============================================================================

class StaffingApp:
    def __init__(self):
        # Carga inicial de datos
        self.data = load_data()
        
        # Lista maestra de horas posibles (desde las 16:00 hasta las 08:00 del día siguiente)
        self.possible_hours = [16, 17, 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8]
        
        # --- COLORES Y ESTILOS (CONSTANTES) ---
        self.COLOR_ACTIVE = "#C6EFCE"    # Verde Excel claro
        self.TEXT_ACTIVE = "#006100"     # Verde texto oscuro
        self.COLOR_INACTIVE = "#FFC7CE"  # Rojo Excel claro
        self.TEXT_INACTIVE = "#9C0006"   # Rojo texto oscuro
        self.COLOR_HEADER_BG = "#F2F2F2"
        self.COLOR_BORDER = "#bfbfbf" 
        self.COLOR_TEXT_HIGHLIGHT = "blue700" 
        self.COLOR_BG_PANEL = "white"        
        self.COLOR_NEUTRAL = "grey200"

        # Paleta de colores rotativa para las tareas
        self.available_colors = ["blue200", "red200", "green200", "amber200", "purple200", "cyan200", "orange200", "pink200", "teal200", "indigo200", "lime200", "brown200"]
        self.task_colors = {} # Se llenará dinámicamente: { "Barra": "blue200", ... }

        # Mapeo de colores de Flet a Hexadecimal para Excel
        self.FLET_TO_HEX = {
            "blue200": "90CAF9", "red200": "EF9A9A", "green200": "A5D6A7", 
            "amber200": "FFE082", "purple200": "CE93D8", "cyan200": "80DEEA", 
            "orange200": "FFCC80", "pink200": "F48FB1", "teal200": "80CBC4", 
            "indigo200": "9FA8DA", "lime200": "E6EE9C", "brown200": "BCAAA4",
            "white": "FFFFFF", "red100": "FFCDD2" 
        }

        # --- ESTADO EN MEMORIA (State Management) ---
        # Estos diccionarios guardan la configuración actual de la UI antes de guardar en JSON
        self.state_hours = {} 
        self.state_D = {} # Disponibilidad
        self.state_Q = {} # Habilidades
        self.state_R = {} # Requerimientos numéricos
        self.state_F = {} # Fijas
        
        self.r_cells = {} # Referencias a los TextFields de la matriz R (para navegación con teclado)
        self.grid_controls = {'D': {}, 'Q': {}} # Referencias a los botones de celda para actualizarlos rápido
        self.bulk_states = {} 

        self.input_people_val = ""
        self.input_tasks_val = ""
        
        # Inicializar horas activas desde el JSON cargado
        for i in range(17):
            active = 0
            if self.data and 'hours' in self.data and i in self.data['hours']:
                active = 1
            self.state_hours[i] = active

        # Controles globales de UI
        self.status_text = ft.Text("Ready.", color="grey600", size=12)
        self.progress_bar = ft.ProgressBar(width=200, color="blue", bgcolor="#eeeeee", visible=False)
        self.btn_optimize = None
        
        # Contenedores principales (Placeholders)
        self.content_matrices = ft.Column(spacing=20) 
        self.content_mandatory = ft.Column(spacing=20) 
        self.container_hours = ft.Column()
        self.page = None

    def main(self, page: ft.Page):
        # Punto de entrada de la aplicación Flet
        self.page = page
        self.page.title = "Staffing Optimizer"
        self.page.window.width = 1200
        self.page.window.height = 900
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 10

        # Cargar valores por defecto en los campos de texto
        def_pers = '\n'.join(self.data.get('people', [])) if self.data else ""
        def_task = '\n'.join(self.data.get('tasks', [])) if self.data else ""
        self.input_people_val = def_pers
        self.input_tasks_val = def_task

        # Campos de entrada de texto multilínea
        self.txt_people = ft.TextField(
            value=def_pers, multiline=True, expand=True, text_size=12, border=ft.InputBorder.NONE,
            on_blur=self.on_input_change # Se activa al salir del foco para regenerar tablas
        )
        self.txt_tasks = ft.TextField(
            value=def_task, multiline=True, expand=True, text_size=12, border=ft.InputBorder.NONE,
            on_blur=self.on_input_change
        )
        
        self.generate_hour_buttons()

        # Helper para obtener parámetros con valor por defecto
        def get_p(k, d): return self.data.get(k, d) if self.data else d
        
        # Helper visual para crear inputs de parámetros (Alpha, Beta, etc.)
        def input_param(label, value):
            return ft.Row([
                ft.Container(
                    content=ft.Text(label, size=11, color="grey800", weight="bold", text_align=ft.TextAlign.LEFT),
                    alignment=ft.alignment.center_left, 
                    expand=True, 
                    padding=ft.padding.only(right=5)
                ),
                ft.Container(
                    content=ft.TextField(
                        value=value, 
                        # 3. TAMAÑO DEL TEXTO: Si agrandas la caja, quizás quieras subir esto a 14
                        text_size=12, 
                        border=ft.InputBorder.NONE, 
                        # 2. ALTO INTERNO: Debe ser igual o ligeramente mayor que el contenedor
                        height=40, 
                        # 4. ALINEACIÓN VERTICAL: Si cambias la altura, ajusta el 'bottom' para centrar el texto
                        content_padding=ft.padding.only(left=10, right=10, bottom=18), 
                        text_align=ft.TextAlign.LEFT
                    ),
                    border=ft.border.all(1, self.COLOR_BORDER), 
                    border_radius=5, 
                    bgcolor="white",
                    
                    # --- CAMBIOS PRINCIPALES AQUÍ ---
                    height=35,  # 1. ALTO DE LA CAJA (Antes 28)
                    width=100,  # 1. ANCHO DE LA CAJA (Antes 80). Pon 120 o 150 si quieres.
                    # --------------------------------
                    
                    alignment=ft.alignment.center_left
                )
            ], spacing=0, vertical_alignment=ft.CrossAxisAlignment.CENTER)

        # Creación de inputs de parámetros
        self.in_alpha = input_param("Alpha (Higher value prioritizes more equal hours) Recommended = 1", str(get_p('alpha', 1.0)))
        self.in_beta = input_param("Beta (Penalizes having two identical tasks in consecutive hours) Recommended = 0.1", str(get_p('beta', 0.1)))
        self.in_gamma = input_param("Gamma (Penalizes having gaps between tasks) Recommended = 0.01", str(get_p('gamma', 0.01)))
        self.in_epsilon = input_param("Epsilon (Penalizes not obeying the mandatory tasks matrix) Recommended = 100", str(get_p('epsilon', 100)))
        self.in_timelimit = input_param("Max Time (sec)", str(get_p('timelimit', 60)))

        # === PANEL IZQUIERDO (Inputs y Configuración) ===
        people_task_row = ft.Row(
            controls=[
                ft.Container(
                    width=180, height=250, 
                    content=ft.Column([
                        ft.Text("1. People", color=self.COLOR_TEXT_HIGHLIGHT, weight="bold", size=20),
                        ft.Container(
                            content=self.txt_people,
                            border=ft.border.all(1, self.COLOR_BORDER),
                            border_radius=5, bgcolor="white", padding=5, expand=True
                        )
                    ], expand=True, spacing=2)
                ),
                ft.Container(
                    width=180, height=250, 
                    content=ft.Column([
                        ft.Text("2. Tasks", color=self.COLOR_TEXT_HIGHLIGHT, weight="bold", size=20),
                        ft.Container(
                            content=self.txt_tasks,
                            border=ft.border.all(1, self.COLOR_BORDER),
                            border_radius=5, bgcolor="white", padding=5, expand=True
                        )
                    ], expand=True, spacing=2)
                ),
            ],
            spacing=10
        )

        # --- NUEVO: Selector de Solver ---
        # Cargamos el solver guardado o por defecto 'highs'
        default_solver = self.data.get('solver', 'highs') if self.data else 'highs'
        
        self.solver_selector = ft.RadioGroup(
            content=ft.Row([
                ft.Radio(value="cbc", label="CBC (Standard)"),
                ft.Radio(value="highs", label="HiGHS (Fast)")
            ]),
            value=default_solver
        )

        config_below = ft.Column([
            ft.Text("3. Active Hours", color=self.COLOR_TEXT_HIGHLIGHT, weight="bold", size=20),
            ft.Container(
                content=self.container_hours, 
                border=ft.border.all(1, self.COLOR_BORDER),
                padding=5, border_radius=5, bgcolor="white",
                width=370
            ),
            ft.Divider(height=10),
            # AÑADIDO AQUI EL TEXTO Y EL SELECTOR
            ft.Text("4. Solver Engine", color=self.COLOR_TEXT_HIGHLIGHT, weight="bold", size=20),
            self.solver_selector,
            ft.Divider(height=10),
            ft.Text("5. Parameters", color=self.COLOR_TEXT_HIGHLIGHT, weight="bold", size=20),
            ft.Column([self.in_alpha, self.in_beta, self.in_gamma, self.in_epsilon, self.in_timelimit], spacing=2)
        ], spacing=10)

        left_panel = ft.Container(
            content=ft.Column([people_task_row, ft.Container(height=10), config_below], spacing=0, scroll=ft.ScrollMode.AUTO),
            width=400, 
            padding=ft.padding.only(right=10),
            border=ft.border.only(right=ft.border.BorderSide(1, self.COLOR_BORDER)),
            alignment=ft.alignment.top_center
        )

        # === PANEL DERECHO (MATRICES - PESTAÑA 1) ===
        matrices_wrapper = ft.Container(
            content=self.content_matrices,
            padding=ft.padding.only(right=25, bottom=25, left=10), 
        )

        # Layout con Scroll bidireccional simulado (Row con scroll + Column con scroll)
        right_panel = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Column(
                        controls=[matrices_wrapper],
                        scroll=ft.ScrollMode.ALWAYS, # Scroll Vertical
                        expand=True 
                    )
                ],
                scroll=ft.ScrollMode.ALWAYS, # Scroll Horizontal
                expand=True,
                vertical_alignment=ft.CrossAxisAlignment.START
            ),
            expand=True
        )

        main_layout = ft.Row(
            controls=[left_panel, right_panel],
            vertical_alignment=ft.CrossAxisAlignment.START,
            spacing=0,
            expand=True
        )

        # === PESTAÑA 2 (OBLIGATORIEDADES) ===
        mandatory_wrapper = ft.Container(
            content=self.content_mandatory,
            padding=ft.padding.only(right=25, bottom=25, left=10),
        )

        mandatory_panel = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Column(
                        controls=[mandatory_wrapper],
                        scroll=ft.ScrollMode.ALWAYS,
                        expand=True
                    )
                ],
                scroll=ft.ScrollMode.ALWAYS,
                expand=True,
                vertical_alignment=ft.CrossAxisAlignment.START
            ),
            expand=True
        )

        # Sistema de Pestañas
        self.tabs_control = ft.Tabs(
            selected_index=0, animation_duration=0, expand=True,
            tabs=[
                ft.Tab(text="1. Configuration & Data", content=ft.Container(content=main_layout, padding=10)),
                ft.Tab(text="2. Mandatory Tasks", content=ft.Container(content=mandatory_panel, padding=10)),
            ]
        )

        # Botón de ejecución
        self.btn_optimize = ft.ElevatedButton(
            "CALCULATE (OPTIMIZE)", on_click=self.run_optimization_thread, 
            height=40, width=200, style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=2), color="white", bgcolor="blue")
        )

        # Barra inferior (Status, Créditos)
        self.page.add(
            ft.Container(content=self.tabs_control, expand=True),
            ft.Container(
                content=ft.Row([
                    ft.Row([
                        self.btn_optimize,
                        ft.VerticalDivider(width=20),
                        self.progress_bar, 
                        self.status_text
                    ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    ft.Row([
                        ft.Text("Developed by ", size=10, color="grey400", italic=True),
                        ft.TextButton(
                            content=ft.Text("Óscar Martínez Zamora", size=10, color="blue400", italic=True),
                            url="https://www.linkedin.com/in/oscarmartinezzamora/",
                            style=ft.ButtonStyle(padding=0, overlay_color="transparent"),
                            tooltip="Ver perfil en LinkedIn"
                        )], spacing=0, vertical_alignment=ft.CrossAxisAlignment.CENTER)
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                padding=10, bgcolor=self.COLOR_BG_PANEL,
                border=ft.border.only(top=ft.border.BorderSide(1, self.COLOR_BORDER))
            )
        )

        # Generar las tablas inicialmente
        self.generate_tables()

    def generate_hour_buttons(self):
        # Genera dinámicamente los botones de selección de hora en dos filas.
        row1 = ft.Row(wrap=False, spacing=2) 
        row2 = ft.Row(wrap=False, spacing=2)
        split_idx = 9
        for i, h in enumerate(self.possible_hours):
            is_active = self.state_hours.get(i, 0)
            btn = ft.Container(
                content=ft.Text(f"{h:02d}h", size=11, weight=ft.FontWeight.BOLD, color="white" if is_active else "black"),
                width=38, height=25, bgcolor="#217346" if is_active else self.COLOR_NEUTRAL, 
                alignment=ft.alignment.center, border_radius=5, on_click=lambda e, idx=i: self.toggle_hour(e, idx)
            )
            if i < split_idx: row1.controls.append(btn)
            else: row2.controls.append(btn)
        self.container_hours.controls = [row1, row2]
        if self.page: self.page.update()

    def toggle_hour(self, e, idx):
        # Callback al hacer click en una hora: cambia estado y regenera tablas.
        curr = self.state_hours[idx]
        self.state_hours[idx] = 1 - curr
        is_active = self.state_hours[idx]
        e.control.bgcolor = "#217346" if is_active else self.COLOR_NEUTRAL
        e.control.content.color = "white" if is_active else "black"
        e.control.update()
        self.generate_tables()

    def on_input_change(self, e):
        # Detecta cambios en los cuadros de texto de Personas/Tareas.
        if e.control == self.txt_people:
            if self.input_people_val == self.txt_people.value: return
            self.input_people_val = self.txt_people.value
        elif e.control == self.txt_tasks:
            if self.input_tasks_val == self.txt_tasks.value: return
            self.input_tasks_val = self.txt_tasks.value
        self.generate_tables()

    def toggle_matrix_btn(self, e):
        # Manejador genérico para clicks en celdas tipo botón (D, Q, F).
        data = e.control.data
        self._update_single_cell(e.control, data['tipo'], data['k1'], data['k2'], data['k3'])

    def _update_single_cell(self, control, tipo, k1, k2, k3, force_val=None):
        """
        Actualiza el estado lógico y visual de una celda específica.
        Cambia el color de fondo y el texto entre YES/NO o Color/Blanco.
        """
        current_val = 0
        if force_val is not None:
            pass 
        else:
            # Obtener valor actual
            if tipo == 'D': current_val = self.state_D.get(k1, {}).get(k2, 1)
            elif tipo == 'Q': current_val = self.state_Q.get(k1, {}).get(k2, 1)
            elif tipo == 'F': current_val = self.state_F.get(k1, {}).get(k2, {}).get(k3, 0)
            force_val = 1 - current_val # Invertir valor

        new_val = force_val
        # Guardar nuevo valor
        if tipo == 'D': self.state_D.setdefault(k1, {})[k2] = new_val
        elif tipo == 'Q': self.state_Q.setdefault(k1, {})[k2] = new_val
        elif tipo == 'F': self.state_F.setdefault(k1, {}).setdefault(k2, {})[k3] = new_val
        
        # Actualización Visual
        bg_color = "white"
        text_color = "grey"
        text_value = ""

        if new_val:
            if tipo == 'F':
                bg_color = self.task_colors.get(k2, self.COLOR_ACTIVE)
                text_value = "" 
            else:
                bg_color = self.COLOR_ACTIVE
                text_color = self.TEXT_ACTIVE
                text_value = "YES"
        else:
            if tipo == 'D' or tipo == 'Q':
                bg_color = self.COLOR_INACTIVE
                text_color = self.TEXT_INACTIVE
                text_value = "NO"
            elif tipo == 'F':
                bg_color = "white"
                text_value = ""
        
        control.bgcolor = bg_color
        control.content.color = text_color
        control.content.value = text_value
        control.update()

    def _get_val_from_memory_or_json(self, tipo, k1, k2=None, k3=None):
        """
        Intenta obtener el valor de una celda desde la memoria (edición actual),
        si no existe, desde el JSON cargado, y si no, devuelve un default.
        """
        if tipo == 'D':
            if k1 in self.state_D and k2 in self.state_D[k1]: return self.state_D[k1][k2]
            if self.data and 'D' in self.data: return int(self.data['D'].get(k1, {}).get(str(k2), 1))
            return 1
        if tipo == 'Q':
            if k1 in self.state_Q and k2 in self.state_Q[k1]: return self.state_Q[k1][k2]
            if self.data and 'Q' in self.data: return int(self.data['Q'].get(k1, {}).get(k2, 1))
            return 1
        if tipo == 'F':
            if k1 in self.state_F and k2 in self.state_F[k1] and k3 in self.state_F[k1][k2]: return self.state_F[k1][k2][k3]
            if self.data and 'F' in self.data: return int(self.data['F'].get(k1, {}).get(k2, {}).get(str(k3), 0))
            return 0
        if tipo == 'R':
            if k1 in self.state_R and k2 in self.state_R[k1]: return self.state_R[k1][k2]
            if self.data and 'R' in self.data: return int(self.data['R'].get(k1, {}).get(str(k2), 0))
            return 0
        return 0

    def create_bulk_action_cell(self, action_type, matrix_type, key, width=28, height=28):
        # Crea el botón pequeño de la cabecera para activar/desactivar toda una fila o columna.
        icon = ft.Icons.SWAP_HORIZ if action_type == 'row' else ft.Icons.SWAP_VERT
        tooltip = f"{action_type} Action {matrix_type}"
        
        if matrix_type == 'R':
            icon = ft.Icons.EXPOSURE_ZERO
            tooltip = "Reset to 0 (Clear)"

        return ft.Container(
            width=width, height=height, 
            content=ft.Icon(icon, size=16, color="grey"), 
            bgcolor="#eeeeee", border_radius=3, alignment=ft.alignment.center,
            tooltip=tooltip,
            # GUARDAMOS LA INFO EN 'data' EN LUGAR DE USAR LAMBDA
            data={'action': action_type, 'matrix': matrix_type, 'key': key},
            on_click=self.on_bulk_click_wrapper, # Llamamos a un wrapper
            ink=True
        )

    # NUEVA FUNCIÓN WRAPPER PARA GESTIONAR EL CLICK
    def on_bulk_click_wrapper(self, e):
        d = e.control.data
        self.execute_bulk_action(d['action'], d['matrix'], d['key'])

    def execute_bulk_action(self, action_type, matrix_type, key):
        """Lógica para modificar filas/columnas enteras al hacer click en el botón bulk."""
        
        # --- CASO 1: Matriz de Requerimientos (R) - Resetea a 0 ---
        if matrix_type == 'R':
            cols = self.indices_horas
            rows = self.tasks
            
            def reset_cell_R(t, h):
                self.state_R.setdefault(t, {})[h] = 0
                try:
                    # Actualizar visualmente el TextField
                    t_idx = self.tasks.index(t)
                    h_idx = self.indices_horas.index(h)
                    tf = self.r_cells.get((t_idx, h_idx))
                    if tf:
                        tf.value = "" # Vacío visualmente es 0
                        tf.update()
                except:
                    pass

            if action_type == 'row':
                for c in cols: reset_cell_R(key, c)
            elif action_type == 'col':
                for r in rows: reset_cell_R(r, key)
            return
        
        # --- CASO 2: Matrices Booleanas (D y Q) - Toggle YES/NO ---
        target_dict = None
        
        if matrix_type == 'D':
            cols = self.indices_horas
            rows = self.people
            target_dict = self.grid_controls['D']
        elif matrix_type == 'Q':
            cols = self.tasks
            rows = self.people
            target_dict = self.grid_controls['Q']
        else:
            return
        
        # --- CORRECCIÓN AQUÍ: Obtener el valor REAL actual ---
        # Leemos el valor de la primera celda usando la función que consulta JSON + Memoria
        current_val = 1 
        
        if action_type == 'row':
            first_col = cols[0] if cols else None
            if first_col is not None:
                # key = Persona, first_col = Hora/Tarea
                current_val = self._get_val_from_memory_or_json(matrix_type, key, first_col)
                
        elif action_type == 'col':
            first_row = rows[0] if rows else None
            if first_row is not None:
                # first_row = Persona, key = Hora/Tarea
                current_val = self._get_val_from_memory_or_json(matrix_type, first_row, key)
        
        # Invertimos el valor encontrado
        new_val = 1 - current_val

        # Aplicar cambio masivo
        if action_type == 'row':
            for c in cols:
                ctrl = target_dict.get((key, c))
                if ctrl:
                    self._update_single_cell(ctrl, matrix_type, key, c, None, force_val=new_val)
        
        elif action_type == 'col':
            for r in rows:
                ctrl = target_dict.get((r, key))
                if ctrl:
                    self._update_single_cell(ctrl, matrix_type, r, key, None, force_val=new_val)

    def generate_tables(self):
        """
        FUNCIÓN CRÍTICA: Reconstruye toda la interfaz de cuadrículas (Grids).
        Se llama cada vez que cambian las personas, tareas u horas activas.
        """
        # 1. Procesar Personas (Eliminar duplicados manteniendo el orden)
        raw_people = [t.strip() for t in self.txt_people.value.split('\n') if t.strip()]
        self.people = list(dict.fromkeys(raw_people)) 

        # 2. Procesar Tareas
        raw_tasks = [t.strip() for t in self.txt_tasks.value.split('\n') if t.strip()]
        self.tasks = list(dict.fromkeys(raw_tasks))

        # Determinar qué horas están seleccionadas
        self.indices_horas = [i for i, v in self.state_hours.items() if v == 1]
        self.indices_horas.sort()

        # Asignar colores a tareas
        self.task_colors = {}
        for i, t in enumerate(self.tasks):
            color = self.available_colors[i % len(self.available_colors)]
            self.task_colors[t] = color

        # Resetear referencias de controles
        self.grid_controls = {'D': {}, 'Q': {}}
        self.r_cells = {} 

        if not self.people or not self.tasks or not self.indices_horas:
            self.status_text.value = "Warning: Missing data (people, tasks or hours)."
            if self.page: self.status_text.update()
            return

        self.status_text.value = "Regenerating spreadsheet view..."
        if self.page: self.status_text.update()

        # Dimensiones de celdas
        CELL_W_NAME = 80 
        CELL_W_HOUR = 35  
        CELL_W_TASK = 50 
        CELL_W_TASK_LABEL = 80 
        CELL_W_BUTTON = 22 
        CELL_H = 22
        FONT_SIZE = 10

        container_d = ft.Column(spacing=2)
        container_r = ft.Column(spacing=2)
        container_q = ft.Column(spacing=2)
        container_f = ft.Column(spacing=2)

        # Helpers para crear celdas de cabecera y nombres
        def cell_header(text, width):
            return ft.Container(
                width=width, height=CELL_H,
                content=ft.Text(text, size=FONT_SIZE, weight="bold", color="black"), 
                alignment=ft.alignment.center,
                bgcolor=self.COLOR_HEADER_BG,
                border=None, border_radius=3
            )

        def cell_name(text, width=CELL_W_NAME): 
            return ft.Container(
                width=width, height=CELL_H,
                content=ft.Text(text, size=FONT_SIZE, color="black"), 
                alignment=ft.alignment.center_left, padding=ft.padding.only(left=5),
                bgcolor="white", border=None, border_radius=3
            )

        # 1. MATRIZ DE DISPONIBILIDAD (D)
        rows_d = []
        header_controls = [ft.Container(width=CELL_W_BUTTON, height=CELL_H), cell_header("Person", CELL_W_NAME)]
        for h in self.indices_horas:
            header_controls.append(cell_header(f"{self.possible_hours[h]:02d}h", CELL_W_HOUR))
        rows_d.append(ft.Row(controls=header_controls, spacing=2))

        for pers in self.people:
            row_ctrls = [
                self.create_bulk_action_cell('row', 'D', pers, width=CELL_W_BUTTON, height=CELL_H),
                cell_name(pers)
            ]
            for h in self.indices_horas:
                row_ctrls.append(self.create_cell_button_scaled("YES", "NO", 'D', pers, h, width=CELL_W_HOUR, height=CELL_H, font_size=FONT_SIZE))
            rows_d.append(ft.Row(controls=row_ctrls, spacing=2))

        # 2. MATRIZ DE REQUERIMIENTOS (R) - Estilo Excel (Inputs Numéricos)
        rows_r = []
        header_r_act = [ft.Container(width=CELL_W_BUTTON + CELL_W_TASK_LABEL, height=CELL_H)]
        for h in self.indices_horas:
            header_r_act.append(self.create_bulk_action_cell('col', 'R', h, width=CELL_W_HOUR, height=CELL_H))
        rows_r.append(ft.Row(controls=header_r_act, spacing=2))

        header_r = [ft.Container(width=CELL_W_BUTTON, height=CELL_H), cell_header("Task", CELL_W_TASK_LABEL)]
        for h in self.indices_horas:
            header_r.append(cell_header(f"{self.possible_hours[h]:02d}h", CELL_W_HOUR)) 
        rows_r.append(ft.Row(controls=header_r, spacing=2))

        for i, t in enumerate(self.tasks):
            row_ctrls = [
                self.create_bulk_action_cell('row', 'R', t, width=CELL_W_BUTTON, height=CELL_H),
                cell_name(t, width=CELL_W_TASK_LABEL)
            ]
            for j, h in enumerate(self.indices_horas):
                val = self._get_val_from_memory_or_json('R', t, h)
                row_ctrls.append(self.create_excel_input(t, h, val, i, j, width=CELL_W_HOUR, height=CELL_H, font_size=FONT_SIZE))
            rows_r.append(ft.Row(controls=row_ctrls, spacing=2))

        # 3. MATRIZ DE HABILIDADES (Q)
        rows_q = []
        header_q_act = [ft.Container(width=CELL_W_NAME, height=CELL_H)]
        for t in self.tasks:
            header_q_act.append(self.create_bulk_action_cell('col', 'Q', t, width=CELL_W_TASK, height=CELL_H))
        rows_q.append(ft.Row(controls=header_q_act, spacing=2))

        header_q = [cell_header("Person", CELL_W_NAME)]
        for t in self.tasks:
            header_q.append(cell_header(t, CELL_W_TASK))
        rows_q.append(ft.Row(controls=header_q, spacing=2))

        for pers in self.people:
            row_ctrls = [cell_name(pers)]
            for t in self.tasks:
                row_ctrls.append(self.create_cell_button_scaled("YES", "NO", 'Q', pers, t, width=CELL_W_TASK, height=CELL_H, font_size=FONT_SIZE))
            rows_q.append(ft.Row(controls=row_ctrls, spacing=2))

        # 4. MATRIZ DE OBLIGATORIOS (F) - Se dibuja una tabla por cada tarea
        list_f = []
        for t in self.tasks:
            color_task = self.task_colors.get(t, self.COLOR_HEADER_BG)
            
            list_f.append(
                ft.Container(
                    content=ft.Text(f" {t} ", size=12, color="black", weight="bold"),
                    bgcolor=color_task, padding=2, border_radius=3
                )
            )
            
            task_rows = []
            h_row = [cell_header("Person", CELL_W_NAME)]
            for h in self.indices_horas:
                h_row.append(cell_header(f"{self.possible_hours[h]:02d}h", CELL_W_HOUR))
            task_rows.append(ft.Row(controls=h_row, spacing=2))

            for pers in self.people:
                r_row = [cell_name(pers)]
                for h in self.indices_horas:
                    r_row.append(self.create_cell_button_scaled("YES", "", 'F', pers, t, h, width=CELL_W_HOUR, height=CELL_H, font_size=FONT_SIZE))
                task_rows.append(ft.Row(controls=r_row, spacing=2))
            
            list_f.append(ft.Column(task_rows, spacing=2))
            list_f.append(ft.Divider(height=20, color="transparent"))

        container_d.controls = rows_d
        container_r.controls = rows_r
        container_q.controls = rows_q
        container_f.controls = list_f

        def title_separator(text):
            return ft.Container(
                content=ft.Text(text, size=14, weight="bold", color=self.COLOR_TEXT_HIGHLIGHT),
                padding=ft.padding.only(top=20, bottom=5)
            )

        # Inyectar controles en las columnas contenedoras
        self.content_matrices.controls = [
            title_separator("1. Availability (D)"), 
            ft.Row([container_d], scroll=ft.ScrollMode.AUTO),
            title_separator("2. Requirements for every Task (R)"), 
            ft.Row([container_r], scroll=ft.ScrollMode.AUTO),
            title_separator("3. Skills/Qualifications (Q)"), 
            ft.Row([container_q], scroll=ft.ScrollMode.AUTO),
            ft.Container(height=30)
        ]

        self.content_mandatory.controls = [
            container_f
        ]

        self.status_text.value = "Matrices generated."
        if self.page: self.page.update()

    def create_cell_button_scaled(self, label_active, label_inactive, tipo, k1, k2, k3=None, width=60, height=28, font_size=10):
        # Crea un botón interactivo (Container con evento click) para las celdas de las matrices booleanas.
        val = self._get_val_from_memory_or_json(tipo, k1, k2, k3)
        
        # Guardar valor inicial en memoria
        if tipo == 'D': self.state_D.setdefault(k1, {})[k2] = val
        elif tipo == 'Q': self.state_Q.setdefault(k1, {})[k2] = val
        elif tipo == 'F': self.state_F.setdefault(k1, {}).setdefault(k2, {})[k3] = val

        bg_color = "white"
        text_color = "grey"
        text_value = label_inactive
        
        border_style = None
        if tipo == 'F': border_style = ft.border.all(1, "#e0e0e0")

        # Determinar estilo inicial
        if val:
            if tipo == 'F':
                bg_color = self.task_colors.get(k2, self.COLOR_ACTIVE)
                text_color = "black" 
                text_value = ""
            else:
                bg_color = self.COLOR_ACTIVE
                text_color = self.TEXT_ACTIVE
                text_value = label_active
        else:
            if tipo == 'D' or tipo == 'Q':
                bg_color = self.COLOR_INACTIVE
                text_color = self.TEXT_INACTIVE
                text_value = label_inactive
            elif tipo == 'F':
                bg_color = "white"
                text_value = ""

        container = ft.Container(
            width=width, height=height, bgcolor=bg_color, 
            border=border_style, 
            border_radius=3,
            alignment=ft.alignment.center,
            content=ft.Text(text_value, size=font_size, color=text_color),
            data={'tipo': tipo, 'k1': k1, 'k2': k2, 'k3': k3}, # Metadata para el manejador de eventos
            on_click=self.toggle_matrix_btn
        )
        
        # Guardar referencia para acceso rápido (Bulk updates)
        if tipo == 'D': self.grid_controls['D'][(k1, k2)] = container
        elif tipo == 'Q': self.grid_controls['Q'][(k1, k2)] = container
            
        return container

    def create_excel_input(self, t, h, val_initial, row_idx, col_idx, width=70, height=20, font_size=10):
        # Crea una celda de input numérico para la matriz de requerimientos (R).
        self.state_R.setdefault(t, {})[h] = val_initial

        def on_change(e):
            val_str = e.control.value
            if not val_str: new_val = 0
            else:
                try: new_val = int(val_str)
                except ValueError: new_val = 0
            self.state_R.setdefault(t, {})[h] = new_val

        def on_focus(e):
            # Seleccionar todo el texto al hacer foco
            e.control.selection_start = 0
            e.control.selection_end = len(e.control.value)
            e.control.update()

        def on_submit(e):
            # Mover foco a la siguiente fila al dar Enter
            next_row = row_idx + 1
            next_cell = self.r_cells.get((next_row, col_idx))
            if next_cell: next_cell.focus()

        display_val = str(val_initial) if val_initial != 0 else ""

        txt_field = ft.TextField(
            value=display_val,
            text_size=font_size,
            width=width, height=height,
            content_padding=ft.padding.only(bottom=21), 
            text_align=ft.TextAlign.CENTER,
            border=ft.InputBorder.NONE,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(allow=True, regex_string=r"^[0-9]*$", replacement_string=""),
            on_change=on_change, on_focus=on_focus, on_submit=on_submit
        )
        
        self.r_cells[(row_idx, col_idx)] = txt_field

        return ft.Container(
            content=txt_field,
            width=width, height=height,
            bgcolor="white",
            border=ft.border.all(1, "#e0e0e0"),
            border_radius=5 
        )

    def run_optimization_thread(self, e):
        # Manejador del botón 'Optimize'. Lanza el cálculo en un hilo aparte.
        if not self.indices_horas:
            self.status_text.value = "Error: No hours."
            self.status_text.update()
            return

        self.btn_optimize.disabled = True
        self.progress_bar.visible = True 
        self.status_text.value = "Optimizing (this may take a while)..."
        self.page.update()

        # Threading evita que la interfaz se congele ("Not Responding") mientras PuLP calcula
        t = threading.Thread(target=self._run_optimization)
        t.start()

    def _run_optimization(self):
        # Función interna que ejecuta el proceso de optimización.
        try:
            model, X, W, W_max, W_min = self.gather_data_and_solve()
            status_txt = LpStatus[model.status]
            self.status_text.value = f"Finished: {status_txt}"
            self.show_results_dialog(model, X, W, W_max, W_min)
        except Exception as ex:
            import traceback
            traceback.print_exc()
            self.status_text.value = f"Error: {str(ex)}"
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Critical Error: {str(ex)}"), bgcolor="red")
            self.page.snack_bar.open = True
            self.page.update()
        finally:
            self.btn_optimize.disabled = False
            self.progress_bar.visible = False 
            self.page.update()

    def gather_data_and_solve(self):
        # Recopila todos los datos de la UI, los guarda y llama al solver.
        def get_val(ctrl):
            try: return float(ctrl.controls[1].content.value)
            except: return 0.0

        # Preparación de diccionarios para JSON
        R_data = {}
        for t in self.tasks:
            R_data[t] = {}
            for h in self.indices_horas:
                val = self.state_R.get(t, {}).get(h, 0)
                R_data[t][h] = val

        F_save = {}
        for pers in self.people:
            F_save[pers] = {}
            for t in self.tasks:
                F_save[pers][t] = {}
                for h in self.indices_horas:
                    F_save[pers][t][str(h)] = self.state_F.setdefault(pers, {}).setdefault(t, {}).get(h, 0)

        # Capturamos el solver seleccionado
        selected_solver = self.solver_selector.value

        final_data = {
            'people': self.people, 'tasks': self.tasks, 'hours': self.indices_horas,
            'D': {k: {str(idx): v for idx, v in d.items()} for k, d in self.state_D.items()},
            'Q': self.state_Q,
            'R': {k: {str(idx): v for idx, v in d.items()} for k, d in R_data.items()},
            'F': F_save,
            'alpha': get_val(self.in_alpha), 'beta': get_val(self.in_beta), 'gamma': get_val(self.in_gamma),
            'epsilon': get_val(self.in_epsilon), 'timelimit': int(get_val(self.in_timelimit)),
            'solver': selected_solver # <--- GUARDAMOS LA SELECCIÓN
        }
        # Persistencia
        save_data(final_data)

        # Preparar datos para el Solver
        solver_data = final_data.copy()
        solver_data['D'] = {k: {int(idx): v for idx, v in d.items()} for k, d in final_data['D'].items()}
        solver_data['R'] = {k: {int(idx): v for idx, v in d.items()} for k, d in final_data['R'].items()}
        F_solver = {}
        for i in self.people:
            F_solver[i] = {}
            for t in self.tasks:
                F_solver[i][t] = {}
                for h in self.indices_horas: 
                    F_solver[i][t][h] = self.state_F.get(i, {}).get(t, {}).get(h, 0)
        solver_data['F'] = F_solver

        return solve_model(solver_data)

    def save_excel_results(self, X, W):
        # Exporta los resultados a un archivo Excel formateado.
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Staffing Plan"

            # Estilos de Excel
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            center_align = Alignment(horizontal="center", vertical="center")
            border_style = Side(border_style="thin", color="000000")
            full_border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

            # Cabeceras
            headers = ["Person"] + [f"{self.possible_hours[h]:02d}h" for h in self.indices_horas] + ["Total"]
            ws.append(headers)

            for col_num, cell in enumerate(ws[1], 1):
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = full_border
                ws.column_dimensions[get_column_letter(col_num)].width = 15

            # Rellenar datos
            for i in self.people:
                row_idx = ws.max_row + 1
                
                cell_name = ws.cell(row=row_idx, column=1, value=i)
                cell_name.font = Font(bold=True)
                cell_name.border = full_border

                for idx_h, h in enumerate(self.indices_horas):
                    assigned_task = ""
                    color_hex = "FFFFFF"
                    
                    # Verificar disponibilidad para marcar rojo si no estaba disponible (pero se asignó)
                    is_available = self._get_val_from_memory_or_json('D', i, h)
                    if not is_available:
                        color_hex = self.FLET_TO_HEX["red100"]

                    # Verificar si se asignó tarea
                    for t in self.tasks:
                        if value(X[i][t][h]) == 1:
                            assigned_task = t
                            flet_color = self.task_colors.get(t, "white")
                            color_hex = self.FLET_TO_HEX.get(flet_color, "FFFFFF")
                            break
                    
                    cell = ws.cell(row=row_idx, column=idx_h + 2, value=assigned_task)
                    cell.alignment = center_align
                    cell.border = full_border
                    cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

                # Columna Total
                val_load = value(W[i]) if value(W[i]) is not None else 0
                cell_total = ws.cell(row=row_idx, column=len(self.indices_horas) + 2, value=int(val_load))
                cell_total.font = Font(bold=True)
                cell_total.alignment = center_align
                cell_total.border = full_border

            filename = "staffing_plan.xlsx"
            wb.save(filename)
            
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Saved: {filename}"), bgcolor="green")
            self.page.snack_bar.open = True
            self.page.update()

            # Intentar abrir el archivo automáticamente
            try:
                if platform.system() == 'Darwin': # macOS
                    subprocess.call(('open', filename))
                elif platform.system() == 'Windows':
                    os.startfile(filename)
                else: # Linux
                    subprocess.call(('xdg-open', filename))
            except Exception as e_open:
                print(f"Could not auto-open file: {e_open}")

        except Exception as ex:
            self.page.snack_bar = ft.SnackBar(ft.Text(f"Excel Error: {str(ex)}"), bgcolor="red")
            self.page.snack_bar.open = True
            self.page.update()

    def show_results_dialog(self, model, X, W, W_max, W_min):
        # Muestra una ventana modal con el resultado de la optimización (grid coloreado y métricas).
        status_txt = LpStatus[model.status]
        
        # Caso: No se encontró solución
        if status_txt != "Optimal":
            content_dlg = ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.WARNING, color="red", size=40),
                    ft.Text(f"No feasible solution found.\nStatus: {status_txt}", size=16, color="red", weight="bold", text_align="center"),
                    ft.Text("Please check constraints (Availability and Requirements) and try again.", text_align="center")
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                height=200, alignment=ft.alignment.center
            )
            title_dlg = ft.Text("Attention", color="red")
            actions_dlg = [ft.TextButton("Close", on_click=lambda e: self.page.close(dlg))]
            
            dlg = ft.AlertDialog(
                title=title_dlg, 
                content=content_dlg,
                actions=actions_dlg,
                shape=ft.RoundedRectangleBorder(radius=5),
                modal=True,
            )
            self.page.open(dlg)
            self.page.update()
            return
        
        # === CASO ÓPTIMO ===
        val_wmax = value(W_max) if value(W_max) is not None else 0
        val_wmin = value(W_min) if value(W_min) is not None else 0
        load_gap = int(val_wmax - val_wmin)

        # --- 1. PRE-CÁLCULO DE MÉTRICAS Y DESCANSOS ---
        total_monotony = 0
        total_breaks = 0
        
        # Diccionario para saber rápidamente qué celdas marcar con borde rojo
        # Clave: (persona, indice_hora_absoluto), Valor: True
        break_cells_map = {} 

        for p in self.people:
            # A) Monotonía
            for t in self.tasks:
                for idx in range(len(self.indices_horas) - 1):
                    h_curr = self.indices_horas[idx]
                    h_next = self.indices_horas[idx+1]
                    if value(X[p][t][h_curr]) == 1 and value(X[p][t][h_next]) == 1:
                        total_monotony += 1
            
            # B) Identificar Descansos Intermedios exactos
            # Construimos lista de índices donde la persona trabaja
            working_indices = []
            for idx_local, h_real in enumerate(self.indices_horas):
                # Ver si trabaja en alguna tarea en esta hora
                works_now = False
                for t in self.tasks:
                    if value(X[p][t][h_real]) == 1:
                        works_now = True
                        break
                if works_now:
                    working_indices.append(idx_local)
            
            # Si trabajó al menos 2 horas separadas, puede haber huecos en medio
            if len(working_indices) >= 2:
                start_idx = working_indices[0]
                end_idx = working_indices[-1]
                
                # Recorremos el rango desde el inicio hasta el fin de su jornada
                # Si un índice NO está en working_indices, es un descanso intermedio.
                current = start_idx + 1
                while current < end_idx:
                    if current not in working_indices:
                        # Es un descanso intermedio
                        h_real_break = self.indices_horas[current]
                        break_cells_map[(p, h_real_break)] = True
                        
                        # Contar solo bloques (si el anterior ya era descanso, no sumamos al contador global, 
                        # pero SÍ marcamos la celda para pintarla)
                        prev_is_break = ((current - 1) > start_idx) and ((current - 1) not in working_indices)
                        if not prev_is_break:
                            total_breaks += 1
                            
                    current += 1

        # Estado local para controlar el Zoom
        zoom_state = {"scale": 1.0}
        
        # Constantes base
        BASE_W_NAME = 100
        BASE_W_HOUR = 50
        BASE_W_TOTAL = 50
        BASE_H_ROW = 20
        BASE_FONT_SIZE = 11
        BASE_FONT_SIZE_SMALL = 10
        
        num_people = len(self.people)
        num_hours = len(self.indices_horas)
        
        # Dimensiones del diálogo
        content_width = BASE_W_NAME + (num_hours * (BASE_W_HOUR + 2)) + BASE_W_TOTAL + 40
        content_height = (num_people + 1) * (BASE_H_ROW + 2) + 40
        
        MIN_WIDTH = 400
        MAX_WIDTH = 1200
        MIN_HEIGHT = 150
        MAX_HEIGHT = 600
        
        dialog_width = max(MIN_WIDTH, min(MAX_WIDTH, content_width))
        dialog_height = max(MIN_HEIGHT, min(MAX_HEIGHT, content_height))

        table_container = ft.Column(spacing=2)
        zoom_label = ft.Text(f"100%", size=12, weight="bold", width=50, text_align=ft.TextAlign.CENTER)

        def build_table(scale):
            W_NAME = int(BASE_W_NAME * scale)
            W_HOUR = int(BASE_W_HOUR * scale)
            W_TOTAL = int(BASE_W_TOTAL * scale)
            H_ROW = int(BASE_H_ROW * scale)
            FONT_SIZE = int(BASE_FONT_SIZE * scale)
            FONT_SIZE_SMALL = int(BASE_FONT_SIZE_SMALL * scale)

            def make_res_cell(content, width, bgcolor="white", border=None):
                return ft.Container(
                    content=content, width=width, height=H_ROW, bgcolor=bgcolor,
                    alignment=ft.alignment.center, border=border, border_radius=3
                )

            rows = []
            
            # Header
            header_cells = [make_res_cell(ft.Text("Person", weight="bold", size=FONT_SIZE), W_NAME, bgcolor="#F2F2F2")]
            for h in self.indices_horas:
                header_cells.append(make_res_cell(ft.Text(f"{self.possible_hours[h]:02d}h", weight="bold", size=FONT_SIZE), W_HOUR, bgcolor="#F2F2F2"))
            header_cells.append(make_res_cell(ft.Text("Total", weight="bold", size=FONT_SIZE), W_TOTAL, bgcolor="#F2F2F2"))
            rows.append(ft.Row(header_cells, spacing=2))

            # Filas de datos
            for i in self.people:
                row_cells = []
                # Nombre persona
                row_cells.append(ft.Container(
                    content=ft.Text(i, size=FONT_SIZE, weight="bold"), width=W_NAME, height=H_ROW, bgcolor="white",
                    alignment=ft.alignment.center_left, padding=ft.padding.only(left=5), border=None, border_radius=3
                ))
                
                # Celdas de horas
                for idx in self.indices_horas:
                    assigned_task = ''
                    bg_color = "white"
                    cell_border = None
                    
                    # Disponibilidad base
                    is_available = self._get_val_from_memory_or_json('D', i, idx)
                    if not is_available:
                        bg_color = "red100"

                    # Chequear tarea asignada
                    for t in self.tasks:
                        if value(X[i][t][idx]) == 1:
                            assigned_task = t
                            bg_color = self.task_colors[t]
                            break
                    
                    # --- LÓGICA DE RESALTADO DE DESCANSOS ---
                    # Si esta celda está en nuestro mapa de descansos, aplicamos borde rojo
                    if (i, idx) in break_cells_map:
                        # Borde rojo grueso para resaltar
                        cell_border = ft.border.all(2, "red")
                        # Opcional: Si quieres que el fondo sea blanco puro para resaltar más
                        # bg_color = "white" 

                    row_cells.append(make_res_cell(
                        ft.Text(assigned_task, color="black", size=FONT_SIZE_SMALL, weight="bold"), 
                        W_HOUR, 
                        bgcolor=bg_color,
                        border=cell_border
                    ))
                
                # Total
                val_load = value(W[i]) if value(W[i]) is not None else 0
                row_cells.append(make_res_cell(ft.Text(str(int(val_load)), weight="bold", size=FONT_SIZE), W_TOTAL))
                rows.append(ft.Row(row_cells, spacing=2))

            return rows

        def update_table():
            table_container.controls = build_table(zoom_state["scale"])
            zoom_label.value = f"{int(zoom_state['scale'] * 100)}%"
            table_container.update()
            zoom_label.update()

        def zoom_in(e):
            if zoom_state["scale"] < 2.0:
                zoom_state["scale"] = min(2.0, zoom_state["scale"] + 0.1)
                update_table()

        def zoom_out(e):
            if zoom_state["scale"] > 0.5:
                zoom_state["scale"] = max(0.5, zoom_state["scale"] - 0.1)
                update_table()

        def zoom_reset(e):
            zoom_state["scale"] = 1.0
            update_table()

        table_container.controls = build_table(zoom_state["scale"])

        zoom_bar = ft.Row(
            controls=[
                ft.IconButton(icon=ft.Icons.ZOOM_OUT, icon_size=20, tooltip="Zoom Out", on_click=zoom_out),
                zoom_label,
                ft.IconButton(icon=ft.Icons.ZOOM_IN, icon_size=20, tooltip="Zoom In", on_click=zoom_in),
                ft.IconButton(icon=ft.Icons.REFRESH, icon_size=18, tooltip="Reset", on_click=zoom_reset),
            ],
            spacing=0,
        )

        content_dlg = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Column(
                        controls=[table_container],
                        scroll=ft.ScrollMode.AUTO,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        expand=True,
                    )
                ],
                scroll=ft.ScrollMode.ALWAYS,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
                expand=True,
            ),
            width=dialog_width,
            height=dialog_height,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            alignment=ft.alignment.center,
        )
        
        # Títulos y colores de advertencia
        color_mon = "red" if total_monotony > 0 else "black"
        color_brk = "red" if total_breaks > 0 else "black"

        title_dlg = ft.Row([
            ft.Text(f"Optimal Plan | Load Δ: {load_gap}", weight="bold", size=16),
            ft.VerticalDivider(width=10),
            ft.Text(f"Same task in 2 consecutive hours: {total_monotony}", weight="bold", size=16, color=color_mon),
            ft.VerticalDivider(width=10),
            ft.Text(f"Rest between tasks: {total_breaks}", weight="bold", size=16, color=color_brk),
        ], alignment=ft.MainAxisAlignment.START)
        
        actions_dlg = [
            ft.Row(
                controls=[
                    zoom_bar,
                    ft.Container(expand=True),
                    ft.ElevatedButton("Download Excel", icon=ft.Icons.DOWNLOAD, 
                                      on_click=lambda e: self.save_excel_results(X, W)),
                    ft.TextButton("Close", on_click=lambda e: self.page.close(dlg))
                ],
                alignment=ft.MainAxisAlignment.START,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                expand=True,
            )
        ]

        dlg = ft.AlertDialog(
            title=title_dlg, 
            content=content_dlg,
            actions=actions_dlg,
            actions_alignment=ft.MainAxisAlignment.CENTER,
            shape=ft.RoundedRectangleBorder(radius=5),
            modal=True,
        )
        self.page.open(dlg)
        self.page.update()
    

if __name__ == "__main__":
    app = StaffingApp()
    ft.app(target=app.main, view=ft.AppView.FLET_APP)

```