'''
1.- No esta agregado aqui el script de la base de datos para yo poderlo copiar y pegar en xampp y crear la base de datos, 
para poder continuar con la revisión y hacer pruebas.
2.- Por seguridad no se puede mostrar los datos de los usuarios, ya que se esta mostrando el usuario y 
la contraseña en el login, esto no es correcto.
'''


import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime
import mysql.connector
import os, json, shutil, webbrowser

try:
    import openpyxl
except:
    openpyxl = None

def conectar():
    return mysql.connector.connect(host="localhost", user="root", password="", database="alerta_temprana")

def cargar_datos_db():
    global GRADOS, GRADO_IDS, datos
    try:
        con = conectar(); cur = con.cursor(dictionary=True)
        cur.execute("SELECT * FROM grados ORDER BY id")
        grados_db = cur.fetchall()
        GRADOS = {}; GRADO_IDS = {}; datos = {}
        for g in grados_db:
            cur.execute("SELECT id, numero_lista, nombre_completo FROM alumnos WHERE grado_id=%s ORDER BY numero_lista", (g['id'],))
            alumnos = cur.fetchall()
            GRADOS[g['nombre']] = [f"{a['numero_lista']} {a['nombre_completo']}" for a in alumnos]
            GRADO_IDS[g['nombre']] = {f"{a['numero_lista']} {a['nombre_completo']}": a['id'] for a in alumnos}
            datos[g['nombre'].split()[0]] = g['total_alumnos']
        con.close()
        return GRADOS
    except Exception as e:
        print("Sin conexión a MySQL, abriendo en modo vacío:", e)
        GRADOS = {"1° Primaria": [], "2° Primaria": [], "3° Primaria": [], "4° Primaria": [], "5° Primaria": [], "6° Primaria": []}
        GRADO_IDS = {k: {} for k in GRADOS}
        datos = {}
        return GRADOS

def cargar_nombres_alumnos():
    try:
        con = conectar(); cur = con.cursor(dictionary=True)
        cur.execute("SELECT nombre_completo FROM alumnos ORDER BY nombre_completo")
        nombres = [r['nombre_completo'] for r in cur.fetchall()]
        con.close(); return nombres
    except:
        nombres=[]
        for g in GRADOS.values():
            for a in g:
                nombres.append(a.split(" ",1)[1] if " " in a else a)
        return nombres

def cargar_masiva_excel(ruta_excel, lbl_status=None):
    if openpyxl is None:
        messagebox.showerror("Falta librería", "Instala openpyxl: pip install openpyxl")
        return False
    try:
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
        if not filas:
            messagebox.showwarning("Vacío", "El Excel está vacío"); return False
        header = [str(c).strip().lower() if c else "" for c in filas[0]]
        idx_grado = -1; idx_num = -1; idx_nombre = -1
        for i, h in enumerate(header):
            if "grado" in h: idx_grado = i
            if "numero" in h or "lista" in h or "no." in h: idx_num = i
            if "nombre" in h: idx_nombre = i
        if idx_grado==-1 or idx_nombre==-1:
            messagebox.showerror("Formato", "Tu Excel debe tener columnas: Grado | Numero_Lista | Nombre_Completo en la fila 1")
            return False
        if idx_num==-1: idx_num = None
        datos_excel = []
        for fila in filas[1:]:
            if fila[idx_nombre] is None or str(fila[idx_nombre]).strip()=="": continue
            grado = str(fila[idx_grado]).strip() if fila[idx_grado] else ""
            nombre = str(fila[idx_nombre]).strip()
            try: num = int(fila[idx_num]) if idx_num is not None and fila[idx_num] is not None else 0
            except: num = 0
            datos_excel.append((grado, num, nombre))
        if not datos_excel:
            messagebox.showwarning("Sin datos", "No encontré alumnos en el Excel"); return False
        if not messagebox.askyesno("¿Borrar todo?", f"Se van a BORRAR todos los alumnos actuales y se van a cargar {len(datos_excel)} alumnos nuevos del Excel.\n¿Estás seguro?"):
            return False
        con = conectar(); cur = con.cursor()
        cur.execute("DELETE FROM asistencia"); cur.execute("DELETE FROM alumnos")
        contadores = {}
        for grado, num, nombre in datos_excel:
            cur.execute("SELECT id FROM grados WHERE nombre=%s", (grado,))
            r = cur.fetchone()
            if r: gid = r[0]
            else:
                cur.execute("INSERT INTO grados (nombre, total_alumnos) VALUES (%s,0)", (grado,)); gid = cur.lastrowid
            if num == 0:
                contadores[grado] = contadores.get(grado, 0) + 1; num = contadores[grado]
            else:
                cur.execute("SELECT id FROM alumnos WHERE grado_id=%s AND numero_lista=%s", (gid, num))
                if cur.fetchone():
                    contadores[grado] = contadores.get(grado, 0) + 1; num = 1000 + contadores[grado]
            cur.execute("INSERT INTO alumnos (grado_id, numero_lista, nombre_completo) VALUES (%s,%s,%s)", (gid, num, nombre))
        cur.execute("UPDATE grados g SET total_alumnos = (SELECT COUNT(*) FROM alumnos a WHERE a.grado_id=g.id)")
        con.commit(); con.close(); cargar_datos_db()
        messagebox.showinfo("Listo", f"Carga masiva exitosa: {len(datos_excel)} alumnos cargados.")
        if lbl_status: lbl_status.config(text=f"✅ {len(datos_excel)} alumnos cargados")
        return True
    except Exception as e:
        messagebox.showerror("Error Excel", str(e)); return False

def agregar_un_alumno(nombre, grado, numero_lista):
    nombre = nombre.strip(); grado = grado.strip()
    if not nombre or not grado:
        messagebox.showwarning("Falta dato", "Escribe nombre y grado"); return False
    try:
        con = conectar(); cur = con.cursor()
        cur.execute("SELECT id FROM grados WHERE nombre=%s", (grado,)); r = cur.fetchone()
        if r: gid = r[0]
        else:
            cur.execute("INSERT INTO grados (nombre, total_alumnos) VALUES (%s,0)", (grado,)); gid = cur.lastrowid
        if numero_lista == 0 or numero_lista == "":
            cur.execute("SELECT COALESCE(MAX(numero_lista),0)+1 FROM alumnos WHERE grado_id=%s", (gid,)); numero_lista = cur.fetchone()[0]
        else:
            numero_lista = int(numero_lista)
            cur.execute("SELECT id FROM alumnos WHERE grado_id=%s AND numero_lista=%s", (gid, numero_lista))
            if cur.fetchone():
                messagebox.showerror("Duplicado", f"Ya existe el número {numero_lista} en {grado}"); con.close(); return False
        cur.execute("INSERT INTO alumnos (grado_id, numero_lista, nombre_completo) VALUES (%s,%s,%s)", (gid, numero_lista, nombre))
        cur.execute("UPDATE grados SET total_alumnos=(SELECT COUNT(*) FROM alumnos WHERE grado_id=%s) WHERE id=%s", (gid, gid))
        con.commit(); con.close(); cargar_datos_db()
        messagebox.showinfo("Guardado", f"Alumno {nombre} agregado a {grado} con número {numero_lista}.")
        return True
    except Exception as e:
        messagebox.showerror("Error DB", str(e)); return False

GUINDA = "#4E0F27"; DORADO = "#C9A86A"; BEIGE = "#F5E6D3"; BLANCO = "#FDF8F0"; BEIGE_CLARO = "#F5F0E6"; BORDER = "#D9C5A0"
GRADOS = cargar_datos_db()
GRADO_ACTUAL = None; ESTADOS = {}; ESTADOS_INC = {}
USUARIO_ACTUAL = {"nombre": "", "usuario": "", "rol": "", "grupo": ""}

# ===== PUNTO 2: DICCIONARIO BILINGÜE PDF =====
CARPETA_DICC = "diccionario_bilingue"
ARCHIVO_DICC = os.path.join(CARPETA_DICC, "diccionario.json")
os.makedirs(CARPETA_DICC, exist_ok=True)

def cargar_lista_diccionario():
    if not os.path.exists(ARCHIVO_DICC): return []
    try:
        with open(ARCHIVO_DICC, "r", encoding="utf-8") as f: return json.load(f)
    except: return []

def guardar_lista_diccionario(lista):
    with open(ARCHIVO_DICC, "w", encoding="utf-8") as f: json.dump(lista, f, ensure_ascii=False, indent=2)

def crear_vista_diccionario(parent):
    tk.Label(parent, text="📖 Diccionario Bilingüe", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=25, pady=(20,2))
    tk.Label(parent, text=f"Subido por: {USUARIO_ACTUAL['nombre']} ({USUARIO_ACTUAL['rol']}) - Todos los docentes y el director pueden ver y agregar PDFs.", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11), wraplength=800, justify="left").pack(anchor="w", padx=25, pady=(0,10))
    card = tk.Frame(parent, bg=BLANCO, highlightbackground="#EADDC8", highlightthickness=1, padx=20, pady=15)
    card.pack(fill="x", padx=25, pady=5)
    tk.Label(card, text="Nombre para mostrar en la plataforma (ej. Diccionario Español - Náhuatl):", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w")
    entry_titulo = tk.Entry(card, font=("Segoe UI", 12)); entry_titulo.pack(fill="x", ipady=8, pady=6)
    entry_titulo.insert(0, "Diccionario Bilingüe")
    lbl_sel = tk.Label(card, text="Ningún PDF seleccionado", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 10)); lbl_sel.pack(anchor="w")
    ruta_sel = {"ruta": ""}
    def elegir_pdf():
        ruta = filedialog.askopenfilename(title="Elige el PDF del diccionario", filetypes=[("PDF", "*.pdf")])
        if ruta:
            ruta_sel["ruta"] = ruta
            lbl_sel.config(text=f"Seleccionado: {os.path.basename(ruta)}")
            if entry_titulo.get().strip() in ("", "Diccionario Bilingüe"):
                entry_titulo.delete(0, "end"); entry_titulo.insert(0, os.path.splitext(os.path.basename(ruta))[0])
    tk.Button(card, text="📂 ELEGIR PDF", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 11, "bold"), bd=0, pady=10, command=elegir_pdf).pack(fill="x", pady=5)
    # Lista
    tk.Label(parent, text="  PDFs guardados en la plataforma:", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=25, pady=(12,5))
    lista_frame = tk.Frame(parent, bg=BLANCO); lista_frame.pack(fill="both", expand=True, padx=25, pady=5)
    tree = ttk.Treeview(lista_frame, columns=("Titulo","Subido por","Fecha","Archivo"), show="headings", height=8)
    tree.heading("Titulo", text="Título"); tree.heading("Subido por", text="Subido por"); tree.heading("Fecha", text="Fecha"); tree.heading("Archivo", text="Archivo")
    tree.column("Titulo", width=260); tree.column("Subido por", width=180); tree.column("Fecha", width=140); tree.column("Archivo", width=200)
    tree.pack(side="left", fill="both", expand=True)
    sb = tk.Scrollbar(lista_frame, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=sb.set); sb.pack(side="right", fill="y")
    def refrescar():
        for i in tree.get_children(): tree.delete(i)
        for d in cargar_lista_diccionario():
            tree.insert("", "end", values=(d.get("titulo",""), d.get("subido_por",""), d.get("fecha",""), d.get("archivo","")))
    def agregar_pdf():
        if not ruta_sel["ruta"]:
            messagebox.showwarning("Falta PDF", "Primero elige un PDF"); return
        titulo = entry_titulo.get().strip() or "Diccionario Bilingüe"
        nombre_archivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.path.basename(ruta_sel['ruta'])}"
        destino = os.path.join(CARPETA_DICC, nombre_archivo)
        try:
            shutil.copy(ruta_sel["ruta"], destino)
            lista = cargar_lista_diccionario()
            lista.append({"titulo": titulo, "archivo": nombre_archivo, "subido_por": f"{USUARIO_ACTUAL['nombre']} ({USUARIO_ACTUAL['rol']})", "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"), "ruta": destino})
            guardar_lista_diccionario(lista)
            refrescar()
            ruta_sel["ruta"] = ""; lbl_sel.config(text="Ningún PDF seleccionado"); entry_titulo.delete(0,"end"); entry_titulo.insert(0, "Diccionario Bilingüe")
            messagebox.showinfo("Guardado", f"PDF '{titulo}' agregado. Ya lo pueden ver docentes y director.")
        except Exception as e: messagebox.showerror("Error", str(e))
    def abrir_pdf():
        sel = tree.selection()
        if not sel: messagebox.showwarning("Elige uno", "Selecciona un PDF de la lista"); return
        vals = tree.item(sel[0])["values"]
        # Buscar ruta real
        for d in cargar_lista_diccionario():
            if d.get("archivo")==vals[3]:
                ruta = d.get("ruta", os.path.join(CARPETA_DICC, d["archivo"]))
                if os.path.exists(ruta):
                    webbrowser.open(os.path.abspath(ruta)); return
        messagebox.showerror("No encontrado", "El archivo ya no existe en la carpeta.")
    def eliminar_pdf():
        sel = tree.selection()
        if not sel: messagebox.showwarning("Elige uno", "Selecciona un PDF para eliminar"); return
        if not messagebox.askyesno("Eliminar", "¿Seguro que quieres eliminar este PDF de la plataforma?"): return
        vals = tree.item(sel[0])["values"]
        lista = cargar_lista_diccionario()
        nueva = []
        for d in lista:
            if d.get("archivo")==vals[3]:
                try: os.remove(d.get("ruta", os.path.join(CARPETA_DICC, d["archivo"])))
                except: pass
            else: nueva.append(d)
        guardar_lista_diccionario(nueva); refrescar()
        messagebox.showinfo("Eliminado", "PDF eliminado.")
    btns = tk.Frame(parent, bg=BLANCO); btns.pack(fill="x", padx=25, pady=10)
    tk.Button(btns, text="➕ AGREGAR PDF A LA PLATAFORMA", bg="#611232", fg="white", font=("Segoe UI", 11, "bold"), bd=0, pady=10, command=agregar_pdf).pack(side="left", fill="x", expand=True, padx=3)
    tk.Button(btns, text="👁 ABRIR / DESCARGAR", bg=DORADO, fg=GUINDA, font=("Segoe UI", 11, "bold"), bd=0, pady=10, command=abrir_pdf).pack(side="left", fill="x", expand=True, padx=3)
    tk.Button(btns, text="🗑 ELIMINAR", bg="#B22222", fg="white", font=("Segoe UI", 11, "bold"), bd=0, pady=10, command=eliminar_pdf).pack(side="left", fill="x", expand=True, padx=3)
    tk.Label(parent, text="Tip: El PDF queda guardado en la carpeta 'diccionario_bilingue' junto a tu programa. Si copias esa carpeta, el diccionario se va contigo.", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 9)).pack(anchor="w", padx=25, pady=5)
    refrescar()

# ===== MOTOR DE ALERTAS - PUNTO 1 - CAMPANITA =====
ALERTAS = []; FILTRO_ALERTA = ["todas"]; RENDER_CAMPANITA = [None]
def cargar_alertas_db():
    global ALERTAS
    try:
        con = conectar(); cur = con.cursor(dictionary=True)
        cur.execute("SELECT * FROM reportes_alerta ORDER BY fecha DESC LIMIT 20")
        rows = cur.fetchall(); con.close(); ALERTAS = []
        for r in rows:
            t = str(r.get('tipo',''))
            if "Retardo" in t: tipo = "retardo"
            elif "tarea" in t.lower(): tipo = "tarea"
            else: tipo = "falta"
            ALERTAS.append({"tipo": tipo, "alumno": f"{r.get('alumno_nombre','')} - {r.get('grado_grupo','')}", "texto": f"{r.get('tipo','')} | Riesgo {r.get('nivel_riesgo','')} | {r.get('descripcion','')}", "hora": str(r.get('fecha','')), "leida": False})
        if not ALERTAS:
            ALERTAS = [
                {"tipo":"falta","alumno":"María López - 3° A","texto":"Falta 3 días seguidos sin justificante","hora":"Hace 10 min","leida":False},
                {"tipo":"retardo","alumno":"Carlos Ruiz - 4° B","texto":"Retardo por 4ta vez esta semana","hora":"Hace 32 min","leida":False},
                {"tipo":"tarea","alumno":"Ana Torres - 5° A","texto":"No entregó tarea de matemáticas","hora":"Hace 1 hora","leida":False},
            ]
    except:
        ALERTAS = [
            {"tipo":"falta","alumno":"María López - 3° A","texto":"Falta 3 días seguidos","hora":"Hace 10 min","leida":False},
            {"tipo":"retardo","alumno":"Carlos Ruiz - 4° B","texto":"Retardo 4ta vez","hora":"Hace 32 min","leida":False},
            {"tipo":"tarea","alumno":"Ana Torres - 5° A","texto":"No entregó tarea","hora":"Hace 1 hora","leida":False},
        ]
def color_por_tipo(tipo):
    if tipo == "falta": return "#e53935"
    if tipo == "retardo": return "#FFC107"
    return "#FF9800"
def agregar_alerta_nueva(tipo, alumno, texto):
    ALERTAS.insert(0, {"tipo":tipo, "alumno":alumno, "texto":texto, "hora":"Ahora mismo", "leida":False})
    if RENDER_CAMPANITA[0]:
        try: RENDER_CAMPANITA[0]()
        except: pass
def crear_campanita(parent_header):
    cargar_alertas_db()
    frame = tk.Frame(parent_header, bg="#611232"); frame.pack(side="right", padx=10)
    lbl_cont = tk.Label(frame, text="3", bg="#e53935", fg="white", font=("Segoe UI", 8, "bold"), width=2); lbl_cont.place(x=28, y=0)
    panel = tk.Frame(ventana, bg=BLANCO, highlightbackground=DORADO, highlightthickness=2, width=380, height=380)
    titulo = tk.Frame(panel, bg=GUINDA); titulo.pack(fill="x")
    tk.Label(titulo, text="  🔔 Alertas", bg=GUINDA, fg="white", font=("Segoe UI", 11, "bold")).pack(side="left", pady=8)
    def marcar_todas():
        for a in ALERTAS: a["leida"] = True
        render()
    tk.Button(titulo, text="Marcar todas leídas", bg=DORADO, fg=GUINDA, font=("Segoe UI", 8, "bold"), relief="flat", command=marcar_todas).pack(side="right", padx=5)
    filtros = tk.Frame(panel, bg=BLANCO); filtros.pack(fill="x", pady=5)
    btns = {}
    def filtrar(t):
        FILTRO_ALERTA[0] = t
        for k,b in btns.items(): b.config(bg=GUINDA if k==t else "#f0f0f0", fg="white" if k==t else "black")
        render()
    for txt, t in [("Todas","todas"),("🔴 Falta","falta"),("🟡 Retardo","retardo"),("🟠 Tarea","tarea")]:
        b = tk.Button(filtros, text=txt, bg="#f0f0f0", font=("Segoe UI", 8), relief="flat", command=lambda tt=t: filtrar(tt)); b.pack(side="left", padx=3); btns[t] = b
    lista_frame = tk.Frame(panel, bg=BLANCO); lista_frame.pack(fill="both", expand=True)
    def render():
        for w in lista_frame.winfo_children(): w.destroy()
        filtradas = [a for a in ALERTAS if FILTRO_ALERTA[0]=="todas" or a["tipo"]==FILTRO_ALERTA[0]]
        if not filtradas: tk.Label(lista_frame, text="No hay alertas de este tipo ✅", bg=BLANCO, fg="gray").pack(pady=20)
        for a in filtradas:
            f = tk.Frame(lista_frame, bg="#FFF3E0" if not a["leida"] else BLANCO, bd=1, relief="solid"); f.pack(fill="x", padx=8, pady=4)
            tk.Label(f, text="●", fg=color_por_tipo(a["tipo"]), bg=f["bg"], font=("Arial", 14)).pack(side="left", padx=5)
            lbl = tk.Label(f, text=f"{a['alumno']}\n{a['texto']}\n{a['hora']}", bg=f["bg"], font=("Segoe UI", 9), justify="left", wraplength=300); lbl.pack(side="left", fill="x", expand=True, pady=5)
            def abrir(al=a):
                al["leida"] = True; render(); messagebox.showinfo("Alerta", f"{al['alumno']}\n{al['texto']}")
            f.bind("<Button-1>", lambda e, al=a: abrir(al)); lbl.bind("<Button-1>", lambda e, al=a: abrir(al))
        sin_leer = len([a for a in ALERTAS if not a["leida"]]); lbl_cont.config(text=str(sin_leer))
        if sin_leer == 0: lbl_cont.place_forget()
        else: lbl_cont.place(x=28, y=0)
    RENDER_CAMPANITA[0] = render; visible = [False]
    def toggle():
        if visible[0]: panel.place_forget(); visible[0] = False
        else: panel.place(x=ventana.winfo_width()-400, y=65, width=380, height=380); visible[0] = True
    tk.Button(frame, text="🔔", bg="white", font=("Arial", 16), relief="flat", width=2, command=toggle).pack()
    render(); return frame

DOCENTES_RAPIDOS = [
    ("👩‍🏫  Profa. Ana  •  1° Primaria", "ana.1a", "ANA-1A-2026"),
    ("👨‍🏫  Profe. Luis  •  2° Primaria", "luis.2b", "LUIS-2B-2026"),
    ("👩‍🏫  Profa. María  •  3° Primaria", "maria.3a", "MARIA-3A-2026"),
    ("👨‍🏫  Profe. José  •  4° Primaria", "jose.4a", "JOSE-4A-2026"),
    ("👨‍🏫  Profe. Pedro  •  5° Primaria", "pedro.5a", "PEDRO-5A-2026"),
    ("👩‍🏫  Profa. Sofía  •  6° Primaria", "sofia.6a", "SOFIA-6A-2026"),
]

class AutocompleteEntry(tk.Entry):
    def __init__(self, master, lista_nombres=[], **kw):
        super().__init__(master, **kw)
        self.lista = lista_nombres; self.listbox = None
        self.bind("<KeyRelease>", self._filtrar)
    def _filtrar(self, e):
        if e.keysym in ("Up","Down","Return"): return
        texto=self.get().lower()
        if self.listbox:
            try: self.listbox.destroy()
            except: pass
            self.listbox=None
        if len(texto)<2: return
        matches=[n for n in self.lista if texto in n.lower()]
        if not matches: return
        self.listbox=tk.Listbox(self.master, bg=BLANCO, fg=GUINDA, relief="solid", bd=1, height=min(5,len(matches)), font=("Arial",11))
        self.listbox.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height(), width=self.winfo_width())
        for m in matches: self.listbox.insert("end", m)
        self.listbox.bind("<<ListboxSelect>>", self._elegir); self.listbox.lift()
    def _elegir(self, e):
        if self.listbox and self.listbox.curselection():
            self.delete(0,"end"); self.insert(0, self.listbox.get(self.listbox.curselection())); self.listbox.destroy(); self.listbox=None

def limpiar():
    for w in ventana.winfo_children(): w.destroy()
def crear_header_login():
    header = tk.Frame(ventana, bg="#611232", height=55); header.pack(fill="x"); header.pack_propagate(False)
    tk.Label(header, text="PRIMARIA NEZAHUALCÓYOTL", font=("Segoe UI", 12, "bold"), bg="#611232", fg="white").pack(side="left", padx=20, pady=10)
    tk.Label(header, text="Alerta Temprana", font=("Segoe UI", 9), bg="#611232", fg=DORADO).pack(side="right", padx=20)
    tk.Frame(ventana, bg=DORADO, height=2).pack(fill="x")
def login():
    if not entrada_user.get() or not entrada_pass.get():
        messagebox.showwarning("Aviso", "Escribe usuario y contraseña"); return
    u = entrada_user.get().strip().lower(); p = entrada_pass.get().strip()
    try:
        con = conectar(); cur = con.cursor(dictionary=True)
        cur.execute("SELECT * FROM usuarios WHERE usuario=%s AND password=%s", (u, p)); user = cur.fetchone(); con.close()
    except Exception as e:
        messagebox.showerror("Error DB", f"No conecta a XAMPP: {e}"); return
    if user:
        USUARIO_ACTUAL["nombre"] = user.get('nombre', user['usuario']); USUARIO_ACTUAL["usuario"] = user['usuario']; USUARIO_ACTUAL["rol"] = user['rol']; USUARIO_ACTUAL["grupo"] = user.get('grupo_asignado') or ""
        vista_directivo() if user['rol']=="director" else vista_docente()
    else: messagebox.showerror("Error", "Usuario o contraseña incorrecta")
def pantalla_login():
    limpiar(); crear_header_login()
    fondo = tk.Frame(ventana, bg=BEIGE); fondo.pack(expand=True, fill="both")
    canvas = tk.Canvas(fondo, bg=BEIGE, highlightthickness=0); scrollbar = tk.Scrollbar(fondo, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=BEIGE); scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    win_id = canvas.create_window((0,0), window=scroll_frame, anchor="n"); canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width)); canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
    card = tk.Frame(scroll_frame, bg=BLANCO, highlightbackground=DORADO, highlightthickness=1, padx=35, pady=25); card.pack(padx=20, pady=20)
    tk.Label(card, text="INICIAR SESIÓN", font=("Segoe UI", 14, "bold"), bg=BLANCO, fg="#611232").pack(pady=(0,15))
    tk.Label(card, text="Usuario:", bg=BLANCO, font=("Segoe UI", 11), anchor="w").pack(fill="x")
    global entrada_user, entrada_pass
    entrada_user = tk.Entry(card, font=("Segoe UI", 11), bg="#FFFBF5", bd=0, highlightthickness=1, highlightbackground="#D9C5A0", width=38); entrada_user.pack(fill="x", ipady=10, pady=5)
    tk.Label(card, text="Contraseña:", bg=BLANCO, font=("Segoe UI", 11), anchor="w").pack(fill="x")
    entrada_pass = tk.Entry(card, font=("Segoe UI", 11), show="•", bg="#FFFBF5", bd=0, highlightthickness=1, highlightbackground="#D9C5A0"); entrada_pass.pack(fill="x", ipady=10, pady=5)
    tk.Label(card, text="director / DIR123", font=("Segoe UI", 8), bg=BLANCO, fg=GUINDA).pack(pady=5)
    tk.Button(card, text="ENTRAR", bg="#611232", fg="white", font=("Segoe UI", 12, "bold"), bd=0, pady=12, cursor="hand2", command=login).pack(fill="x", pady=12)
    tk.Label(card, text="Acceso rápido - toca tu nombre:", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 8)).pack(pady=(2,6))
    def llenar(u,p):
        entrada_user.delete(0,"end"); entrada_user.insert(0,u); entrada_pass.delete(0,"end"); entrada_pass.insert(0,p)
    for texto, u, p in DOCENTES_RAPIDOS:
        tk.Button(card, text=f"{texto}\n{u}  |  {p}", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 9), bd=0, anchor="w", justify="left", cursor="hand2", command=lambda u=u,p=p: llenar(u,p)).pack(fill="x", pady=2, ipady=3)
    entrada_pass.bind("<Return>", lambda e: login())
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

def vista_docente():
    limpiar()
    global GRADO_ACTUAL, ESTADOS, ESTADOS_INC
    grupo_inicial = USUARIO_ACTUAL["grupo"] if USUARIO_ACTUAL["grupo"] in GRADOS else "1° Primaria"
    GRADO_ACTUAL = tk.StringVar(value=grupo_inicial); ESTADOS = {}; ESTADOS_INC = {}
    for g in GRADOS:
        ESTADOS[g] = {a: tk.StringVar(value="Asistio") for a in GRADOS[g]}; ESTADOS_INC[g] = {a: tk.BooleanVar(value=False) for a in GRADOS[g]}
    header_top = tk.Frame(ventana, bg="#611232", height=55); header_top.pack(side="top", fill="x"); header_top.pack_propagate(False)
    tk.Label(header_top, text=f"  {USUARIO_ACTUAL['nombre']} • {USUARIO_ACTUAL['grupo']}", bg="#611232", fg="white", font=("Segoe UI", 11, "bold")).pack(side="left", pady=12)
    crear_campanita(header_top)
    menu = tk.Frame(ventana, bg=GUINDA, width=290); menu.pack(side="left", fill="y"); menu.pack_propagate(False)
    tk.Label(menu, text="DOCENTE\nPANEL", bg=GUINDA, fg=BLANCO, font=("Segoe UI", 18, "bold"), justify="left").pack(anchor="w", padx=20, pady=20)
    tk.Frame(menu, bg=DORADO, height=3, width=60).pack(anchor="w", padx=20, pady=5)
    tk.Label(menu, text=f"{USUARIO_ACTUAL['nombre']}\n{USUARIO_ACTUAL['grupo']}", bg=GUINDA, fg=DORADO, font=("Segoe UI", 10, "bold"), anchor="w", justify="left").pack(anchor="w", padx=20, pady=(5,0))
    contenido = tk.Frame(ventana, bg=BLANCO); contenido.pack(side="right", fill="both", expand=True)

    def crear_vista_justificante(parent):
        nombres = cargar_nombres_alumnos()
        alumno_var = tk.StringVar(); grado_var = tk.StringVar(value=USUARIO_ACTUAL["grupo"] or "3° A"); docente_var = tk.StringVar(value=USUARIO_ACTUAL['nombre'])
        fecha_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y")); motivo_var = tk.StringVar(value="Enfermedad"); faltas_var = tk.StringVar(value="1"); tutor_var = tk.StringVar(); telefono_var = tk.StringVar()
        canvas = tk.Canvas(parent, bg="#EDE6D9", highlightthickness=0); scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#EDE6D9"); cw = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width)); scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        contenedor = tk.Frame(scrollable_frame, bg="#FFF9F0", highlightbackground=BORDER, highlightthickness=1); contenedor.pack(fill="both", expand=True, padx=30, pady=20)
        tk.Label(contenedor, text="FORMATO DE JUSTIFICANTE", font=("Arial", 16, "bold"), bg="#FFF9F0", fg="#611232").pack(pady=12)
        f1 = tk.LabelFrame(contenedor, text=" Datos del alumno ", font=("Arial", 12, "bold"), bg=BLANCO, fg="#611232", padx=15, pady=10, bd=0, highlightbackground=BORDER, highlightthickness=1); f1.pack(fill="x", padx=15, pady=8); f1.columnconfigure(1, weight=1)
        tk.Label(f1, text="Nombre del alumno:", bg=BLANCO, font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=8)
        entry_alumno = AutocompleteEntry(f1, nombres, font=("Arial", 12)); entry_alumno.grid(row=0, column=1, sticky="ew", padx=10, ipady=8)
        tk.Label(f1, text="Grado y grupo:", bg=BLANCO, font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=8)
        ttk.Combobox(f1, textvariable=grado_var, values=["1° A","1° B","2° A","2° B","3° A","3° B","4° A","4° B","5° A","5° B","6° A","6° B"], font=("Arial", 12), state="readonly").grid(row=1, column=1, sticky="ew", padx=10)
        tk.Label(f1, text="Docente:", bg=BLANCO, font=("Arial", 12)).grid(row=2, column=0, sticky="w", pady=8)
        tk.Entry(f1, textvariable=docente_var, font=("Arial", 12), state="disabled", disabledbackground="#FFFBF5", disabledforeground=GUINDA).grid(row=2, column=1, sticky="ew", padx=10, ipady=8)
        f2 = tk.LabelFrame(contenedor, text=" Motivo de la inasistencia ", font=("Arial", 12, "bold"), bg=BLANCO, fg="#611232", padx=15, pady=10, bd=0, highlightbackground=BORDER, highlightthickness=1); f2.pack(fill="x", padx=15, pady=8); f2.columnconfigure(1, weight=1)
        tk.Label(f2, text="Fecha de falta:", bg=BLANCO, font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=8); tk.Entry(f2, textvariable=fecha_var, font=("Arial", 12), width=18).grid(row=0, column=1, sticky="w", padx=10, ipady=6)
        tk.Label(f2, text="No. de días:", bg=BLANCO, font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=8); tk.Spinbox(f2, from_=1, to=30, textvariable=faltas_var, font=("Arial", 12), width=10).grid(row=1, column=1, sticky="w", padx=10)
        tk.Label(f2, text="Motivo:", bg=BLANCO, font=("Arial", 12)).grid(row=2, column=0, sticky="w", pady=8); ttk.Combobox(f2, textvariable=motivo_var, values=["Enfermedad","Cita médica","Asunto familiar","Trámite","Otro"], font=("Arial", 12), state="readonly").grid(row=2, column=1, sticky="ew", padx=10)
        tk.Label(f2, text="Descripción:", bg=BLANCO, font=("Arial", 12)).grid(row=3, column=0, sticky="nw", pady=8); descripcion = tk.Text(f2, height=4, font=("Arial", 12)); descripcion.grid(row=3, column=1, sticky="ew", padx=10)
        f3 = tk.LabelFrame(contenedor, text=" Datos del tutor ", font=("Arial", 12, "bold"), bg=BLANCO, fg="#611232", padx=15, pady=10, bd=0, highlightbackground=BORDER, highlightthickness=1); f3.pack(fill="x", padx=15, pady=8); f3.columnconfigure(1, weight=1)
        tk.Label(f3, text="Nombre del tutor:", bg=BLANCO, font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=8); tk.Entry(f3, textvariable=tutor_var, font=("Arial", 12)).grid(row=0, column=1, sticky="ew", padx=10, ipady=8)
        tk.Label(f3, text="Teléfono:", bg=BLANCO, font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=8)
        def solo_numero(P): return P.isdigit() or P == ""
        vcmd = ventana.register(solo_numero); tk.Entry(f3, textvariable=telefono_var, font=("Arial", 12), validate="key", validatecommand=(vcmd, '%P')).grid(row=1, column=1, sticky="ew", padx=10, ipady=8)
        def limpiar_j():
            entry_alumno.delete(0,"end"); alumno_var.set(""); tutor_var.set(""); telefono_var.set(""); descripcion.delete("1.0", "end"); fecha_var.set(datetime.now().strftime("%d/%m/%Y")); faltas_var.set("1"); motivo_var.set("Enfermedad")
        def generar():
            alumno_var.set(entry_alumno.get())
            if not alumno_var.get().strip(): messagebox.showwarning("Falta dato", "Escribe el nombre del alumno"); return
            if not tutor_var.get().strip(): messagebox.showwarning("Falta dato", "Escribe el nombre del tutor"); return
            try:
                fecha_db = datetime.strptime(fecha_var.get().strip(), "%d/%m/%Y").date(); con = conectar(); cur = con.cursor()
                cur.execute("INSERT INTO justificantes (alumno_nombre, grado_grupo, docente, fecha_falta, num_dias, motivo, descripcion, tutor, telefono) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)", (alumno_var.get(), grado_var.get(), docente_var.get(), fecha_db, int(faltas_var.get()), motivo_var.get(), descripcion.get("1.0","end").strip(), tutor_var.get(), telefono_var.get()))
                con.commit(); con.close(); messagebox.showinfo("Guardado", "Justificante guardado en la base de datos")
            except Exception as e: messagebox.showerror("Error DB", str(e))
        btns = tk.Frame(contenedor, bg="#FFF9F0"); btns.pack(pady=12)
        tk.Button(btns, text="LIMPIAR", bg=BLANCO, fg="#611232", font=("Arial", 12, "bold"), padx=20, pady=10, command=limpiar_j).pack(side="left", padx=10)
        tk.Button(btns, text="GENERAR JUSTIFICANTE", bg="#611232", fg="white", font=("Arial", 13, "bold"), padx=30, pady=11, command=generar).pack(side="left", padx=10)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    def crear_vista_reportes(parent):
        nombres = cargar_nombres_alumnos()
        alumno_var = tk.StringVar(); grado_var = tk.StringVar(value=USUARIO_ACTUAL["grupo"]); docente_var = tk.StringVar(value=USUARIO_ACTUAL['nombre'])
        fecha_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y")); tipo_var = tk.StringVar(value="Inasistencia"); riesgo_var = tk.StringVar(value="Medio"); apercibimiento_var = tk.BooleanVar(); apoyo_var = tk.BooleanVar()
        canvas = tk.Canvas(parent, bg="#EDE6D9", highlightthickness=0); scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#EDE6D9"); cw = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width)); scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        contenedor = tk.Frame(scrollable_frame, bg="#FFF8FA", highlightbackground=BORDER, highlightthickness=1); contenedor.pack(fill="both", expand=True, padx=30, pady=20)
        f1 = tk.LabelFrame(contenedor, text=" Datos del alumno en riesgo ", font=("Arial", 11, "bold"), bg=BLANCO, fg="#611232", padx=15, pady=10, bd=0, highlightbackground=BORDER, highlightthickness=1); f1.pack(fill="x", pady=8, padx=15); f1.columnconfigure(1, weight=1)
        tk.Label(f1, text="Alumno:", bg=BLANCO, font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=6); entry_riesgo = AutocompleteEntry(f1, nombres, font=("Arial", 11)); entry_riesgo.grid(row=0, column=1, sticky="ew", padx=10, pady=4, ipady=7)
        tk.Label(f1, text="Grado y grupo:", bg=BLANCO, font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=6); ttk.Combobox(f1, textvariable=grado_var, values=["1° A","1° B","2° A","2° B","3° A","3° B","4° A","4° B","5° A","5° B","6° A","6° B"], font=("Arial", 11), state="readonly").grid(row=1, column=1, sticky="ew", padx=10, pady=4)
        tk.Label(f1, text="Docente:", bg=BLANCO, font=("Arial", 11)).grid(row=2, column=0, sticky="w", pady=6); tk.Entry(f1, textvariable=docente_var, font=("Arial", 11), state="disabled", disabledbackground="#FFFBF5", disabledforeground=GUINDA).grid(row=2, column=1, sticky="ew", padx=10, pady=4, ipady=7)
        f2 = tk.LabelFrame(contenedor, text=" Detalle de la incidencia ", font=("Arial", 11, "bold"), bg=BLANCO, fg="#611232", padx=15, pady=10, bd=0, highlightbackground=BORDER, highlightthickness=1); f2.pack(fill="x", pady=8, padx=15); f2.columnconfigure(1, weight=1)
        tk.Label(f2, text="Fecha:", bg=BLANCO, font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=6); tk.Entry(f2, textvariable=fecha_var, font=("Arial", 11), width=18).grid(row=0, column=1, sticky="w", padx=10, pady=4, ipady=6)
        tk.Label(f2, text="Tipo:", bg=BLANCO, font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=6); ttk.Combobox(f2, textvariable=tipo_var, values=["Inasistencia","Retardo","No entrega tareas","Conducta","Otro"], font=("Arial", 11), state="readonly").grid(row=1, column=1, sticky="ew", padx=10, pady=4)
        tk.Label(f2, text="Nivel de riesgo:", bg=BLANCO, font=("Arial", 11)).grid(row=2, column=0, sticky="w", pady=6); riesgo_frame = tk.Frame(f2, bg=BLANCO); riesgo_frame.grid(row=2, column=1, sticky="w", padx=10)
        for r in ["Bajo", "Medio", "Alto"]: tk.Radiobutton(riesgo_frame, text=r, variable=riesgo_var, value=r, bg=BLANCO, font=("Arial", 11)).pack(side="left", padx=10)
        tk.Label(f2, text="Descripción:", bg=BLANCO, font=("Arial", 11)).grid(row=3, column=0, sticky="nw", pady=6); descripcion = tk.Text(f2, height=3, font=("Arial", 11)); descripcion.grid(row=3, column=1, sticky="ew", padx=10, pady=4)
        tk.Label(f2, text="Acción propuesta:", bg=BLANCO, font=("Arial", 11)).grid(row=4, column=0, sticky="nw", pady=6); accion = tk.Text(f2, height=3, font=("Arial", 11)); accion.grid(row=4, column=1, sticky="ew", padx=10, pady=4)
        f3 = tk.Frame(contenedor, bg=BLANCO, highlightbackground=BORDER, highlightthickness=1); f3.pack(fill="x", pady=10, padx=15)
        tk.Checkbutton(f3, text="Generar Apercibimiento", variable=apercibimiento_var, bg=BLANCO, font=("Arial", 11)).pack(side="left", padx=15, pady=10)
        tk.Checkbutton(f3, text="Generar Plan de recuperación", variable=apoyo_var, bg=BLANCO, font=("Arial", 11)).pack(side="left", padx=15, pady=10)
        def limpiar_r():
            entry_riesgo.delete(0,"end"); alumno_var.set(""); grado_var.set(USUARIO_ACTUAL["grupo"]); tipo_var.set("Inasistencia"); riesgo_var.set("Medio"); fecha_var.set(datetime.now().strftime("%d/%m/%Y")); descripcion.delete("1.0", "end"); accion.delete("1.0", "end"); apercibimiento_var.set(False); apoyo_var.set(False)
        def generar_r():
            alumno_var.set(entry_riesgo.get())
            if not alumno_var.get().strip(): messagebox.showwarning("Falta dato", "Escribe el nombre del alumno"); return
            try:
                fecha_db = datetime.strptime(fecha_var.get().strip(), "%d/%m/%Y").date(); con = conectar(); cur = con.cursor()
                cur.execute("INSERT INTO reportes_alerta (alumno_nombre, grado_grupo, docente, fecha, tipo, nivel_riesgo, descripcion, accion_propuesta, generar_apercibimiento, generar_plan) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (alumno_var.get(), grado_var.get(), docente_var.get(), fecha_db, tipo_var.get(), riesgo_var.get(), descripcion.get("1.0","end").strip(), accion.get("1.0","end").strip(), int(apercibimiento_var.get()), int(apoyo_var.get())))
                con.commit(); con.close()
                tipo_camp = "falta" if tipo_var.get()=="Inasistencia" else "retardo" if tipo_var.get()=="Retardo" else "tarea"
                agregar_alerta_nueva(tipo_camp, f"{alumno_var.get()} - {grado_var.get()}", f"{tipo_var.get()} | Riesgo {riesgo_var.get()} | {descripcion.get('1.0','end').strip()}")
                messagebox.showinfo("Guardado", "Reporte guardado y alerta enviada 🔔")
            except Exception as e: messagebox.showerror("Error DB", str(e))
        botones = tk.Frame(contenedor, bg="#FFF8FA"); botones.pack(pady=18)
        tk.Button(botones, text="GENERAR REPORTE", bg="#611232", fg="white", font=("Arial", 12, "bold"), padx=25, pady=10, bd=0, command=generar_r).pack(side="left", padx=10)
        tk.Button(botones, text="LIMPIAR", bg=BLANCO, fg="#611232", font=("Arial", 11, "bold"), padx=20, pady=10, bd=1, relief="solid", command=limpiar_r).pack(side="left", padx=10)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    def crear_vista_asistencia(parent):
        header = tk.Frame(parent, bg=GUINDA); header.pack(fill="x")
        global titulo_grado, lista_frame
        titulo_grado = tk.Label(header, text=f"  {GRADO_ACTUAL.get()} - Lista Completa", bg=GUINDA, fg=BLANCO, font=("Segoe UI", 13, "bold"), pady=14, padx=15); titulo_grado.pack(side="left")
        tk.Label(header, text=f"{USUARIO_ACTUAL['nombre']} • {USUARIO_ACTUAL['grupo']}", bg=GUINDA, fg=DORADO, font=("Segoe UI", 10, "bold")).pack(side="right", padx=15)
        tk.Frame(parent, bg=DORADO, height=4).pack(fill="x")
        kpi_frame = tk.Frame(parent, bg=BEIGE_CLARO, padx=10, pady=10); kpi_frame.pack(fill="x")
        global l1_num, l2_num, l3_num, l4_num
        def kpi_card(titulo):
            f = tk.Frame(kpi_frame, bg=BLANCO, padx=10, pady=8, highlightbackground=DORADO, highlightthickness=1); f.pack(side="left", padx=5, expand=True, fill="x")
            tk.Label(f, text=titulo, bg=BLANCO, fg=GUINDA, font=("Segoe UI", 8, "bold")).pack(anchor="w"); lbl = tk.Label(f, text="0", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 18, "bold")); lbl.pack(anchor="w"); return lbl
        l1_num = kpi_card("ASISTIERON"); l2_num = kpi_card("FALTARON"); l3_num = kpi_card("RETARDOS"); l4_num = kpi_card("INCIDENCIAS")
        enc = tk.Frame(parent, bg=BEIGE, padx=10, pady=8); enc.pack(fill="x"); tk.Label(enc, text="NOMBRE DEL ALUMNO", bg=BEIGE, fg=GUINDA, font=("Segoe UI", 9, "bold"), width=40, anchor="w").pack(side="left")
        canvas_frame = tk.Frame(parent, bg=BLANCO); canvas_frame.pack(fill="both", expand=True)
        global canvas; canvas = tk.Canvas(canvas_frame, bg=BLANCO, highlightthickness=0)
        scroll = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview); canvas.configure(yscrollcommand=scroll.set); scroll.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        global lista_frame; lista_frame = tk.Frame(canvas, bg=BLANCO); win_id = canvas.create_window((0,0), window=lista_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width)); lista_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units")); cargar_grado()
        def guardar_asistencia():
            g = GRADO_ACTUAL.get(); hoy = datetime.now().date(); hora = datetime.now().strftime("%H:%M:%S")
            try:
                con = conectar(); cur = con.cursor()
                for nombre in GRADOS[g]:
                    aid = GRADO_IDS[g][nombre]; estado = ESTADOS[g][nombre].get(); inc = 1 if ESTADOS_INC[g][nombre].get() else 0
                    cur.execute("INSERT INTO asistencia (alumno_id, fecha, hora_llegada, estado, tiene_incidencia) VALUES (%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE estado=VALUES(estado), hora_llegada=VALUES(hora_llegada), tiene_incidencia=VALUES(tiene_incidencia)", (aid, hoy, hora, estado, inc))
                con.commit(); con.close(); messagebox.showinfo("Guardado", f"Asistencia de {g} guardada en la base de datos")
            except Exception as e: messagebox.showerror("Error DB", str(e))
        tk.Button(parent, text="GUARDAR ASISTENCIA EN BASE DE DATOS", bg=GUINDA, fg="white", font=("Segoe UI", 11, "bold"), pady=10, command=guardar_asistencia).pack(fill="x", padx=10, pady=10)
    def actualizar_kpi():
        g = GRADO_ACTUAL.get(); llego = sum(1 for n in GRADOS[g] if ESTADOS[g][n].get() == "Asistio"); falto = sum(1 for n in GRADOS[g] if ESTADOS[g][n].get() == "Falto"); ret = sum(1 for n in GRADOS[g] if ESTADOS[g][n].get() == "Retardo"); inc = sum(1 for n in GRADOS[g] if ESTADOS_INC[g][n].get() == True)
        l1_num.config(text=f"{llego}"); l2_num.config(text=f"{falto}"); l3_num.config(text=f"{ret}"); l4_num.config(text=f"{inc}"); titulo_grado.config(text=f"  {g} - Lista Completa")
    def cargar_grado():
        for w in lista_frame.winfo_children(): w.destroy()
        g = GRADO_ACTUAL.get()
        for i,nombre in enumerate(GRADOS[g]):
            bg_fila = BLANCO if i%2==0 else BEIGE_CLARO; fila = tk.Frame(lista_frame, bg=bg_fila); fila.pack(fill="x")
            tk.Label(fila, text=f"{i+1:02d}", bg=bg_fila, fg=GUINDA, font=("Segoe UI", 10, "bold"), width=4).pack(side="left", padx=(15,5), pady=9)
            tk.Label(fila, text=nombre.split(" ",1)[1], bg=bg_fila, fg=GUINDA, font=("Segoe UI", 11), width=35, anchor="w").pack(side="left")
            cont = tk.Frame(fila, bg=bg_fila); cont.pack(side="right", padx=15)
            for txt,val in [("Asistió","Asistio"),("Faltó","Falto"),("Retardo","Retardo")]:
                tk.Radiobutton(cont, text=txt, variable=ESTADOS[g][nombre], value=val, bg=bg_fila, fg=GUINDA, selectcolor=DORADO, font=("Segoe UI", 9, "bold"), indicatoron=0, width=8, command=actualizar_kpi).pack(side="left", padx=3)
            tk.Checkbutton(cont, text="Incidencia", variable=ESTADOS_INC[g][nombre], bg=bg_fila, fg="#8B0000", selectcolor="#FFD9D9", font=("Segoe UI", 9, "bold"), command=actualizar_kpi).pack(side="left", padx=10)
        actualizar_kpi()
    botones_docente = []
    def marcar_seleccionado(btn_activo):
        for b in botones_docente:
            if b == btn_activo: b.config(bg="#6B1636", fg=DORADO)
            else: b.config(bg=GUINDA, fg=BLANCO)
    def mostrar_opcion(op, btn=None):
        if btn: marcar_seleccionado(btn)
        for w in contenido.winfo_children(): w.destroy()
        if op == "Asistencia": crear_vista_asistencia(contenido)
        elif op == "Justificantes": crear_vista_justificante(contenido)
        elif op == "Reportes": crear_vista_reportes(contenido)
        elif op == "Diccionario": crear_vista_diccionario(contenido)
        elif op == "Mi Grupo":
            header = tk.Frame(contenido, bg=BLANCO); header.pack(fill="x", padx=30, pady=(25,10))
            tk.Label(header, text=f"Mi Grupo - {USUARIO_ACTUAL['grupo']}", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 22, "bold")).pack(anchor="w")
            tk.Label(header, text=f"Docente: {USUARIO_ACTUAL['nombre']} - Ciclo 2025-2026", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", pady=(5,15))
            cards = tk.Frame(contenido, bg=BLANCO); cards.pack(fill="x", padx=20, pady=10)
            card = tk.Frame(cards, bg=BLANCO, highlightbackground="#EADDC8", highlightthickness=1); card.pack(fill="x", padx=10, pady=10)
            tk.Frame(card, bg=GUINDA, height=6).pack(fill="x"); inner = tk.Frame(card, bg=BLANCO, padx=15, pady=15); inner.pack(fill="both", expand=True)
            tk.Label(inner, text=USUARIO_ACTUAL['grupo'], bg=BLANCO, fg=GUINDA, font=("Segoe UI", 16, "bold")).pack(anchor="w")
            tk.Label(inner, text=f"{len(GRADOS.get(USUARIO_ACTUAL['grupo'],[]))} alumnos", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 20, "bold")).pack(anchor="w", pady=(12,0))
            tk.Label(inner, text=USUARIO_ACTUAL['nombre'], bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w")
            tk.Button(inner, text="Ver lista →", bg=BEIGE_CLARO, fg=GUINDA, bd=0, font=("Segoe UI", 9, "bold"), cursor="hand2", pady=6, command=lambda: mostrar_opcion("Asistencia", botones_docente[1])).pack(fill="x", pady=(12,0))
        else: tk.Label(contenido, text=f"{op}\n\nMódulo en desarrollo", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 16, "bold")).pack(pady=50)
    def btn_menu(texto, cmd_op):
        b = tk.Button(menu, text=f"  {texto}", bg=GUINDA, fg=BLANCO, font=("Segoe UI", 11, "bold"), bd=0, anchor="w", pady=15, cursor="hand2", activebackground="#6B1636", activeforeground=DORADO)
        b.config(command=lambda b=b, op=cmd_op: mostrar_opcion(op, b)); b.pack(fill="x", padx=10, pady=2); botones_docente.append(b); return b
    btn_menu("Mi Grupo", "Mi Grupo"); btn_menu("Asistencia", "Asistencia"); btn_menu("Justificantes", "Justificantes"); btn_menu("Reportes Alertas", "Reportes"); btn_menu("📖 Diccionario Bilingüe", "Diccionario"); btn_menu("Calificaciones", "Calificaciones")
    tk.Button(menu, text="  Cerrar Sesión", bg=GUINDA, fg=DORADO, font=("Segoe UI", 10, "bold"), bd=0, anchor="w", pady=15, cursor="hand2", command=pantalla_login).pack(side="bottom", fill="x", padx=10, pady=20)
    mostrar_opcion("Mi Grupo", botones_docente[0])

def vista_directivo():
    limpiar(); ventana.configure(bg=BLANCO)
    header_top = tk.Frame(ventana, bg="#611232", height=55); header_top.pack(side="top", fill="x"); header_top.pack_propagate(False)
    tk.Label(header_top, text=f"  Director: {USUARIO_ACTUAL['nombre']}", bg="#611232", fg="white", font=("Segoe UI", 11, "bold")).pack(side="left", pady=12)
    crear_campanita(header_top)
    menu = tk.Frame(ventana, bg="#611232", width=280); menu.pack(side="left", fill="y"); menu.pack_propagate(False)
    tk.Label(menu, text="DIRECTOR\nPANEL", bg="#611232", fg=BLANCO, font=("Segoe UI", 18, "bold"), justify="left").pack(anchor="w", padx=20, pady=20)
    tk.Frame(menu, bg=DORADO, height=3, width=60).pack(anchor="w", padx=20, pady=5)
    tk.Label(menu, text=f"{USUARIO_ACTUAL['nombre']}", bg="#611232", fg=DORADO, font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w", padx=20, pady=(5,0))
    contenido = tk.Frame(ventana, bg=BLANCO); contenido.pack(side="right", fill="both", expand=True)
    def vista_resumen(parent):
        try:
            con = conectar(); cur = con.cursor()
            cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=CURRENT_DATE AND estado='Asistio'"); asist = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=CURRENT_DATE AND estado='Falto'"); falt = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=CURRENT_DATE AND estado='Retardo'"); ret = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=CURRENT_DATE AND tiene_incidencia=1"); inc = cur.fetchone()[0]
            con.close()
        except: asist, falt, ret, inc = 0,0,0,0
        izq = tk.Frame(parent, bg=GUINDA, width=280); izq.pack(side="left", fill="y"); izq.pack_propagate(False)
        tk.Label(izq, text="Control de\nAsistencia", bg=GUINDA, fg=BLANCO, font=("Segoe UI", 18, "bold"), justify="left").pack(anchor="w", padx=20, pady=15)
        def tarjeta(titulo, num, sub):
            c = tk.Frame(izq, bg=BLANCO); c.pack(fill="x", padx=15, pady=4); tk.Frame(c, bg=DORADO, height=2).pack(fill="x"); f = tk.Frame(c, bg=BLANCO, padx=8, pady=5); f.pack(fill="x")
            tk.Label(f, text=titulo, bg=BLANCO, fg=GUINDA, font=("Segoe UI", 7, "bold")).pack(anchor="w"); tk.Label(f, text=str(num), bg=BLANCO, fg=GUINDA, font=("Segoe UI", 18, "bold")).pack(anchor="w"); tk.Label(f, text=sub, bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 7)).pack(anchor="w")
        tarjeta("ASISTIERON", asist, "Presentes hoy"); tarjeta("FALTARON", falt, "Sin justificar"); tarjeta("RETARDOS", ret, "Llegadas tarde"); tarjeta("INCIDENCIAS", inc, "Reportes / incidencias")
        cont = tk.Frame(parent, bg=BLANCO); cont.pack(side="right", fill="both", expand=True)
        tk.Label(cont, text="Resumen General - Datos reales de la base de datos", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=15, pady=20)
        tk.Label(cont, text=f"Hoy: {asist} asistieron, {falt} faltaron, {ret} retardos, {inc} incidencias.", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", padx=15)
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from datetime import timedelta
            con = conectar(); cur = con.cursor(); dias=[]; vals_asis=[]; vals_falt=[]
            for i in range(6,-1,-1):
                d = datetime.now().date() - timedelta(days=i); dias.append(d.strftime("%d/%m"))
                cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=%s AND estado='Asistio'", (d,)); vals_asis.append(cur.fetchone()[0])
                cur.execute("SELECT COUNT(*) FROM asistencia WHERE fecha=%s AND estado='Falto'", (d,)); vals_falt.append(cur.fetchone()[0])
            cur.execute("SELECT a.nombre_completo, COUNT(*) as c FROM asistencia s JOIN alumnos a ON a.id=s.alumno_id WHERE s.estado='Falto' GROUP BY a.nombre_completo ORDER BY c DESC LIMIT 5"); top = cur.fetchall(); con.close()
            fig, axes = plt.subplots(1,3, figsize=(9,2.8)); fig.patch.set_facecolor(BLANCO)
            axes[0].pie([asist,falt,ret] if (asist+falt+ret)>0 else [1], labels=["Asistió","Faltó","Retardo"] if (asist+falt+ret)>0 else ["Sin datos"], colors=["#4E0F27","#C9A86A","#B22222"], autopct="%1.0f%%"); axes[0].set_title("Hoy", color=GUINDA, fontsize=9, fontweight="bold")
            axes[1].bar(dias, vals_asis, color="#4E0F27", label="Asistió"); axes[1].bar(dias, vals_falt, bottom=vals_asis, color="#C9A86A", label="Faltó"); axes[1].set_title("Últimos 7 días", color=GUINDA, fontsize=9, fontweight="bold"); axes[1].legend(fontsize=7)
            if top: axes[2].barh([t[0].split()[0] for t in top], [t[1] for t in top], color="#B22222"); axes[2].set_title("Más faltas", color=GUINDA, fontsize=9, fontweight="bold")
            for ax in axes: ax.set_facecolor(BLANCO); ax.tick_params(colors=GUINDA, labelsize=7)
            canvas = FigureCanvasTkAgg(fig, master=cont); canvas.draw(); canvas.get_tk_widget().pack(fill="x", padx=15, pady=10)
        except Exception as e: tk.Label(cont, text=f"Gráficas no disponibles: {e}", bg=BLANCO, fg="#8A7A70").pack(anchor="w", padx=15)
    def vista_actividad_docente(parent):
        header = tk.Frame(parent, bg=BLANCO); header.pack(fill="x", padx=30, pady=(20,10))
        tk.Label(header, text="✦ Actividad Docente - Últimos Movimientos", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(header, text="Resumen de justificantes, asistencias y reportes registrados en la DB", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", pady=5)
        try:
            con = conectar(); cur = con.cursor(dictionary=True)
            cur.execute("SELECT COUNT(*) as c FROM justificantes"); n_j = cur.fetchone()['c']
            cur.execute("SELECT COUNT(*) as c FROM asistencia WHERE fecha=CURRENT_DATE"); n_a = cur.fetchone()['c']
            cur.execute("SELECT COUNT(*) as c FROM reportes_alerta"); n_r = cur.fetchone()['c']
            cur.execute("SELECT COUNT(*) as c FROM usuarios WHERE rol='docente'"); n_d = cur.fetchone()['c']
            cur.execute("SELECT * FROM justificantes ORDER BY fecha_registro DESC LIMIT 5"); ult_j = cur.fetchall(); con.close()
        except: n_j=n_a=n_r=n_d=0; ult_j=[]
        kpi_frame = tk.Frame(parent, bg=BLANCO, padx=20, pady=10); kpi_frame.pack(fill="x")
        def kpi(titulo, num, color, emoji):
            f = tk.Frame(kpi_frame, bg=BLANCO, padx=15, pady=12, highlightbackground="#EADDC8", highlightthickness=1); f.pack(side="left", padx=8, expand=True, fill="x")
            tk.Frame(f, bg=color, height=4).pack(fill="x", pady=(0,8)); tk.Label(f, text=f"{emoji} {titulo}", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 8, "bold")).pack(anchor="w"); tk.Label(f, text=str(num), bg=BLANCO, fg=GUINDA, font=("Segoe UI", 22, "bold")).pack(anchor="w"); tk.Label(f, text="registros en DB", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 8)).pack(anchor="w")
        kpi("JUSTIFICANTES", n_j, DORADO, "📄"); kpi("ASISTENCIAS HOY", n_a, GUINDA, "✅"); kpi("REPORTES ALERTA", n_r, "#B22222", "⚠"); kpi("DOCENTES ACTIVOS", n_d, "#8B5A2B", "👩‍🏫")
        tk.Label(parent, text="  Últimos justificantes:", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=30, pady=(10,5))
        for j in ult_j:
            card = tk.Frame(parent, bg=BEIGE_CLARO, padx=12, pady=8); card.pack(fill="x", padx=30, pady=3); tk.Label(card, text=f"• {j['fecha_registro']} - Justificante: {j['alumno_nombre']} - {j['motivo']} - {j['grado_grupo']}", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 10)).pack(anchor="w")
    def vista_reportes_director(parent):
        tk.Label(parent, text="Reportes - Todos los grupos", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=25, pady=(20,2))
        tk.Label(parent, text="Aquí el director ve todos los reportes y justificantes que suben los 6 docentes", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", padx=25, pady=(0,10))
        filtros = tk.Frame(parent, bg=BLANCO); filtros.pack(fill="x", padx=25, pady=5)
        tk.Label(filtros, text="Ver:", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 10, "bold")).pack(side="left")
        tipo_var = tk.StringVar(value="Todos"); grupo_var = tk.StringVar(value="Todos los grupos"); buscar_var = tk.StringVar()
        ttk.Combobox(filtros, textvariable=tipo_var, values=["Todos","Reportes de Alerta","Justificantes"], state="readonly", width=20).pack(side="left", padx=8)
        ttk.Combobox(filtros, textvariable=grupo_var, values=["Todos los grupos"]+list(GRADOS.keys()), state="readonly", width=18).pack(side="left", padx=8)
        tk.Entry(filtros, textvariable=buscar_var, font=("Segoe UI", 10), width=28).pack(side="left", padx=8)
        tk.Label(filtros, text="Buscar alumno...", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 8)).pack(side="left")
        lbl_total = tk.Label(parent, text="", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")); lbl_total.pack(anchor="w", padx=25, pady=5)
        kpi = tk.Frame(parent, bg=BLANCO); kpi.pack(fill="x", padx=25, pady=5)
        lbl_alertas = tk.Label(kpi, text="⚠ Alertas: 0", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 10, "bold"), padx=12, pady=6); lbl_alertas.pack(side="left", padx=(0,8))
        lbl_just = tk.Label(kpi, text="📄 Justificantes: 0", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 10, "bold"), padx=12, pady=6); lbl_just.pack(side="left", padx=8)
        lbl_alto = tk.Label(kpi, text="🔴 Riesgo Alto: 0", bg="#FFE0E0", fg="#8B0000", font=("Segoe UI", 10, "bold"), padx=12, pady=6); lbl_alto.pack(side="left", padx=8)
        cols = ("Fecha","Tipo","Alumno","Grupo","Docente","Detalle"); tree = ttk.Treeview(parent, columns=cols, show="headings", height=14)
        for c in cols: tree.heading(c, text=c); tree.column(c, width=140 if c not in ("Alumno","Detalle") else 200)
        tree.column("Fecha", width=110); tree.column("Tipo", width=130); tree.pack(fill="both", expand=True, padx=25, pady=10)
        detalle = tk.Label(parent, text="Selecciona un reporte para ver el detalle completo", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 10), wraplength=900, justify="left", anchor="w", padx=12, pady=8); detalle.pack(fill="x", padx=25, pady=(0,10))
        def on_select(e):
            sel = tree.selection()
            if not sel: return
            vals = tree.item(sel[0])["values"]; detalle.config(text=f"{vals[1]}  •  {vals[2]}  •  {vals[3]}  •  Docente: {vals[4]}  •  Fecha: {vals[0]}\nDetalle: {vals[5]}")
        tree.bind("<<TreeviewSelect>>", on_select)
        def cargar():
            for i in tree.get_children(): tree.delete(i)
            try:
                con = conectar(); cur = con.cursor(dictionary=True)
                cur.execute("SELECT COUNT(*) as c FROM reportes_alerta"); n_a = cur.fetchone()['c']
                cur.execute("SELECT COUNT(*) as c FROM justificantes"); n_j = cur.fetchone()['c']
                cur.execute("SELECT COUNT(*) as c FROM reportes_alerta WHERE nivel_riesgo='Alto'"); n_alto = cur.fetchone()['c']
                lbl_alertas.config(text=f"⚠ Alertas: {n_a}"); lbl_just.config(text=f"📄 Justificantes: {n_j}"); lbl_alto.config(text=f"🔴 Riesgo Alto: {n_alto}"); rows = []
                if tipo_var.get() in ("Todos","Reportes de Alerta"):
                    cur.execute("SELECT * FROM reportes_alerta ORDER BY fecha DESC")
                    for r in cur.fetchall(): rows.append((str(r['fecha']), "⚠ Alerta - "+r['tipo'], r['alumno_nombre'], r['grado_grupo'], r['docente'], f"{r['nivel_riesgo']} | {r['descripcion']} | Acción: {r['accion_propuesta']}"))
                if tipo_var.get() in ("Todos","Justificantes"):
                    cur.execute("SELECT * FROM justificantes ORDER BY fecha_falta DESC")
                    for r in cur.fetchall(): rows.append((str(r['fecha_falta']), "📄 Justificante - "+r['motivo'], r['alumno_nombre'], r['grado_grupo'], r['docente'], f"{r['num_dias']} día(s) | Tutor: {r['tutor']} | {r['descripcion']}"))
                con.close(); filtrados = []
                for row in rows:
                    if grupo_var.get() != "Todos los grupos" and grupo_var.get() not in row[3]: continue
                    if buscar_var.get().strip() and buscar_var.get().strip().lower() not in row[2].lower(): continue
                    filtrados.append(row)
                for row in filtrados: tree.insert("", "end", values=row)
                lbl_total.config(text=f"Mostrando {len(filtrados)} reportes")
            except Exception as e: lbl_total.config(text=f"Error: {e}")
        tk.Button(filtros, text="🔍 BUSCAR", bg=GUINDA, fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=14, pady=4, command=cargar).pack(side="left", padx=8)
        tk.Button(filtros, text="🔄 ACTUALIZAR", bg=DORADO, fg=GUINDA, font=("Segoe UI", 9, "bold"), bd=0, padx=12, pady=4, command=cargar).pack(side="left", padx=4); cargar()
    def vista_control_docentes(parent):
        GUINDA_D = "#7A1020"; GUINDA_CLARO = "#A83246"; CREMA = "#FFF8E7"; DOCENTES = ["Profa. Ana L.", "Profe. Luis G.", "Profa. Maria P.", "Profe. Jose T.", "Profe. Pedro M.", "Profa. Sofia R."]
        canvas = tk.Canvas(parent, bg=CREMA, highlightthickness=0); scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview); scroll_frame = tk.Frame(canvas, bg=CREMA)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))); window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw"); canvas.configure(yscrollcommand=scrollbar.set); canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width)); canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        header = tk.Frame(scroll_frame, bg=GUINDA_D, height=70); header.pack(fill="x"); tk.Label(header, text="CONTROL DOCENTE - ASISTENCIA DE PROFESORES", font=("Arial Black", 14), bg=GUINDA_D, fg=CREMA).pack(pady=20)
        lbl_cont = tk.Label(scroll_frame, text="0 registros", font=("Arial", 10, "bold"), bg=BEIGE, fg=GUINDA_D); lbl_cont.pack(fill="x", padx=20, pady=10)
        form = tk.Frame(scroll_frame, bg=BLANCO, padx=15, pady=15); form.pack(fill="x", padx=20)
        tk.Label(form, text="DOCENTE", font=("Arial", 9, "bold"), bg=BLANCO, fg=GUINDA_D).pack(anchor="w"); var_docente = tk.StringVar(value="Selecciona"); tk.OptionMenu(form, var_docente, *DOCENTES).pack(fill="x", pady=5, ipady=4)
        tk.Label(form, text="MATERIA / GRUPO", font=("Arial", 9, "bold"), bg=BLANCO, fg=GUINDA_D).pack(anchor="w", pady=(10,0)); entry_materia = tk.Entry(form, bg=CREMA, fg=GUINDA_D, relief="flat", font=("Arial", 11)); entry_materia.pack(fill="x", ipady=7, pady=5); entry_materia.insert(0, "Matemáticas 3A")
        lista = tk.Listbox(scroll_frame, bg=BLANCO, fg=GUINDA_D, font=("Consolas", 9), relief="flat", height=12); lbl_status = tk.Label(scroll_frame, text="Listo para registrar", bg=BEIGE, fg=GUINDA_D, font=("Arial", 9))
        def guardar(estado):
            docente = var_docente.get(); materia = entry_materia.get().strip()
            if docente == "Selecciona": messagebox.showwarning("Falta", "Selecciona un docente"); return
            if materia == "": materia = "General"
            try:
                con = conectar(); cur = con.cursor(); cur.execute("INSERT INTO control_docente (docente, materia_grupo, estado, fecha_hora) VALUES (%s,%s,%s,%s)", (docente, materia, estado, datetime.now())); con.commit(); con.close(); lbl_status.config(text=f"✓ {estado} de {docente} guardado en DB"); actualizar()
            except Exception as e: messagebox.showerror("Error DB", str(e))
        def actualizar():
            lista.delete(0, tk.END)
            try:
                con = conectar(); cur = con.cursor(dictionary=True); cur.execute("SELECT * FROM control_docente ORDER BY fecha_hora DESC LIMIT 25"); rows = cur.fetchall(); con.close()
                for r in rows: lista.insert(tk.END, f"  {r['fecha_hora'].strftime('%d/%m/%Y %H:%M')} - {r['docente']} | {r['materia_grupo']} | {r['estado']}")
                lbl_cont.config(text=f"{len(rows)} registros")
            except: lista.insert(0, "  Sin registros aún...")
        tk.Button(form, text="✓ ASISTENCIA", bg=GUINDA_D, fg="white", font=("Arial", 10, "bold"), relief="flat", pady=10, command=lambda: guardar("Asistencia")).pack(fill="x", pady=3)
        tk.Button(form, text="● RETARDO", bg=DORADO, fg="white", font=("Arial", 10, "bold"), relief="flat", pady=8, command=lambda: guardar("Retardo")).pack(fill="x", pady=3)
        tk.Button(form, text="✕ FALTA", bg=GUINDA_CLARO, fg="white", font=("Arial", 10, "bold"), relief="flat", pady=8, command=lambda: guardar("Falta")).pack(fill="x", pady=3)
        lbl_status.pack(fill="x", padx=20, pady=10); lista.pack(fill="both", expand=True, padx=20, pady=5); actualizar(); canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    def vista_agregar_alumno(parent):
        tk.Label(parent, text="➕ Agregar Alumno", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=25, pady=(20,2))
        tk.Label(parent, text="El alumno que agregues aquí le aparecerá al docente en su lista de asistencia.", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", padx=25, pady=(0,15))
        card = tk.Frame(parent, bg=BLANCO, highlightbackground="#EADDC8", highlightthickness=1, padx=20, pady=20); card.pack(fill="x", padx=25, pady=10)
        tk.Label(card, text="Nombre completo del alumno:", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w"); entry_nombre = tk.Entry(card, font=("Segoe UI", 12)); entry_nombre.pack(fill="x", ipady=8, pady=6)
        tk.Label(card, text="Grado:", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8,0)); grado_var = tk.StringVar(value="1° Primaria"); ttk.Combobox(card, textvariable=grado_var, values=list(GRADOS.keys()) if GRADOS else ["1° Primaria","2° Primaria","3° Primaria","4° Primaria","5° Primaria","6° Primaria"], font=("Segoe UI", 11), state="readonly").pack(fill="x", ipady=4, pady=6)
        tk.Label(card, text="Número de lista (déjalo vacío para asignar automático):", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8,0)); entry_num = tk.Entry(card, font=("Segoe UI", 12)); entry_num.pack(fill="x", ipady=8, pady=6)
        def guardar():
            try: num = int(entry_num.get().strip()) if entry_num.get().strip() else 0
            except: num = 0
            if agregar_un_alumno(entry_nombre.get(), grado_var.get(), num): entry_nombre.delete(0,"end"); entry_num.delete(0,"end")
        tk.Button(card, text="AGREGAR ALUMNO", bg="#611232", fg="white", font=("Segoe UI", 12, "bold"), bd=0, pady=12, command=guardar).pack(fill="x", pady=12)
    def vista_carga_masiva(parent):
        tk.Label(parent, text="📂 Carga Masiva de Excel", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=25, pady=(20,2))
        tk.Label(parent, text="Sube un Excel y reemplaza TODOS los alumnos actuales.", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 11)).pack(anchor="w", padx=25, pady=(0,5))
        tk.Label(parent, text="⚠ OJO: Esto BORRA los nombres que están y mete los del Excel.", bg="#FFE0E0", fg="#8B0000", font=("Segoe UI", 11, "bold"), padx=12, pady=8).pack(anchor="w", padx=25, pady=10)
        card = tk.Frame(parent, bg=BLANCO, highlightbackground="#EADDC8", highlightthickness=1, padx=20, pady=20); card.pack(fill="x", padx=25, pady=10)
        tk.Label(card, text="Formato del Excel (fila 1):", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(card, text="Grado  |  Numero_Lista  |  Nombre_Completo\nEjemplo:  1° Primaria  |  1  |  Juan Perez Lopez", bg=BEIGE_CLARO, fg=GUINDA, font=("Consolas", 11), justify="left", padx=10, pady=10).pack(fill="x", pady=8)
        lbl_archivo = tk.Label(card, text="Ningún archivo seleccionado", bg=BLANCO, fg="#8A7A70", font=("Segoe UI", 10)); lbl_archivo.pack(anchor="w", pady=5)
        lbl_status = tk.Label(card, text="", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 10, "bold")); lbl_status.pack(anchor="w", pady=5); ruta_var = {"ruta": ""}
        def elegir():
            ruta = filedialog.askopenfilename(title="Selecciona tu Excel", filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
            if ruta: ruta_var["ruta"] = ruta; lbl_archivo.config(text=f"Archivo: {ruta}")
        def cargar():
            if not ruta_var["ruta"]: messagebox.showwarning("Falta archivo", "Primero elige un Excel"); return
            cargar_masiva_excel(ruta_var["ruta"], lbl_status)
        tk.Button(card, text="📂 ELEGIR EXCEL", bg=BEIGE_CLARO, fg=GUINDA, font=("Segoe UI", 11, "bold"), bd=0, pady=10, command=elegir).pack(fill="x", pady=6)
        tk.Button(card, text="🚀 CARGAR Y REEMPLAZAR TODO", bg="#611232", fg="white", font=("Segoe UI", 12, "bold"), bd=0, pady=12, command=cargar).pack(fill="x", pady=6)
        try:
            con = conectar(); cur = con.cursor(); cur.execute("SELECT COUNT(*) FROM alumnos"); total = cur.fetchone()[0]; con.close(); tk.Label(card, text=f"Alumnos actuales en la base: {total}", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=8)
        except: pass
    botones_director = []
    def marcar_director(btn_activo):
        for b in botones_director:
            if b == btn_activo: b.config(bg="#8B1A4A", fg=DORADO)
            else: b.config(bg="#611232", fg=BLANCO)
    def mostrar_director(op, btn=None):
        if btn: marcar_director(btn)
        for w in contenido.winfo_children(): w.destroy()
        if op == "Resumen": vista_resumen(contenido)
        elif op == "Actividad": vista_actividad_docente(contenido)
        elif op == "Docentes": vista_control_docentes(contenido)
        elif op == "Reportes": vista_reportes_director(contenido)
        elif op == "AgregarAlumno": vista_agregar_alumno(contenido)
        elif op == "CargaMasiva": vista_carga_masiva(contenido)
        elif op == "Diccionario": crear_vista_diccionario(contenido)
        else: tk.Label(contenido, text=f"{op}\n\nMódulo en desarrollo", bg=BLANCO, fg=GUINDA, font=("Segoe UI", 16, "bold")).pack(pady=50)
    def btn_menu(texto, op, color=BLANCO):
        b = tk.Button(menu, text=f"  {texto}", bg="#611232", fg=color, font=("Segoe UI", 11, "bold"), bd=0, anchor="w", pady=15, cursor="hand2", activebackground="#8B1A4A", activeforeground=DORADO)
        b.config(command=lambda b=b, op=op: mostrar_director(op, b)); b.pack(fill="x", padx=10, pady=2); botones_director.append(b); return b
    btn_menu("Resumen General", "Resumen", BLANCO); btn_menu("Actividad Docente", "Actividad", DORADO); btn_menu("Docentes", "Docentes"); btn_menu("Reportes", "Reportes"); btn_menu("➕ Agregar Alumno", "AgregarAlumno"); btn_menu("📂 Carga Masiva Excel", "CargaMasiva"); btn_menu("📖 Diccionario Bilingüe", "Diccionario"); btn_menu("Configuración", "Config")
    tk.Button(menu, text="  Cerrar Sesión", bg="#611232", fg=DORADO, font=("Segoe UI", 10, "bold"), bd=0, anchor="w", pady=15, cursor="hand2", command=pantalla_login).pack(side="bottom", fill="x", padx=10, pady=20)
    mostrar_director("Resumen", botones_director[0])

ventana = tk.Tk()
ventana.state('zoomed')
ventana.configure(bg=BEIGE)
pantalla_login()
ventana.mainloop()





# BASE DE DATOS PARA ESTE CODIGO
# 
# # Tabla alumnos
# # id_alumno es la llave primaria
# # nombre es el nombre del alumno
# # matricula es unica, no se puede repetir
# # grado es el grado del alumno
# # grupo es el grupo del alumno
# # correo es el correo del alumno
# cursor.execute("CREATE TABLE IF NOT EXISTS alumnos (id_alumno INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, matricula TEXT UNIQUE NOT NULL, grado TEXT, grupo TEXT, correo TEXT)")
#
# # Tabla docentes
# # id_docente es la llave primaria
# # nombre es el nombre del docente
# # especialidad es lo que da el docente
# # correo es unico
# cursor.execute("CREATE TABLE IF NOT EXISTS docentes (id_docente INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, especialidad TEXT, correo TEXT UNIQUE)")
#
# # Tabla materias
# # id_materia es la llave primaria
# # nombre_materia es el nombre de la materia
# # id_docente se conecta con la tabla docentes
# cursor.execute("CREATE TABLE IF NOT EXISTS materias (id_materia INTEGER PRIMARY KEY AUTOINCREMENT, nombre_materia TEXT NOT NULL, id_docente INTEGER, FOREIGN KEY (id_docente) REFERENCES docentes(id_docente))")
#
# # Tabla calificaciones
# # id_calificacion es la llave primaria
# # id_alumno se conecta con alumnos
# # id_materia se conecta con materias
# # calificacion es la nota
# # parcial es 1, 2 o 3
# cursor.execute("CREATE TABLE IF NOT EXISTS calificaciones (id_calificacion INTEGER PRIMARY KEY AUTOINCREMENT, id_alumno INTEGER, id_materia INTEGER, calificacion REAL, parcial INTEGER, FOREIGN KEY (id_alumno) REFERENCES alumnos(id_alumno), FOREIGN KEY (id_materia) REFERENCES materias(id_materia))")
#
# # Tabla asistencias
# # id_asistencia es la llave primaria
# # id_alumno se conecta con alumnos
# # fecha es la fecha de la asistencia
# # asistio es 1 si vino y 0 si falto
# cursor.execute("CREATE TABLE IF NOT EXISTS asistencias (id_asistencia INTEGER PRIMARY KEY AUTOINCREMENT, id_alumno INTEGER, fecha TEXT, asistio INTEGER, FOREIGN KEY (id_alumno) REFERENCES alumnos(id_alumno))")
#
# # Para guardar un alumno seria:
# # cursor.execute("INSERT INTO alumnos (nombre, matricula, grado, grupo, correo) VALUES (?, ?, ?, ?, ?)", (nombre, matricula, grado, grupo, correo))
# # Para guardar un docente seria:
# # cursor.execute("INSERT INTO docentes (nombre, especialidad, correo) VALUES (?, ?, ?)", (nombre, especialidad, correo))
# # Para guardar una materia seria:
# # cursor.execute("INSERT INTO materias (nombre_materia, id_docente) VALUES (?, ?)", (nombre_materia, id_docente))
# # Para guardar una calificacion seria:
# # cursor.execute("INSERT INTO calificaciones (id_alumno, id_materia, calificacion, parcial) VALUES (?, ?, ?, ?)", (id_alumno, id_materia, calificacion, parcial))
# # Para consultar calificaciones seria:
# # cursor.execute("SELECT alumnos.nombre, materias.nombre_materia, calificaciones.calificacion FROM calificaciones JOIN alumnos ON calificaciones.id_alumno = alumnos.id_alumno JOIN materias ON calificaciones.id_materia = materias.id_materia")
# # Para guardar asistencia seria:
# # cursor.execute("INSERT INTO asistencias (id_alumno, fecha, asistio) VALUES (?, ?, ?)", (id_alumno, fecha, asistio))


