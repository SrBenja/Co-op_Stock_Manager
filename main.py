import urllib.request
import urllib.error
import webbrowser
import tkinter.simpledialog as simpledialog
import subprocess
import tempfile
import time
import json as _json
import sys, os
import sqlite3
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd  #type: ignore
from datetime import datetime
from reportlab.lib.pagesizes import A4  #type: ignore
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph  #type: ignore
from reportlab.lib import colors  #type: ignore
from reportlab.lib.styles import getSampleStyleSheet  #type: ignore

# --- CONSTANTES GLOBALES ---
DB_NAME      = "stock_co-op.db"
TABLE_NAME   = "productos"
BACKUP_DIR   = "backup"
CONFIG_DIR   = "config"
MAX_UNDO     = 30

VERSION = "v0.1.0"   # incrementar esto cada vez que publique una nueva versión

def _norm_tag(s):
    """
    Normaliza una etiqueta de versión para comparar:
    - asegura que sea str,
    - quita espacios en los extremos,
    - elimina prefijo 'v' o 'V' si existe.
    Ej: ' v0.1.1 ' -> '0.1.1'
    """
    if s is None:
        return ""
    return str(s).strip().lstrip("vV").strip()


def _ask_update_custom(parent, title, message, width=760):
    """
    Diálogo modal centrado en pantalla con botones: Sí | No | Cancelar.
    - width: anchura en píxeles que queremos para el dialog (wraplength se ajusta).
    - Cierra con X o Escape -> "cancel".
    - Devuelve: "yes", "no" o "cancel".
    - Hereda modo claro/oscuro según current_theme global.
    """
    # Crear Toplevel modal
    win = tk.Toplevel(parent)
    win.title(title)
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    # Asegurar visible
    try:
        win.deiconify()
    except Exception:
        pass

    # Recuperar tema actual (si existe)
    theme = globals().get("current_theme", None)

    # Colores coherentes con apply_light/apply_dark
    if theme == "dark":
        bg = "#2e2e2e"
        lbl_fg = "white"
        btn_bg = "#3e3e3e"
        btn_fg = "white"
        btn_hover = "#505050"
        pad_xy = (14, 12)
    else:
        bg = "SystemButtonFace"
        lbl_fg = "black"
        btn_bg = "SystemButtonFace"
        btn_fg = "black"
        btn_hover = "SystemHighlight"
        pad_xy = (12, 10)

    # Estilos ttk locales (no tocar globales)
    style = ttk.Style(win)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    frame_style = "UpdateDlg.TFrame"
    label_style = "UpdateDlg.TLabel"
    button_style = "UpdateDlg.TButton"

    style.configure(frame_style, background=bg)
    style.configure(label_style, background=bg, foreground=lbl_fg, wraplength=width)
    style.configure(button_style,
                    background=btn_bg,
                    foreground=btn_fg,
                    padding=6,
                    relief="raised")
    style.map(button_style,
              background=[("active", btn_hover), ("!disabled", btn_bg)],
              foreground=[("active", btn_fg), ("!disabled", btn_fg)],
              relief=[("pressed", "sunken"), ("!pressed", "raised")])

    # Aplicar background al Toplevel por si el gestor lo respeta
    try:
        win.configure(bg=bg)
    except Exception:
        pass

    # Cerrar con X -> cancelar
    def _on_close():
        win._result = "cancel"
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", _on_close)

    # Contenedor principal
    container = ttk.Frame(win, style=frame_style, padding=pad_xy)
    container.pack(fill="both", expand=True)

    # Label de mensaje
    lbl = ttk.Label(container, text=message, style=label_style, justify="left")
    lbl.pack(fill="both", expand=True)

    # Frame para botones
    btn_frame = ttk.Frame(container, style=frame_style)
    btn_frame.pack(pady=(10, 6))

    def choose(res):
        win._result = res
        win.destroy()

    # Crear botones: Sí | No | Cancelar
    b_yes = ttk.Button(btn_frame, text="Sí", width=12, command=lambda: choose("yes"), style=button_style)
    b_no = ttk.Button(btn_frame, text="No", width=12, command=lambda: choose("no"), style=button_style)
    b_cancel = ttk.Button(btn_frame, text="Cancelar", width=12, command=lambda: choose("cancel"), style=button_style)

    b_yes.pack(side="left", padx=6)
    b_no.pack(side="left", padx=6)
    b_cancel.pack(side="left", padx=6)

    # Escape => cancelar
    win.bind("<Escape>", lambda e: choose("cancel"))

    # Foco por defecto en NO (para que X sea acción natural de cancelar)
    try:
        b_no.focus_set()
    except Exception:
        pass

    # ---------- Forzar cálculo de tamaño y luego centrar correctamente ----------
    try:
        # Forzar layout
        win.update_idletasks()
        # Tamaño requerido por widgets
        req_w = win.winfo_reqwidth()
        req_h = win.winfo_reqheight()
        # Asegurar ancho mínimo
        final_w = max(req_w, width)
        final_h = req_h
        # Calcular posición centrada en pantalla
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, (sw - final_w) // 2)
        y = max(0, (sh - final_h) // 2)
        win.geometry(f"{final_w}x{final_h}+{x}+{y}")
        win.update_idletasks()
    except Exception:
        # Fallback básico
        try:
            win.update()
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            final_w = width
            final_h = win.winfo_height() or 200
            x = max(0, (sw - final_w) // 2)
            y = max(0, (sh - final_h) // 2)
            win.geometry(f"{final_w}x{final_h}+{x}+{y}")
        except Exception:
            pass

    # Topmost momentáneo para asegurar visibilidad
    try:
        win.lift()
        win.attributes("-topmost", True)
        win.after(200, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

    win.wait_window()
    return getattr(win, "_result", "cancel")


def check_updates():
    """
    Consulta la última release en GitHub (usa github_repo en config o el repo por defecto).
    Usa github_token en config para aumentar cuota si está presente.
    """
    cfg = load_config()
    default_repo = "SrBenja/Co-op_Stock_Manager"
    repo = cfg.get("github_repo", default_repo)
    if "github_repo" not in cfg:
        cfg["github_repo"] = repo
        save_config(cfg)

    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    try:
        headers = {"User-Agent": f"Co-op_Stock_Manager/{VERSION} (+https://github.com/{repo})"}
        token = cfg.get("github_token")
        if token:
            headers["Authorization"] = f"token {token}"
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            j = _json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            messagebox.showinfo("Actualizaciones", "No se encontraron releases en el repositorio. Seguramente los haya pronto. Disculpe las molestias.")
            return
        elif e.code == 403:
            messagebox.showerror("Actualizaciones", "Límite de peticiones alcanzado (rate limit). Considere añadir un token en config/config.json como 'github_token' para aumentar la cuota o espere una hora.\nPara más información, visite Ayuda y actualizaciones --> Cómo funciona --> Actualizaciones")
            return
        else:
            messagebox.showerror("Actualizaciones", f"Error al consultar GitHub: {e}")
            return
    except Exception as e:
        messagebox.showerror("Actualizaciones", f"No se pudo comprobar la actualización:\n{e}")
        return

    latest_tag = j.get("tag_name") or j.get("name") or ""
    body = j.get("body", "") or ""
    html_url = j.get("html_url", "")

    if not latest_tag:
        messagebox.showinfo("Actualizaciones", "No se pudo leer la etiqueta de la última release.")
        return

    if _norm_tag(latest_tag) == _norm_tag(VERSION):
        messagebox.showinfo("Actualizaciones", f"Estás en la última versión: {VERSION}")
        return

    summary = (body[:800] + "...") if len(body) > 800 else body
    msg = (f"Versión instalada: {VERSION}\n"
           f"Versión disponible: {latest_tag}\n\n"
           f"{summary}\n\n"
           "¿Descargar e instalar automáticamente (Sí), abrir la release en el navegador (No) o cancelar?")

    # parent: intenta usar root si existe para que la ventana sea modal sobre la app
    parent = globals().get("root", None)
    choice = _ask_update_custom(parent, "Actualización disponible", msg)

    if choice == "yes":
        download_and_install_release_exe(repo, preferred_asset_name="Co-op_Stock_Manager.exe")
    elif choice == "no":
        if html_url:
            webbrowser.open(html_url)
        else:
            messagebox.showinfo("Actualizaciones", "No hay URL de la release para abrir.")
    else:
        # cancelar o cerrar por X -> no hacemos nada
        return


# RUTAS ABSOLUTAS
if getattr(sys, "frozen", False):
    BASE_DIR   = os.path.dirname(sys.executable)
else:
    BASE_DIR   = os.path.dirname(os.path.abspath(__file__))

DB_PATH      = os.path.join(BASE_DIR, DB_NAME)
BACKUP_PATH  = os.path.join(BASE_DIR, BACKUP_DIR)
CONFIG_PATH  = os.path.join(BASE_DIR, CONFIG_DIR)
CONFIG_FILE  = os.path.join(CONFIG_PATH, "config.json")

def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return _json.load(f)
    except (FileNotFoundError, _json.JSONDecodeError):
        return {}

def save_config(cfg: dict):
    os.makedirs(CONFIG_PATH, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        _json.dump(cfg, f, indent=2)

# -- DEFAULT GITHUB REPO (para que la app lo tenga sin editar config manualmente) --
DEFAULT_REPO = "SrBenja/Co-op_Stock_Manager"

def ensure_default_config():
    """
    Si no existe config/config.json, crea uno mínimo con github_repo = DEFAULT_REPO.
    No añade token ni datos privados — así cualquiera que descargue el código tendrá
    el repo preconfigurado sin necesidad de editar archivos.
    """
    cfg = load_config()  # load_config ya devuelve {} si no existe o es inválido
    changed = False
    if "github_repo" not in cfg or not cfg.get("github_repo"):
        cfg["github_repo"] = DEFAULT_REPO
        changed = True
    # No agregamos github_token aquí (el usuario debe añadirlo si lo desea)
    if changed:
        save_config(cfg)

def ensure_dirs():
    os.makedirs(BACKUP_PATH, exist_ok=True)
    os.makedirs(CONFIG_PATH, exist_ok=True)

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        # Si es la primera vez, creamos con la nueva columna "orden"
        conn.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                categoria       TEXT,
                codigo          TEXT,
                descripcion     TEXT,
                cantidad        INTEGER,
                precio_lista    REAL,
                iva             INTEGER,
                bnf             REAL    DEFAULT 0,
                precio_final    REAL,
                importe         REAL,
                fecha_retiro    TEXT,
                orden           INTEGER DEFAULT 0
            )
        """)
        # Si la columna "orden" no existe (en bases antiguas), la agregamos
        cursor = conn.execute(f"PRAGMA table_info({TABLE_NAME})")
        cols = [row[1] for row in cursor.fetchall()]
        if "orden" not in cols:
            conn.execute(f"ALTER TABLE {TABLE_NAME} ADD COLUMN orden INTEGER DEFAULT 0")

# todas las columnas en la BD, incl. 'id'
COLUMNS = [
    "id", "categoria", "orden", "cantidad", "codigo", "descripcion",
    "iva", "precio_lista", "bnf",
    "precio_final", "importe", "fecha_retiro"
]
VISIBLE_COLUMNS = [c for c in COLUMNS if c not in ("id", "categoria", "orden")]
IMPORT_COLUMNS  = [c for c in COLUMNS if c not in ("id", "categoria", "orden")]

COLUMN_LABELS = {
    "cantidad":           "Cantidad",
    "codigo":             "Cod. Art.",
    "descripcion":        "Concepto",
    "iva":                "IVA",
    "precio_lista":       "P. lista",
    "bnf":                "BNF",
    "precio_final":       "Precio",
    "importe":            "Importe",
    "fecha_retiro":       "Fecha de retiro"
}
# 4. FUNCIONES DE NEGOCIO (CRUD, snapshots, backup, import/export, imprimir)
undo_stack, redo_stack = [], []
current_id = None

def _parse_number_from_db(value):
    """
    Convierte valores provenientes de la BD a float:
    - acepta int/float
    - acepta cadenas como '123.45 $', '21 %', '1.234,56 $'
    - devuelve 0.0 si no puede parsear
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "":
        return 0.0
    # Quitamos símbolos y espacios
    s = s.replace("$", "").replace("%", "").strip()
    # Normalizamos miles/decimal: '1.234,56' -> '1234.56'
    # Si hay más de un punto y también una coma, asumimos formato europeo
    if s.count(".") > 1 and "," in s:
        s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _refresh_tree(rows):
    tree.delete(*tree.get_children())

    for row in rows:
        # construimos los valores en el orden de VISIBLE_COLUMNS
        vals = [row[COLUMNS.index(col)] for col in VISIBLE_COLUMNS]
        cantidad = row[COLUMNS.index("cantidad")]
        tags     = ("bajo_stock",) if cantidad < 5 else ()
        tree.insert(
            "", tk.END,
            iid=str(row[0]),
            values=vals,
            tags=tags
        )
    update_status()

class DropdownMenu:
    RAINBOW = ["#FF0000", "#FFA500", "#FFFF00",
               "#008000", "#0000FF", "#4B0082", "#EE82EE"]
    # Lista de todas las instancias creadas
    _instances = []

    def __init__(self, master, items, dark_bg="#2e2e2e", light_bg="white"):
        self.master     = master
        self.items      = items
        self.top        = None
        self.dark_bg    = dark_bg
        self.light_bg   = light_bg
        self._root      = master.winfo_toplevel()
        self._motion_id = None

        # Registramos esta instancia
        DropdownMenu._instances.append(self)

    def show(self, x, y):
        # Cerrar cualquier otro menú
        for dd in DropdownMenu._instances:
            if dd is not self:
                dd.close()

        # Cerrar este si estaba abierto
        self.close()

        bg = self.dark_bg if current_theme == "dark" else self.light_bg
        fg = "white"      if current_theme == "dark" else "black"

        # Crear Toplevel
        self.top = tk.Toplevel(self.master)
        self.top.overrideredirect(True)
        self.top.transient(self.master)
        self.top.lift()
        self.top.geometry(f"+{x}+{y}")
        self.top.config(bg=bg)

        frm = tk.Frame(self.top, bd=1, relief="solid", bg=bg)
        frm.pack()

        idx = 0
        for text, cmd, sep in self.items:
            if sep:
                ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=2)
                continue

            btn = tk.Button(
                frm,
                text=text,
                anchor="w",
                relief="flat",
                bg=bg,
                fg=fg,
                activeforeground=fg,
                command=lambda c=cmd: (self.close(), c())
            )
            btn.pack(fill="x", padx=5, pady=2)

            color = self.RAINBOW[idx % len(self.RAINBOW)]
            idx += 1
            btn.bind("<Enter>", lambda e, b=btn, c=color: b.config(bg=c))
            btn.bind("<Leave>", lambda e, b=btn, bg=bg: b.config(bg=bg))

        # Cierre automático: salida del menú o fuera del botón
        self.top.bind("<Leave>", lambda e: self._check_leave())
        self._motion_id = self._root.bind("<Motion>", lambda e: self._check_leave(), add="+")

    def _check_leave(self):
        if not self.top:
            return

        # Posición del cursor
        x, y = self._root.winfo_pointerx(), self._root.winfo_pointery()

        # Área del menú
        ax, ay = self.top.winfo_rootx(), self.top.winfo_rooty()
        w,  h  = self.top.winfo_width(), self.top.winfo_height()

        # Área del botón maestro
        bx, by = self.master.winfo_rootx(), self.master.winfo_rooty()
        bw, bh = self.master.winfo_width(), self.master.winfo_height()

        inside_menu   = (ax <= x <= ax+w and ay <= y <= ay+h)
        inside_button = (bx <= x <= bx+bw and by <= y <= by+bh)

        if not (inside_menu or inside_button):
            self.close()

    def close(self):
        if self.top:
            self.top.destroy()
            self.top = None
        if self._motion_id:
            self._root.unbind("<Motion>", self._motion_id)
            self._motion_id = None

def cargar_datos():
    cat = CATEGORIES[current_cat_idx]
    sql = f"SELECT {', '.join(COLUMNS)} FROM {TABLE_NAME} WHERE categoria = ? ORDER BY orden, id"
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(sql, (cat,)).fetchall()
    _refresh_tree(rows)

def buscar(event=None):
    term = entry_search.get().strip()
    cat  = CATEGORIES[current_cat_idx]
    if term:
        sql = f"""
            SELECT {', '.join(COLUMNS)}
              FROM {TABLE_NAME}
             WHERE categoria = ?
               AND (codigo LIKE ? OR descripcion LIKE ?)
        """
        pat = f"%{term}%"
        params = (cat, pat, pat)
    else:
        sql    = f"SELECT {', '.join(COLUMNS)} FROM {TABLE_NAME} WHERE categoria = ?"
        params = (cat,)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(sql, params).fetchall()
    _refresh_tree(rows)

# ORDENAR COLUMNAS
_sort_state = {col: False for col in VISIBLE_COLUMNS}

def sort_column(col):
    # Recolectamos (valor, iid) y convertimos a numérico si podemos
    data = [(tree.set(k, col), k) for k in tree.get_children('')]
    try:
        data = [(float(v), k) for v, k in data]
    except:
        pass

    # Ordenamos según el estado anterior (asc/desc)
    data.sort(reverse=_sort_state[col])

    # Recorremos e insertamos cada fila en la nueva posición
    for new_pos, (_, iid) in enumerate(data):
        tree.move(iid, '', new_pos)

    # Persistimos el nuevo orden en la BD
    with sqlite3.connect(DB_PATH) as conn:
        for new_pos, (_, iid) in enumerate(data):
            conn.execute(
                f"UPDATE {TABLE_NAME} SET orden = ? WHERE id = ?",
                (new_pos, iid)
            )

    # Invertimos criterio para la próxima vez que hagas clic en el encabezado
    _sort_state[col] = not _sort_state[col]

def update_status():
    """
    Calcula número total de productos y suma de importes.
    Parsea importes que estén guardados como texto con '$' o '%' y suma correctamente.
    """
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()
        # Total productos
        cur.execute(f"SELECT COUNT(*) FROM {TABLE_NAME}")
        cnt = cur.fetchone()[0]
        # Recuperamos todos los importes crudos y los parseamos
        cur.execute(f"SELECT importe FROM {TABLE_NAME}")
        rows = cur.fetchall()

    total = 0.0
    for (imp_raw,) in rows:
        total += _parse_number_from_db(imp_raw)

    # Mostrar símbolo $ en el status
    status_var.set(f"Total productos: {cnt}    |    Valor total del stock: {total:.2f} $")

def snapshot():
    """Guarda el estado actual de la tabla en undo_stack, limitando su tamaño."""
    with sqlite3.connect(DB_PATH) as conn:
        df = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME}", conn)
    undo_stack.append(df)
    # Si superamos el límite, descartamos el más antiguo
    if len(undo_stack) > MAX_UNDO:
        undo_stack.pop(0)
    # Cada vez que das un snapshot, el redo ya no tiene sentido
    redo_stack.clear()

def _get_table_columns(table: str = TABLE_NAME):
    """Devuelve la lista de columnas (en orden) de la tabla SQLite `table`."""
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(f"PRAGMA table_info({table})")
        cols = [row[1] for row in cur.fetchall()]
    return cols


def _reindex_df_to_table_schema(df: pd.DataFrame, table: str = TABLE_NAME) -> pd.DataFrame:
    """
    Reindexa `df` para que coincida exactamente con las columnas de `table`.
    - Si faltan columnas en df, se añaden con valores NaN (luego convertidos a None).
    - Si df tiene columnas extras, se descartan.
    - Devuelve un DataFrame con las columnas en el mismo orden que la tabla.
    """
    cols = _get_table_columns(table)
    if df is None:
        # devolvemos un DataFrame vacío con las columnas correctas
        return pd.DataFrame(columns=cols)
    # Hacemos copia para no mutar el original
    df2 = df.copy()
    # Reindexar para que tenga exactamente las columnas de la tabla (añade NaN si faltan)
    df2 = df2.reindex(columns=cols)
    return df2


def _restore_df_preserve_schema(df: pd.DataFrame, table: str = TABLE_NAME):
    """
    Restaura los contenidos de `df` en la tabla `table` sin reemplazar el esquema.
    - Limpia la tabla actual y re-inserta las filas del DataFrame reindexado
      para coincidir con las columnas de la tabla.
    """
    # Si df es None o vacío: borramos filas y salimos
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(f"DELETE FROM {table}")
            conn.commit()
        return

    # Reindexamos el DataFrame para que coincida exactamente con el esquema de la tabla
    df_reindexed = _reindex_df_to_table_schema(df, table)

    # Convertir NaN -> None para que sqlite reciba NULL
    df_clean = df_reindexed.where(pd.notnull(df_reindexed), None)

    cols = list(df_clean.columns)
    if not cols:
        # Si por alguna razón no hay columnas, limpiamos y salimos
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(f"DELETE FROM {table}")
            conn.commit()
        return

    placeholders = ",".join(["?"] * len(cols))
    insert_sql = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})"

    rows = []
    for _, r in df_clean.iterrows():
        vals = [r[c] for c in cols]
        rows.append(tuple(vals))

    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()
        # Borrar el contenido actual preservando esquema (PK, AUTOINCREMENT, índices)
        cur.execute(f"DELETE FROM {table}")
        # Insertar filas (si hay)
        if rows:
            cur.executemany(insert_sql, rows)
        conn.commit()


def deshacer(event=None):
    if not undo_stack:
        messagebox.showinfo(
            title="Deshacer",
            message="Nada para deshacer."
        )
        return

    # Guardamos el estado actual antes de restaurar (para poder rehacer luego)
    with sqlite3.connect(DB_PATH) as conn:
        estado_actual = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME}", conn)
    redo_stack.append(estado_actual)

    # Tomamos el estado previo y lo eliminamos de undo_stack
    df_prev = undo_stack.pop()

    # Restauramos ese estado en la base de datos SIN reemplazar el esquema
    try:
        _restore_df_preserve_schema(df_prev, TABLE_NAME)
    except Exception as e:
        messagebox.showerror("Deshacer", f"No se pudo restaurar el estado:\n{e}")
        return

    # Refrescamos la vista
    limpiar_form()
    cargar_datos()


def rehacer(event=None):
    if not redo_stack:
        messagebox.showinfo(
            title="Rehacer",
            message="Nada para rehacer."
        )
        return

    # Guardamos el estado actual para poder deshacerlo luego
    with sqlite3.connect(DB_PATH) as conn:
        estado_actual = pd.read_sql_query(f"SELECT * FROM {TABLE_NAME}", conn)
    undo_stack.append(estado_actual)

    # Recuperamos el siguiente estado de redo_stack
    df_next = redo_stack.pop()

    # Restauramos ese estado en la base de datos sin reemplazar el esquema
    try:
        _restore_df_preserve_schema(df_next, TABLE_NAME)
    except Exception as e:
        messagebox.showerror("Rehacer", f"No se pudo restaurar el estado:\n{e}")
        return

    # Limpiamos el formulario y refrescamos la vista
    limpiar_form()
    cargar_datos()


def backup_db():
    """Copia la base con timestamp en BACKUP_PATH."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{TABLE_NAME}_backup_{ts}.db"
    dst   = os.path.join(BACKUP_PATH, fname)
    shutil.copyfile(DB_PATH, dst)
    return dst

def restore_backup():
    """Restaura desde un .db en BACKUP_PATH."""
    path = filedialog.askopenfilename(
        initialdir=BACKUP_PATH,
        title="Seleccionar backup para restaurar",
        filetypes=[("SQLite DB","*.db")]
    )
    if not path:
        return

    if not messagebox.askyesno(
        title="Restaurar backup",
        message=f"¿Restaurar desde:\n{os.path.basename(path)}?"
    ):
        return

    # Reemplazamos el fichero real de datos
    shutil.copyfile(path, DB_PATH)
    limpiar_form()
    cargar_datos()
    messagebox.showinfo(
        title="Restauración",
        message="Restauración completada correctamente."
    )

def manual_backup():
    dst = backup_db()
    messagebox.showinfo(
        title="Backup manual",
        message=f"Backup creado:\n{os.path.basename(dst)}"
    )

def importar_csv():
    path = filedialog.askopenfilename(
        title="Seleccionar CSV para importar",
        filetypes=[("CSV", "*.csv")]
    )
    if not path:
        return

    # Leemos todo como strings para preservar símbolos y celdas vacías
    df = pd.read_csv(path, dtype=str)

    # Normalizamos nombres de columnas (si vienen)
    df = df.rename(columns={
        "Cantidad":           "cantidad",
        "Cod. Art.":          "codigo",
        "Concepto":           "descripcion",
        "Descripción":        "descripcion",
        "% IVA":              "iva",
        "IVA":                "iva",
        "P. lista":           "precio_lista",
        "BNF":                "bnf",
        "Precio":             "precio_final",
        "Importe":            "importe",
        "Fecha de retiro":    "fecha_retiro"
    })

    # Si no viene 'categoria', asumimos la categoría actual para todas las filas
    if "categoria" not in df.columns:
        df["categoria"] = CATEGORIES[current_cat_idx]

    # Permitimos que venga 'orden' — no será motivo de error.
    expected = set(IMPORT_COLUMNS)
    present = set(df.columns) - {"categoria", "orden"}
    missing = [c for c in expected if c not in present]
    if missing:
        return messagebox.showerror(
            "Importación inválida",
            f"Faltan columnas obligatorias en el CSV: {missing}"
        )

    # Normalizamos textos: evitar "nan" -> dejar vacío
    for text_col in ("codigo", "descripcion", "fecha_retiro"):
        if text_col in df.columns:
            df[text_col] = df[text_col].fillna("").astype(str)

    # Parseo de columnas numéricas usando _parse_number_from_db (acepta $, % y formatos)
    numeric_cols = ["cantidad", "precio_lista", "iva", "bnf", "precio_final", "importe"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(lambda v: _parse_number_from_db(v) if v is not None and str(v).strip() != "" else 0.0)

    # Aseguramos tipos enteros para cantidad e iva
    try:
        if "cantidad" in df.columns:
            df["cantidad"] = df["cantidad"].astype(int)
        if "iva" in df.columns:
            df["iva"] = df["iva"].astype(int)
    except Exception:
        return messagebox.showerror("Importación inválida", "Las columnas 'cantidad' e 'iva' deben ser valores enteros válidos.")

    # Para precio_lista, bnf, precio_final, importe -> float
    for col in ("precio_lista", "bnf", "precio_final", "importe"):
        if col in df.columns:
            df[col] = df[col].astype(float)

    # Orden
    if "orden" in df.columns:
        def _parse_orden(x):
            try:
                return int(float(x))
            except Exception:
                return -1
        df["orden"] = df["orden"].fillna(-1).apply(_parse_orden)
    else:
        df["orden"] = -1

    # Guardamos snapshot para deshacer
    snapshot()

    # Insertar en DB: si alguna fila tiene orden == -1, calculamos el next_orden por su categoría
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()
        cats_with_missing = df.loc[df["orden"] == -1, "categoria"].unique().tolist()
        next_orden_map = {}
        for cat in cats_with_missing:
            q = conn.execute(f"SELECT COALESCE(MAX(orden), -1) FROM {TABLE_NAME} WHERE categoria = ?", (cat,)).fetchone()[0]
            next_orden_map[cat] = q + 1

        insert_sql = f"""
            INSERT INTO {TABLE_NAME} (
                categoria, codigo, descripcion, cantidad,
                iva, precio_lista, bnf, precio_final,
                importe, fecha_retiro, orden
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """
        rows_to_insert = []
        for _, r in df.iterrows():
            categoria = r["categoria"]
            codigo    = r.get("codigo", "") or ""
            descripcion = r.get("descripcion", "") or ""
            cantidad  = int(r.get("cantidad", 0) or 0)
            iva_val   = int(r.get("iva", 0) or 0)
            precio_lista_val = float(r.get("precio_lista", 0.0) or 0.0)
            bnf_val   = float(r.get("bnf", 0.0) or 0.0)
            precio_final_val = float(r.get("precio_final", 0.0) or 0.0)
            importe_val = float(r.get("importe", 0.0) or 0.0)
            retiro    = r.get("fecha_retiro", "") or ""

            orden_val = int(r.get("orden", -1) or -1)
            if orden_val == -1:
                orden_val = next_orden_map.setdefault(categoria, next_orden_map.get(categoria, 0))
                next_orden_map[categoria] = orden_val + 1

            # --- Aquí formateamos de vuelta con símbolos ANTES de insertar ---
            precio_lista_str = f"{precio_lista_val:.1f} $"
            precio_final_str = f"{precio_final_val:.3f} $"
            importe_str      = f"{importe_val:.2f} $"
            iva_str          = f"{iva_val} %"

            rows_to_insert.append((
                categoria, codigo, descripcion, cantidad,
                iva_str, precio_lista_str, bnf_val,
                precio_final_str, importe_str, retiro, orden_val
            ))

        cur.executemany(insert_sql, rows_to_insert)
        conn.commit()

    # Refrescamos vista
    cargar_datos()
    messagebox.showinfo("Importación", f"{len(rows_to_insert)} productos importados correctamente.")


def exportar_csv():
    """Exporta a CSV todos los productos de todas las categorías,
    incluyendo las columnas 'categoria' y 'orden', respetando el orden actual."""
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV","*.csv")],
        title="Guardar todos los productos como CSV"
    )
    if not path:
        return

    import csv
    try:
        with sqlite3.connect(DB_PATH) as conn, open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            # Cabecera: categoria, orden y luego las columnas visibles (en tu orden actual)
            headers = ["categoria", "orden"] + VISIBLE_COLUMNS
            writer.writerow(headers)

            # Leemos directamente de la DB, ordenando por categoría y luego por orden
            cursor = conn.execute(f"""
                SELECT categoria, orden, {', '.join(VISIBLE_COLUMNS)}
                  FROM {TABLE_NAME}
                 ORDER BY categoria, orden
            """)
            for row in cursor:
                writer.writerow(row)

        messagebox.showinfo("Exportar CSV", f"Todos los productos exportados correctamente a:\n{path}")
    except Exception as e:
        messagebox.showerror("Error al exportar CSV", str(e))


def imprimir_stock():
    # Pedir ruta de guardado
    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        title="Guardar informe como"
    )
    if not path:
        return

    # Preparamos documento y estilos
    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=40, rightMargin=40,
        topMargin=40, bottomMargin=40
    )
    styles = getSampleStyleSheet()
    styleN = styles["BodyText"]
    styleN.wordWrap = 'CJK'  # activa el wrap

    # Encabezados y recogida de datos
    data = []
    cols_to_print = [c for c in COLUMNS if c != "id"]
    header_pars = [Paragraph(COLUMN_LABELS[c], styleN) for c in cols_to_print]
    data.append(header_pars)

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.execute(f"SELECT {', '.join(COLUMNS)} FROM {TABLE_NAME}")
        for row in cursor:
            vals = []
            for idx, c in enumerate(COLUMNS):
                if c == "id":
                    continue
                vals.append(Paragraph(str(row[idx]), styleN))
            data.append(vals)

    # Anchos iguales
    page_w, page_h = A4
    usable_w = page_w - doc.leftMargin - doc.rightMargin
    n_cols   = len(cols_to_print)
    colWidths = [usable_w / n_cols] * n_cols

    # Tabla con encabezado repetido
    table = Table(data, colWidths=colWidths, repeatRows=1)
    table.setStyle(TableStyle([
        ('GRID',       (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0),   colors.lightblue),
        ('VALIGN',     (0,0), (-1,-1),  'TOP'),
        ('WORDWRAP',   (0,0), (-1,-1),  True),
    ]))

    # Generamos el PDF y lo abrimos
    doc.build([table])
    messagebox.showinfo("Imprimir", f"PDF generado en:\n{path}")
    if os.name == "nt":
        os.startfile(path)
    else:
        os.system(f'xdg-open "{path}"')

def calcular_importe(event=None):
    # Leemos y limpiamos el contenido de P. lista
    raw = entry_precio_lista.get().strip().rstrip(" $").replace(",", ".")
    if not raw:
        entry_precio_final.delete(0, tk.END)
        entry_importe.delete(0, tk.END)
        return

    try:
        pl = float(raw)
        iva = int(entry_iva.get().strip().rstrip("%") or 0)
        pf = pl * (1 + iva/100)

        # Precio con dos decimales y $
        entry_precio_final.delete(0, tk.END)
        entry_precio_final.insert(0, f"{pf:.2f} $")

        # Importe = cantidad * pf
        cantidad = int(entry_cantidad.get().strip() or 0)
        imp = round(cantidad * pf, 2)
        entry_importe.delete(0, tk.END)
        entry_importe.insert(0, f"{imp:.2f} $")
    except ValueError:
        # si falla el parseo, no hacemos nada
        pass

def formatear_precio_lista(event=None):
    try:
        # Quitamos espacios
        s = entry_precio_lista.get().replace(" ", "")
        # Si detectamos miles y decimal → 1.234,56 => 1234.56
        if "." in s and "," in s:
            s2 = s.replace(".", "").replace(",", ".")
        else:
            # sólo decimal con coma o punto
            s2 = s.replace(",", ".")
        # Convertimos a float y formateamos a 3 decimales
        pl = float(s2 or 0)
        entry_precio_lista.delete(0, tk.END)
        entry_precio_lista.insert(0, f"{pl:.3f}")
    except:
        pass

def limpiar_form():
    global current_id
    current_id = None

    # Limpio todos los Entry…
    for widget in (
        entry_cantidad, entry_codigo, entry_descripcion,
        entry_iva, entry_precio_lista, entry_bnf,
        entry_precio_final, entry_importe, entry_retiro
    ):
        widget.delete(0, tk.END)

    # Deselecciono cualquier fila
    for sel in tree.selection():
        tree.selection_remove(sel)

def guardar_producto():
    global current_id

    # Validar código
    codigo = entry_codigo.get().strip()
    if not codigo:
        return messagebox.showwarning("Validación", "El campo Cod. Art. es obligatorio.")

    # Helper para limpiar texto numérico y símbolos
    def clean_numeric(text):
        return text.strip().rstrip(" $%").replace(",", ".")

    # Validar numéricos
    try:
        cantidad     = int(clean_numeric(entry_cantidad.get()) or 0)
        precio_lista = float(clean_numeric(entry_precio_lista.get()) or 0)
        iva          = int(clean_numeric(entry_iva.get()) or 0)
        bnf          = float(clean_numeric(entry_bnf.get()) or 0)
    except ValueError:
        return messagebox.showwarning(
            "Validación",
            "Cantidad, Precio de Lista, IVA y BNF deben ser números válidos."
        )

    # Validaciones de rango
    if cantidad < 0:
        return messagebox.showwarning("Validación", "La cantidad no puede ser negativa.")
    if precio_lista < 0:
        return messagebox.showwarning("Validación", "El Precio de Lista no puede ser negativo.")
    if not (0 <= iva <= 100):
        return messagebox.showwarning("Validación", "El IVA debe estar entre 0 y 100 %.")
    if bnf < 0:
        return messagebox.showwarning("Validación", "El BNF no puede ser negativo.")

    # PRE-CHECK de duplicado al crear
    if current_id is None:
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                f"SELECT COUNT(*) FROM {TABLE_NAME} WHERE codigo = ?",
                (codigo,)
            )
            if cur.fetchone()[0] > 0:
                messagebox.showwarning(
                    "Código duplicado",
                    f"Ya existe un producto con Cod. Art. = {codigo}.\n"
                    "Continuaremos de todas formas."
                )

    # Snapshot para deshacer
    snapshot()

    # Preparar datos
    descripcion      = entry_descripcion.get().strip()
    retiro           = entry_retiro.get().strip()
    precio_final     = round(precio_lista * (1 + iva/100), 3)
    importe          = round(cantidad * precio_final, 2)

    precio_lista_str = f"{precio_lista:.1f} $"
    precio_final_str = f"{precio_final:.3f} $"
    importe_str      = f"{importe:.2f} $"
    iva_str          = f"{iva} %"

    cat = CATEGORIES[current_cat_idx]

    # INSERT o UPDATE con orden
    try:
        with sqlite3.connect(DB_PATH) as conn:
            if current_id:
                # Al editar, no cambiamos el campo orden
                sql = f"""
                    UPDATE {TABLE_NAME}
                    SET categoria    = ?,
                        codigo       = ?,
                        descripcion  = ?,
                        cantidad     = ?,
                        precio_lista = ?,
                        iva          = ?,
                        bnf          = ?,
                        precio_final = ?,
                        importe      = ?,
                        fecha_retiro = ?
                    WHERE id = ?
                """
                params = (
                    cat, codigo, descripcion, cantidad,
                    precio_lista_str, iva_str, bnf,
                    precio_final_str, importe_str,
                    retiro, current_id
                )
            else:
                # Al insertar, calculamos next_orden para esta categoría
                cur = conn.execute(
                    f"SELECT COALESCE(MAX(orden), -1) FROM {TABLE_NAME} WHERE categoria = ?",
                    (cat,)
                )
                next_orden = cur.fetchone()[0] + 1

                sql = f"""
                    INSERT INTO {TABLE_NAME} (
                        categoria, codigo, descripcion, cantidad,
                        precio_lista, iva, bnf,
                        precio_final, importe,
                        fecha_retiro, orden
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """
                params = (
                    cat, codigo, descripcion, cantidad,
                    precio_lista_str, iva_str, bnf,
                    precio_final_str, importe_str,
                    retiro, next_orden
                )

            conn.execute(sql, params)
    except Exception as e:
        messagebox.showerror("Error al guardar", str(e))

    # Limpiar y recargar
    limpiar_form()
    cargar_datos()

# — Portapapeles interno para Copy/Paste de productos —
_clipboard = []  # aquí guardaremos los rows completos

def editar_producto():
    global current_id
    sel = tree.selection()
    if not sel:
        # Mensaje informativo si no hay nada seleccionado
        messagebox.showinfo(
            title="Editar",
            message="Por favor, seleccioná un producto primero."
        )
        return

    # Limpiamos el formulario
    limpiar_form()

    # Obtenemos el ID real del registro desde el IID de la fila
    current_id = int(sel[0])

    # Recuperamos los valores que están en el Treeview (solo las columnas visibles)
    vals = tree.item(sel[0])["values"]

    # Insertamos cada valor en su Entry correspondiente, respetando VISIBLE_COLUMNS
    for idx, col in enumerate(VISIBLE_COLUMNS):
        entries[col].insert(0, vals[idx])

def eliminar_producto():
    sel = tree.selection()
    if not sel:
        return messagebox.showinfo("Eliminar", "Por favor, seleccioná un producto primero.")
    iid = sel[0]
    pid = int(iid)  # aquí sí tomo directamente el IID como PK

    if not messagebox.askyesno("Eliminar", "¿Confirmar eliminación?"):
        return

    snapshot()  # guardo estado para poder deshacer

    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(f"DELETE FROM {TABLE_NAME} WHERE id=?", (pid,))

    limpiar_form()
    cargar_datos()

# --- INICIALIZACIÓN ---
ensure_dirs()
ensure_default_config()
init_db()

# Creamos la ventana
root = tk.Tk()
# Configuraciones de la ventana
root.title("Co-op Stock Manager")
root.geometry("1000x600")
# BARRA DE ESTADO
status_var = tk.StringVar()

# --- MODO OSCURO / CLARO Y MENÚ PERSONALIZADO ---
style = ttk.Style(root)
cfg = load_config()
current_theme = cfg.get("theme", "light")

# Estilo base para el botón de limpiar búsqueda
style.configure(
    "SearchClear.TButton",
    background="SystemButtonFace",
    foreground="black",
    relief="flat"
)
style.map(
    "SearchClear.TButton",
    background=[("active", "SystemHighlight")],
    foreground=[("active", "black")]
)

def toggle_theme():
    global current_theme
    if current_theme == "light":
        apply_dark()
        current_theme = "dark"
    else:
        apply_light()
        current_theme = "light"
    save_config({"theme": current_theme})

# ------------------ Ventana de Ayuda: "Cómo funciona" ------------------
pages = [
    ("Pantalla principal — resumen",
     "Breve:\n"
     "- Formulario superior: Cantidad | Cod. Art. (obligatorio) | Concepto | IVA | P. lista | BNF | Precio (calc.) | Importe (calc.) | Fecha.\n"
     "- Botones: Guardar, Editar, Eliminar, Limpiar.\n"
     "- Tabla (Treeview): muestra productos por categoría; las filas con bajo stock se resaltan.\n"
     "- Búsqueda: campo dinámico; selector de categoría ◀ ▶.\n"
     "- Barra de estado: total de productos y valor total del stock."
    ),

    ("Importar y exportar CSV: Lo esencial",
     "- Exporta todas las filas con las columnas: categoría, orden y las columnas visibles.\n"
     "- Los campos de precio e IVA se formatean con símbolos (por ejemplo: \"100.0 $\", \"21 %\").\n"
     "- Al importar se aceptan formatos con o sin símbolos y se normalizan (coma/punto).\n"
     "- Si faltan columnas obligatorias, la importación se aborta y se muestra un error.\n"
     "- Consejo: exporta desde la aplicación para obtener un CSV reimportable de forma fiable."
    ),

    ("Backup y restauración — resumen",
     "- Backup automático al cerrar: copia del archivo de base de datos en la carpeta backup/ con marca temporal.\n"
     "- Backup manual: Opciones → Hacer backup manual.\n"
     "- Restaurar: Opciones → Restaurar backup... (se pedirá confirmación antes de reemplazar los datos activos).\n"
     "- Recomendación: conserva copias externas si necesitas historial más largo."
    ),

    ("Orden y arrastrar/soltar",
     "- Reordena filas arrastrando y soltando dentro de la tabla o con botones para mover arriba/abajo.\n"
     "- El orden se guarda en la base de datos (columna 'orden') por categoría.\n"
     "- Al exportar, las filas se escriben ordenadas por categoría y orden."
    ),

    ("Deshacer y rehacer",
     "- La aplicación guarda snapshots antes de operaciones importantes (importar, guardar, eliminar).\n"
     "- Ctrl+Z deshace; Ctrl+Y rehace. El historial tiene un límite configurado (MAX_UNDO).\n"
     "- Si necesitas recuperar estados muy antiguos, usa los backups en backup/."
    ),

    ("Modo oscuro y claro: lo básico",
     "- Alterna en Opciones → Modo Oscuro/Claro; la preferencia se guarda automáticamente.\n"
     "- El tema afecta entradas, la tabla y colores de alerta; la interfaz se adapta al redimensionar.\n"
     "- Si ves texto poco visible, alternar el modo una vez suele forzar el redibujo."
    ),

    ("Atajos útiles",
     "- Ctrl+C — copiar selección interna.\n"
     "- Ctrl+V — pegar/duplicar en la categoría actual.\n"
     "- Ctrl+Z / Ctrl+Y — deshacer / rehacer.\n"
     "- Clic en el encabezado de una columna — ordenar por esa columna (clic repetido invierte el orden)."
    ),

    ("Actualizaciones",
     "Esta aplicación se encuentra en desarrollo. Aún no se sabe cuándo será la versión definitiva.\n\n"
     "Cada actualización puede incluir correcciones de errores, mejoras y nuevas características.\n"
     "El registro de cambios de cada versión se publicará en la release de cada actualización.\n\n"
     "Este es un proyecto personal; sin embargo, puedes enviar solicitudes que serán evaluadas.\n\n"
     "Para comprobar si tienes la versión más reciente, ve a: Ayuda y actualizaciones → Buscar actualizaciones.\n" 
     "O simplemente cuando inicies la aplicación, automáticamente buscará si hay una actualización disponible.\n\n"
     "Visita el repositorio en GitHub: https://github.com/SrBenja/Co-op_Stock_Manager\n\n"
    ),

    ("Preguntas frecuentes",
     "- ¿No funciona Deshacer? Es posible que el historial haya alcanzado su límite; usa backups si necesitas recuperar estados remotos.\n"
     "- ¿Dónde están los backups? En la carpeta backup/ dentro del directorio de la aplicación.\n"
     "- ¿Hay duplicados en Cod. Art.? La aplicación advierte al guardar si el código ya existe; revisa antes de insertar si quieres evitar duplicados.\n"
     "- ¿Límite de buscar actualizaciones alcanzado? Espere 60 minutos y vuelva a intentar. También puede agregar un token personal de GitHub en config/config.json. Edite el archivo y agregué en cualquier sitio lo siguiente: (agregue comillas dobles)github_token(agregue comillas dobles): (agregue comillas dobles)AQUÍ AGREGUE SU TOKEN(agregue comillas dobles)"
    )
]

def show_help_window():
    """Abrir la ventana de ayuda (TOC + contenido)."""
    # Si ya está abierta, la traemos al frente
    for w in root.winfo_children():
        if isinstance(w, tk.Toplevel) and w.title() == "Cómo funciona":
            w.lift()
            return

    win = tk.Toplevel(root)
    win.title("Cómo funciona")

    # Tamaño inicial y centrado
    init_w, init_h = 900, 600
    win.geometry(f"{init_w}x{init_h}")
    win.minsize(600, 400)
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - init_w) // 2
    y = (sh - init_h) // 2
    win.geometry(f"+{x}+{y}")

    win.transient(root)  # no modal

    # Ancho fijo del índice: 30
    max_chars = 30
    left_w = tk.Frame(win, width=max_chars * 8)
    left_w.pack(side="left", fill="y")
    left_w.pack_propagate(False)

    right_w = tk.Frame(win)
    right_w.pack(side="left", fill="both", expand=True)

    # ÍNDICE (Listbox) sin scrollbar
    lbl_toc = ttk.Label(left_w, text="Índice", anchor="w")
    lbl_toc.pack(fill="x", padx=8, pady=(8, 0))

    toc_list = tk.Listbox(left_w, exportselection=False, activestyle="none",
                          selectmode="browse", width=max_chars)
    toc_list.pack(fill="both", expand=True, padx=8, pady=(6,8))

    for i, (title, _) in enumerate(pages):
        toc_list.insert("end", f"{i+1}. {title}")

    # Header
    header = tk.Frame(right_w)
    header.pack(fill="x", padx=8, pady=(8, 0))
    title_lbl = ttk.Label(header, text="", font=("TkDefaultFont", 12, "bold"), anchor="w")
    title_lbl.pack(side="left", fill="x", expand=True)

    # Contenido (Text) sin scrollbar visible
    content_frame = tk.Frame(right_w)
    content_frame.pack(fill="both", expand=True, padx=8, pady=8)
    content_text = tk.Text(content_frame, wrap="word", state="disabled",
                           padx=8, pady=6, relief="flat", borderwidth=0)
    content_text.pack(fill="both", expand=True)

    # Aplicar colores iniciales según tema (si existe current_theme)
    try:
        if current_theme == "dark":
            left_bg = "#2e2e2e"
            left_fg = "white"
            sel_bg = "#3366AA"   # selección visible en oscuro
            sel_fg = "white"
            txt_bg = "#3e3e3e"
            txt_fg = "white"
            title_fg = "white"
        else:
            left_bg = "SystemButtonFace"
            left_fg = "black"
            sel_bg = "#cce6ff"   # selección visible en claro
            sel_fg = "black"
            txt_bg = "white"
            txt_fg = "black"
            title_fg = "black"
        left_w.configure(bg=left_bg)
        toc_list.configure(bg=left_bg, fg=left_fg, selectbackground=sel_bg, selectforeground=sel_fg, highlightthickness=0)
        content_text.configure(background=txt_bg, foreground=txt_fg, insertbackground=txt_fg)
        title_lbl.configure(foreground=title_fg)
    except Exception:
        pass

    # Footer centrado (Atrás | Siguiente | indicador)
    footer = tk.Frame(right_w)
    footer.pack(fill="x", padx=8, pady=(0,8))
    center_frame = tk.Frame(footer)
    center_frame.pack(expand=True)
    btn_back = ttk.Button(center_frame, text="◀ Atrás")
    btn_next = ttk.Button(center_frame, text="Siguiente ▶")
    page_lbl = ttk.Label(center_frame, text="0 / 0")
    btn_back.pack(side="left", padx=(0,6))
    btn_next.pack(side="left")
    page_lbl.pack(side="left", padx=(8,0))

    current_index = {"i": 0}

    # Inserta texto y detecta URLs simples (http...)
    def _insert_with_links_and_subtitle(text_widget, s, idx):
        text_widget.config(state="normal")
        text_widget.delete("1.0", "end")
        i = 0
        L = len(s)
        # Insertar texto y detectar URLs como antes
        while i < L:
            p = s.find("http", i)
            if p == -1:
                text_widget.insert("end", s[i:])
                break
            if p > i:
                text_widget.insert("end", s[i:p])
            j = p
            while j < L and not s[j].isspace():
                j += 1
            url = s[p:j].rstrip(".,);:")
            start_index = text_widget.index("end-1c")
            text_widget.insert("end", url)
            end_index = text_widget.index("end-1c")
            tag_name = f"link_{start_index.replace('.', '_')}"
            text_widget.tag_add(tag_name, start_index, end_index)
            text_widget.tag_config(tag_name, foreground="blue", underline=1)
            text_widget.tag_bind(tag_name, "<Enter>", lambda e, w=text_widget: w.config(cursor="hand2"))
            text_widget.tag_bind(tag_name, "<Leave>", lambda e, w=text_widget: w.config(cursor=""))
            def _open_url(event, url=url):
                try:
                    webbrowser.open(url)
                except Exception:
                    pass
            text_widget.tag_bind(tag_name, "<Button-1>", _open_url)
            i = j

        text_widget.config(state="disabled")

    def _show_index(idx):
        idx = max(0, min(len(pages) - 1, idx))
        current_index["i"] = idx
        title, body = pages[idx]
        title_lbl.config(text=title)
        _insert_with_links_and_subtitle(content_text, body, idx)
        toc_list.selection_clear(0, "end")
        toc_list.selection_set(idx)
        toc_list.see(idx)
        page_lbl.config(text=f"{idx+1} / {len(pages)}")
        if idx > 0:
            btn_back.state(["!disabled"])
        else:
            btn_back.state(["disabled"])
        if idx < len(pages) - 1:
            btn_next.state(["!disabled"])
        else:
            btn_next.state(["disabled"])

    def on_toc_select(event=None):
        sel = toc_list.curselection()
        if not sel:
            return
        _show_index(sel[0])

    def on_back():
        _show_index(current_index["i"] - 1)

    def on_next():
        _show_index(current_index["i"] + 1)

    btn_back.config(command=on_back)
    btn_next.config(command=on_next)
    toc_list.bind("<<ListboxSelect>>", on_toc_select)

    # -------------------------
    # Deseleccionar índice al hacer clic dentro de la ventana (solo para esta ventana)
    # -------------------------
    def _click_in_help(event):
        try:
            # Solo actuamos si el toplevel del widget clicado es esta ventana 'win'
            if event.widget.winfo_toplevel() is win:
                # Si el clic no fue dentro del toc_list, deseleccionamos
                # (si el clic fue en toc_list, la selección debe mantenerse)
                w = event.widget
                is_in_toc = False
                while w:
                    if w is toc_list:
                        is_in_toc = True
                        break
                    if w is win:
                        break
                    w = getattr(w, "master", None)
                if not is_in_toc:
                    try:
                        toc_list.selection_clear(0, "end")
                    except Exception:
                        pass
        except Exception:
            pass

    # Bind local: bind a todos los widgets hijos de win mediante bind_all pero filtramos por toplevel==win
    # De esta forma no interfiere con otras ventanas; la acción solo se ejecuta cuando el clic pertenece a 'win'.
    win.bind_all("<Button-1>", _click_in_help, add="+")

    # También quitar la selección si el usuario escribe en el content_text
    def _on_content_key(event=None):
        try:
            toc_list.selection_clear(0, "end")
        except Exception:
            pass
    content_text.bind("<Key>", _on_content_key)

    # -------------------------
    # Actualizar colores cuando cambie current_theme
    # -------------------------
    def _apply_theme(ct):
        try:
            if ct == "dark":
                left_bg = "#2e2e2e"
                left_fg = "white"
                sel_bg = "#3366AA"
                sel_fg = "white"
                txt_bg = "#3e3e3e"
                txt_fg = "white"
                title_fg = "white"
            else:
                left_bg = "SystemButtonFace"
                left_fg = "black"
                sel_bg = "#cce6ff"
                sel_fg = "black"
                txt_bg = "white"
                txt_fg = "black"
                title_fg = "black"
            left_w.configure(bg=left_bg)
            toc_list.configure(bg=left_bg, fg=left_fg, selectbackground=sel_bg, selectforeground=sel_fg)
            content_text.configure(background=txt_bg, foreground=txt_fg, insertbackground=txt_fg)
            title_lbl.configure(foreground=title_fg)
        except Exception:
            pass

    def _poll_theme(prev=[None]):
        try:
            ct = current_theme
        except Exception:
            ct = None
        if ct != prev[0]:
            prev[0] = ct
            _apply_theme(ct)
        try:
            win.after(500, _poll_theme)
        except Exception:
            pass

    # Inicializar
    _show_index(0)
    _poll_theme()

    # Forzar foco en la ventana de ayuda (no roba foco de manera disruptiva)
    try:
        win.focus_force()
    except Exception:
        pass

# ------------------ fin de la ventana de Ayuda ------------------


# — Definición de ítems para cada menú —
file_items = [
    ("Importar CSV",    importar_csv,      False),
    ("Exportar CSV",    exportar_csv,      False),
    ("Imprimir",        imprimir_stock,    False),
]
opt_items = [
    ("Modo Oscuro/Claro", toggle_theme,     False),
    ("Restaurar backup...", restore_backup, False),
    ("Hacer backup manual", manual_backup,  False),
]
help_items = [
    ("Buscar actualizaciones", lambda: check_updates(), False),
    ("Cómo funciona", lambda: show_help_window(), False),
]

def _download_file(url, dst, timeout=30):
    """Descarga con streaming para no quedarnos sin memoria."""
    req = urllib.request.Request(url, headers={"User-Agent": "Coop-Stock-Updater"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dst, "wb") as out:
        chunk = resp.read(8192)
        while chunk:
            out.write(chunk)
            chunk = resp.read(8192)


def download_and_install_release_exe(repo, preferred_asset_name=None):
    """
        Updater:
      - descarga silenciosa del asset .exe en BACKUP_PATH,
      - crea un .bat temporal que espera al PID, reemplaza el exe y lanza un VBS limpiador,
      - lanza el .bat de forma oculta (no muestra consola),
      - no arranca la app nueva (debe abrirse manualmente),
      - todos los try/except están correctamente emparejados.
    """
    # Consultar la última release
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    try:
        req = urllib.request.Request(api_url, headers={"User-Agent": f"Co-op_Stock_Manager/{VERSION}"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            j = _json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        messagebox.showerror("Actualizaciones", f"No se pudo consultar GitHub:\n{e}")
        return

    latest_tag = j.get("tag_name") or j.get("name") or ""
    if not latest_tag:
        messagebox.showinfo("Actualizaciones", "No se encontró etiqueta en la última release.")
        return
    if _norm_tag(latest_tag) == _norm_tag(VERSION):
        messagebox.showinfo("Actualizaciones", f"Estás en la última versión: {VERSION}")
        return

    # Seleccionar asset .exe
    assets = j.get("assets", []) or []
    chosen = None
    for a in assets:
        name = a.get("name", "")
        if preferred_asset_name and name == preferred_asset_name:
            chosen = a
            break
    if not chosen:
        for a in assets:
            if a.get("name", "").lower().endswith(".exe"):
                chosen = a
                break
    if not chosen:
        messagebox.showinfo("Actualizaciones", "No se encontró un ejecutable en la última release.")
        return

    download_url = chosen.get("browser_download_url")
    asset_name = chosen.get("name")

    # Diálogo inicial: confirmar descarga
    initial_msg = (
        f"Versión instalada: {VERSION}\n"
        f"Versión disponible: {latest_tag}\n\n"
        f"Ejecutable: {asset_name}\n\n"
        "¿Descargar la actualización ahora?"
    )
    if not messagebox.askyesno("Actualizar - Descargar", initial_msg):
        return

    # Descargar silenciosamente
    try:
        os.makedirs(BACKUP_PATH, exist_ok=True)
    except Exception:
        pass

    tmp_path = os.path.join(BACKUP_PATH, f"update_{int(time.time())}_{asset_name}")
    if not tmp_path.lower().endswith(".exe"):
        tmp_path = tmp_path + ".exe"

    try:
        _download_file(download_url, tmp_path)
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        messagebox.showerror("Actualizaciones", f"Error al descargar la actualización:\n{e}")
        return

    # Diálogo final: confirmar aplicar (mensaje claro)
    final_msg = (
        "La actualización se descargó correctamente.\n\n"
        "La aplicación se cerrará ahora para aplicar la actualización.\n\n"
        "Deberás iniciarla manualmente después de que termine el proceso de actualización.\n\n"
        "¿Continuar y cerrar la aplicación ahora?"
    )
    if not messagebox.askokcancel("Actualizar - Aplicar ahora", final_msg):
        return

    # Preparar paths y PID
    try:
        if getattr(sys, "frozen", False):
            exe_path = sys.executable
        else:
            exe_path = os.path.abspath(sys.argv[0])
        backup_path = exe_path + ".bak"
        current_pid = os.getpid()
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        messagebox.showerror("Actualizaciones", f"Error determinando paths para la actualización:\n{e}")
        return

    # Evitar reemplazar python.exe en modo dev
    try:
        if os.path.basename(exe_path).lower().startswith("python"):
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            messagebox.showerror(
                "Actualizaciones",
                "La aplicación parece ejecutarse con python.exe (modo desarrollo). Empaqueta la app como .exe para usar el actualizador."
            )
            return
    except Exception:
        pass

    # Crear .bat y .vbs
    batch_path = os.path.join(BACKUP_PATH, f"updater_{int(time.time())}.bat")
    vbs_path = os.path.join(BACKUP_PATH, f"cleaner_{int(time.time())}.vbs")

    batch_contents = f"""@echo off
setlocal enableextensions enabledelayedexpansion
set "NEW={tmp_path}"
set "TARGET={exe_path}"
set "BACKUP={backup_path}"
set "PID={current_pid}"

REM esperar a que el PID termine
:wait_pid
tasklist /FI "PID eq %PID%" 2>NUL | findstr /R /C:"%PID%" >NUL
if "%ERRORLEVEL%"=="0" (
    timeout /t 1 /nobreak >NUL
    goto wait_pid
)

REM intento de reemplazo con reintentos
set /A tries=0
:copy_try
if exist "%TARGET%" (
    move /Y "%TARGET%" "%BACKUP%" >NUL 2>&1
)
copy /Y "%NEW%" "%TARGET%" >NUL 2>&1
if exist "%TARGET%" (
    if exist "%NEW%" del /F /Q "%NEW%" >NUL 2>&1
    goto launch_cleaner
) else (
    set /A tries+=1
    if %tries% GEQ 20 goto final_try
    timeout /t 1 /nobreak >NUL
    goto copy_try
)

:final_try
move /Y "%NEW%" "%TARGET%" >NUL 2>&1
if exist "%NEW%" del /F /Q "%NEW%" >NUL 2>&1

:launch_cleaner
REM Llamar al VBS invisible que borrará BACKUP y este batch
wscript.exe "{vbs_path}" "%BACKUP%" "%~f0"
exit /B 0
"""

    vbs_contents = (
        'Set args = WScript.Arguments\n'
        'backup = args(0)\n'
        'batch = args(1)\n'
        'WScript.Sleep 1000\n'
        'On Error Resume Next\n'
        'Set fso = CreateObject("Scripting.FileSystemObject")\n'
        'tries = 0\n'
        'Do While tries < 15\n'
        '  If fso.FileExists(backup) Then\n'
        '    fso.DeleteFile backup, True\n'
        '  End If\n'
        '  If Not fso.FileExists(backup) Then Exit Do\n'
        '  WScript.Sleep 1000\n'
        '  tries = tries + 1\n'
        'Loop\n'
        'On Error Resume Next\n'
        'If fso.FileExists(batch) Then\n'
        '  On Error Resume Next\n'
        '  fso.DeleteFile batch, True\n'
        'End If\n'
        'On Error Resume Next\n'
        'If fso.FileExists(WScript.ScriptFullName) Then\n'
        '  On Error Resume Next\n'
        '  fso.DeleteFile WScript.ScriptFullName, True\n'
        'End If\n'
    )

    # Escribir archivos al disco
    try:
        with open(batch_path, "w", encoding="utf-8") as f:
            f.write(batch_contents)
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs_contents)
    except Exception as e:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        messagebox.showerror("Actualizaciones", f"No se pudo preparar el actualizador:\n{e}")
        return

    # Helper: intento directo como último recurso (función local, maneja sus excepciones)
    def _attempt_direct_replace(tmp_p, exe_p, bak_p):
        try:
            # mover exe actual a bak (si existe)
            if os.path.exists(exe_p):
                try:
                    os.replace(exe_p, bak_p)
                except Exception:
                    try:
                        os.rename(exe_p, bak_p)
                    except Exception:
                        pass
            # mover tmp a target
            try:
                os.replace(tmp_p, exe_p)
                # intento borrar bak
                try:
                    if os.path.exists(bak_p):
                        os.remove(bak_p)
                except Exception:
                    pass
                return True
            except Exception:
                # cleanup tmp si no se pudo mover
                try:
                    if os.path.exists(tmp_p):
                        os.remove(tmp_p)
                except Exception:
                    pass
                return False
        except Exception:
            return False

    # Lanzar el .bat de forma oculta: preferimos cmd /c call con CREATE_NO_WINDOW (shell=False)
    started_batch = False
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(['cmd', '/c', 'call', batch_path], shell=False, creationflags=creationflags)
        started_batch = True
    except Exception:
        # fallback: intentar shell=True con call
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            cmdline = f'call "{batch_path}"'
            subprocess.Popen(cmdline, shell=True, creationflags=creationflags)
            started_batch = True
        except Exception:
            started_batch = False

    # Si no logramos lanzar el batch, intentamos replace directo como último recurso
    if not started_batch:
        _attempt_direct_replace(tmp_path, exe_path, backup_path)

    # Cerrar la app para liberar handles (batch/VBS harán la limpieza)
    try:
        root.destroy()
    except Exception:
        pass
    os._exit(0)


# — Creamos el frame que contendrá los Menubuttons —
menubar_frame = tk.Frame(root)
menubar_frame.pack(fill="x")

def create_menu_button(text, items):
    # Estilos según tema
    bg = "#2e2e2e" if current_theme == "dark" else "SystemButtonFace"
    fg = "white"     if current_theme == "dark" else "black"

    # Creamos el Menubutton
    mb = tk.Menubutton(
        menubar_frame,
        text=text,
        relief="flat",
        bg=bg, fg=fg,
        activebackground=bg,
        activeforeground=fg
    )
    mb.pack(side="left", padx=2, pady=2)

    # Instanciamos aquí su DropdownMenu usando el Menubutton como master
    dropdown = DropdownMenu(mb, items)

    # Bind de apertura: cierra los demás y muestra este
    def open_it(e):
        mb.config(relief="flat")
        x = mb.winfo_rootx()
        y = mb.winfo_rooty() + mb.winfo_height()
        dropdown.show(x, y)

    mb.bind("<Button-1>", open_it)
    return mb

# — Finalmente, creamos cada botón ligado a su lista desplegable —
file_mb = create_menu_button("Archivo", file_items)   # dentro usa DropdownMenu(mb, items)
opt_mb  = create_menu_button("Opciones", opt_items)
help_mb = create_menu_button("Ayuda y actualizaciones", help_items)

# Helper
def _force_entry_caret_and_selection(entry, insert_color, select_bg, select_fg, insert_width=2):
    """Forza opciones de caret y selección en un widget Entry o ttk.Entry."""
    if entry is None:
        return
    # intento normal (funciona para tk.Entry y a veces para ttk.Entry)
    try:
        entry.configure(insertbackground=insert_color)
    except Exception:
        pass
    try:
        entry.configure(selectbackground=select_bg)
    except Exception:
        pass
    try:
        entry.configure(selectforeground=select_fg)
    except Exception:
        pass
    # intento nativo vía Tk (funciona para ttk.Entry)
    try:
        entry.tk.call(entry._w, 'configure', '-insertbackground', insert_color)
    except Exception:
        pass
    try:
        entry.tk.call(entry._w, 'configure', '-selectbackground', select_bg)
    except Exception:
        pass
    try:
        entry.tk.call(entry._w, 'configure', '-selectforeground', select_fg)
    except Exception:
        pass
    try:
        entry.tk.call(entry._w, 'configure', '-insertwidth', str(int(insert_width)))
    except Exception:
        pass
    # update visual
    try:
        entry.update_idletasks()
    except Exception:
        pass


def _darken_hex(hex_color: str, factor: float = 0.6) -> str:
    """Devuelve una versión más oscura de hex_color (ej. '#FF0000'), factor entre 0 y 1."""
    try:
        s = hex_color.strip()
        if s.startswith("#"):
            s = s[1:]
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        r = max(0, min(255, int(r * factor)))
        g = max(0, min(255, int(g * factor)))
        b = max(0, min(255, int(b * factor)))
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return hex_color


def apply_light():
    # Ventana principal y estilo base
    root.configure(bg="SystemButtonFace")
    style.theme_use("clam")

    # ENTRADAS
    style.configure("TEntry", fieldbackground="white", foreground="black")
    style.map("TEntry",
        fieldbackground=[("!disabled", "white"), ("focus", "white")],
        foreground=[("!disabled", "black"), ("focus", "black")]
    )
    try:
        entry_search.config(bg="white", fg="black", insertbackground="black")
    except Exception:
        pass
    for ent in entries.values():
        try: ent.configure(background="white", foreground="black")
        except Exception: pass

    # CARET / SELECCIÓN (helper)
    def _apply_light_later():
        for ent in entries.values():
            _force_entry_caret_and_selection(ent, insert_color="black", select_bg="#cce6ff", select_fg="black", insert_width=2)
        try:
            _force_entry_caret_and_selection(entry_search, insert_color="black", select_bg="#cce6ff", select_fg="black", insert_width=2)
        except Exception:
            pass
    _apply_light_later()
    try: root.after(100, _apply_light_later)
    except Exception: pass

    # -------------------------
    # BOTONES: creamos/actualizamos estilo Action.TButton (para todos los botones de acción)
    action_style = "Action.TButton"
    hover_bg = "SystemHighlight"
    normal_bg = "SystemButtonFace"
    style.configure(action_style,
                    background=normal_bg,
                    foreground="black",
                    relief="raised",
                    padding=4)
    style.map(action_style,
              background=[("active", hover_bg), ("!disabled", normal_bg)],
              foreground=[("active", "black"), ("!disabled", "black")],
              relief=[("active", "raised"), ("!disabled", "raised")])
    # Asignar este estilo a todos los ttk.Buttons dentro de los contenedores relevantes
    parents = [frm, content, search_frame, cat_select_frame]
    for p in parents:
        try:
            for w in p.winfo_children():
                # Si es ttk.Button, forzamos style = Action.TButton
                if isinstance(w, ttk.Button):
                    try:
                        w.configure(style=action_style)
                    except Exception:
                        pass
                # Si es tk.Button, configuramos activebackground como fallback
                elif isinstance(w, tk.Button):
                    try:
                        if not getattr(w, "_action_btn_bound", False):
                            try:
                                orig_bg = w.cget("background")
                            except Exception:
                                orig_bg = normal_bg
                            # enter/leave para refuerzo visual en tk.Button
                            def _tk_enter(e, widget=w, bg=hover_bg):
                                try: widget.config(background=bg)
                                except Exception: pass
                            def _tk_leave(e, widget=w, bg=orig_bg):
                                try: widget.config(background=bg)
                                except Exception: pass
                            w.bind("<Enter>", _tk_enter, add="+")
                            w.bind("<Leave>", _tk_leave, add="+")
                            # también configura activebackground para que se muestre el clic nativo
                            try: w.config(activebackground=hover_bg, activeforeground="black")
                            except Exception: pass
                            w._action_btn_bound = True
                    except Exception:
                        pass
        except Exception:
            pass

    # -------------------------
    # MENUBUTTONS: paleta arcoíris on-hover, y forzamos estado por defecto según tema
    menubar_frame.config(bg="SystemButtonFace")
    # restaurar paleta original (clara)
    try:
        DropdownMenu.RAINBOW = ["#FF0000", "#FFA500", "#FFFF00", "#008000", "#0000FF", "#4B0082", "#EE82EE"]
    except Exception:
        pass

    try:
        rainbow = DropdownMenu.RAINBOW
        targets = (file_mb, opt_mb, help_mb)
        for i, mb in enumerate(targets):
            color = rainbow[i % len(rainbow)]
            # ON ENTER: arcoíris (vivo)
            def _on_enter(e, m=mb, c=color):
                try: m.config(bg=c, fg="black", activebackground=c, activeforeground="black")
                except Exception: pass
            # ON LEAVE: restaurar al fondo del tema (light)
            def _on_leave(e, m=mb):
                try: m.config(bg="SystemButtonFace", fg="black", activebackground="SystemButtonFace", activeforeground="black")
                except Exception: pass
            # ON CLICK: mantener el color mientras se abre
            def _on_click(e, m=mb, c=color):
                try: m.config(bg=c, fg="black", activebackground=c, activeforeground="black")
                except Exception: pass
            # ON FOCUS_OUT: restaurar
            def _on_focus_out(e, m=mb):
                try: m.config(bg="SystemButtonFace", fg="black", activebackground="SystemButtonFace", activeforeground="black")
                except Exception: pass

            # añadimos binds sin remover <Button-1> original que abre el dropdown
            mb.bind("<Enter>", _on_enter, add="+")
            mb.bind("<Leave>", _on_leave, add="+")
            mb.bind("<Button-1>", _on_click, add="+")
            mb.bind("<FocusOut>", _on_focus_out, add="+")
            mb._hover_bound = True
            # forzamos color por defecto acorde al tema (esto arregla el problema del "pegado")
            try: mb.config(bg="SystemButtonFace", fg="black", activebackground="SystemButtonFace", activeforeground="black")
            except Exception: pass
    except Exception:
        pass

    # -------------------------
    # RESTO DE LA IU
    search_frame.config(bg="SystemButtonFace")
    label_search.config(bg="SystemButtonFace", fg="black")
    style.configure("TFrame", background="SystemButtonFace")
    style.configure("TLabel", background="SystemButtonFace", foreground="black")
    style.configure("Treeview", background="white", fieldbackground="white", foreground="black")
    style.map("Treeview", background=[("selected", "#cce6ff")], foreground=[("selected", "black")])
    tree.tag_configure("bajo_stock", background="#ffdddd", foreground="black")

    try: root.update_idletasks()
    except Exception: pass


def apply_dark():
    # Ventana principal y estilo base
    root.configure(bg="#2e2e2e")
    style.theme_use("clam")

    # ENTRADAS
    style.configure("TEntry", fieldbackground="#3e3e3e", foreground="white")
    style.map("TEntry",
        fieldbackground=[("!disabled", "#3e3e3e"), ("focus", "#3e3e3e")],
        foreground=[("!disabled", "white"), ("focus", "white")]
    )
    try:
        entry_search.config(bg="#3e3e3e", fg="white", insertbackground="white")
    except Exception:
        pass
    for ent in entries.values():
        try: ent.configure(background="#3e3e3e", foreground="white")
        except Exception: pass

    # CARET / SELECCIÓN
    def _apply_dark_later():
        for ent in entries.values():
            _force_entry_caret_and_selection(ent, insert_color="white", select_bg="#3366AA", select_fg="white", insert_width=2)
        try:
            _force_entry_caret_and_selection(entry_search, insert_color="white", select_bg="#3366AA", select_fg="white", insert_width=2)
        except Exception:
            pass
    _apply_dark_later()
    try: root.after(100, _apply_dark_later)
    except Exception: pass

    # -------------------------
    # BOTONES: estilo Action.TButton para dark
    action_style = "Action.TButton"
    normal_bg = "#3e3e3e"
    hover_bg = "#505050"
    style.configure(action_style,
                    background=normal_bg,
                    foreground="white",
                    relief="raised",
                    padding=4)
    style.map(action_style,
              background=[("active", hover_bg), ("!disabled", normal_bg)],
              foreground=[("active", "white"), ("!disabled", "white")],
              relief=[("active", "raised"), ("!disabled", "raised")])
    # asignar estilo a ttk.Buttons y configurar tk.Buttons
    parents = [frm, content, search_frame, cat_select_frame]
    for p in parents:
        try:
            for w in p.winfo_children():
                if isinstance(w, ttk.Button):
                    try: w.configure(style=action_style)
                    except Exception: pass
                elif isinstance(w, tk.Button):
                    try:
                        if not getattr(w, "_action_btn_bound", False):
                            try: orig_bg = w.cget("background")
                            except Exception: orig_bg = normal_bg
                            def _tk_enter(e, widget=w, bg=hover_bg):
                                try: widget.config(background=bg)
                                except Exception: pass
                            def _tk_leave(e, widget=w, bg=orig_bg):
                                try: widget.config(background=bg)
                                except Exception: pass
                            w.bind("<Enter>", _tk_enter, add="+")
                            w.bind("<Leave>", _tk_leave, add="+")
                            try: w.config(activebackground=hover_bg, activeforeground="white")
                            except Exception: pass
                            w._action_btn_bound = True
                    except Exception: pass
        except Exception: pass

    # -------------------------
    # MENUBUTTONS: fondo neutro y rainbow oscurecido on-hover
    menubar_frame.config(bg="#2e2e2e")
    try:
        base_rainbow = ["#FF0000", "#FFA500", "#FFFF00", "#008000", "#0000FF", "#4B0082", "#EE82EE"]
        dark_rainbow = [_darken_hex(c, 0.45) for c in base_rainbow]
        DropdownMenu.RAINBOW = dark_rainbow
    except Exception:
        pass

    try:
        rainbow = DropdownMenu.RAINBOW
        targets = (file_mb, opt_mb, help_mb)
        for i, mb in enumerate(targets):
            color = rainbow[i % len(rainbow)]
            def _on_enter(e, m=mb, c=color):
                try: m.config(bg=c, fg="white", activebackground=c, activeforeground="white")
                except Exception: pass
            def _on_leave(e, m=mb):
                try: m.config(bg="#2e2e2e", fg="white", activebackground="#5a5a5a", activeforeground="white")
                except Exception: pass
            def _on_click(e, m=mb, c=color):
                try: m.config(bg=c, fg="white", activebackground=c, activeforeground="white")
                except Exception: pass
            def _on_focus_out(e, m=mb):
                try: m.config(bg="#2e2e2e", fg="white", activebackground="#5a5a5a", activeforeground="white")
                except Exception: pass

            mb.bind("<Enter>", _on_enter, add="+")
            mb.bind("<Leave>", _on_leave, add="+")
            mb.bind("<Button-1>", _on_click, add="+")
            mb.bind("<FocusOut>", _on_focus_out, add="+")
            mb._hover_bound = True
            # forzamos estado por defecto acorde al tema
            try: mb.config(bg="#2e2e2e", fg="white", activebackground="#5a5a5a", activeforeground="white")
            except Exception: pass
    except Exception: pass

    # -------------------------
    # RESTO DE LA IU
    search_frame.config(bg="#2e2e2e")
    label_search.config(bg="#2e2e2e", fg="white")
    style.configure("TFrame", background="#2e2e2e")
    style.configure("TLabel", background="#2e2e2e", foreground="white")
    style.configure("Treeview", background="#3e3e3e", fieldbackground="#3e3e3e", foreground="white")
    style.map("Treeview", background=[("selected", "#3366AA")], foreground=[("selected", "white")])
    tree.tag_configure("bajo_stock", background="#800000", foreground="white")

    try: root.update_idletasks()
    except Exception: pass



# — CONTENIDO PRINCIPAL bajo la franja de menú —
content = ttk.Frame(root)
content.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
content.columnconfigure(0, weight=1)
content.rowconfigure(2, weight=1)

# BLOQUE DE BÚSQUEDA
MIN_SEARCH_WIDTH = 15
MAX_SEARCH_WIDTH = 50
# Hacemos search_frame hijo de content
search_frame = tk.Frame(content)
search_frame.grid(row=1, column=0, sticky="ew", pady=(0,5))

# Etiqueta
label_search = tk.Label(search_frame, text="Buscar:")
label_search.pack(side=tk.LEFT, padx=(0,5))

# Entry de búsqueda con ancho dinámico
entry_search = tk.Entry(search_frame, width=MIN_SEARCH_WIDTH)
entry_search.pack(side=tk.LEFT, padx=(0,5))

def on_search_key(event=None):
    txt = entry_search.get()
    new_w = min(MAX_SEARCH_WIDTH, max(MIN_SEARCH_WIDTH, len(txt) + 1))
    entry_search.config(width=new_w)
    buscar()

entry_search.bind("<KeyRelease>", on_search_key)


# — Variables de categoría —
CATEGORIES = ["Plomería", "Gas", "Electricidad"]
current_cat_idx = 0

# — Portapapeles interno para Copy/Paste de productos —
_clipboard = []

def copiar_seleccion(event=None):
    """Guarda en _clipboard las filas seleccionadas, sin mostrar mensajes."""
    global _clipboard
    sel = tree.selection()
    if not sel:
        return
    _clipboard = []
    with sqlite3.connect(DB_PATH) as conn:
        for iid in sel:
            row = conn.execute(
                f"SELECT {', '.join(COLUMNS)} FROM {TABLE_NAME} WHERE id = ?",
                (iid,)
            ).fetchone()
            if row:
                _clipboard.append(row)

def pegar_seleccion(event=None):
    """Duplica en la categoría actual los productos guardados en _clipboard de forma segura."""
    if not _clipboard:
        return

    nuevo_cat = CATEGORIES[current_cat_idx]
    with sqlite3.connect(DB_PATH) as conn:
        # Determinamos el siguiente orden disponible
        cur = conn.execute(
            f"SELECT COALESCE(MAX(orden), -1) FROM {TABLE_NAME} WHERE categoria = ?",
            (nuevo_cat,)
        )
        next_orden = cur.fetchone()[0] + 1

        insert_sql = f"""
            INSERT INTO {TABLE_NAME} (
                categoria, codigo, descripcion, cantidad,
                precio_lista, iva, bnf,
                precio_final, importe,
                fecha_retiro, orden
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """

        rows_to_insert = []
        for row in _clipboard:
            # row es una tupla con columnas en el orden de COLUMNS
            def _get(col):
                v = row[COLUMNS.index(col)]
                if v is None:
                    return ""
                return v

            # valores textuales
            codigo      = str(_get("codigo")) or ""
            descripcion = str(_get("descripcion")) or ""
            retiro      = str(_get("fecha_retiro")) or ""

            # valores numéricos: usamos _parse_number_from_db para mayor tolerancia
            try:
                cantidad_val = int(_parse_number_from_db(_get("cantidad")))
            except Exception:
                cantidad_val = 0
            try:
                iva_val = int(_parse_number_from_db(_get("iva")))
            except Exception:
                iva_val = 0
            try:
                precio_lista_val = float(_parse_number_from_db(_get("precio_lista")))
            except Exception:
                precio_lista_val = 0.0
            try:
                bnf_val = float(_parse_number_from_db(_get("bnf")))
            except Exception:
                bnf_val = 0.0
            try:
                precio_final_val = float(_parse_number_from_db(_get("precio_final")))
            except Exception:
                precio_final_val = 0.0
            try:
                importe_val = float(_parse_number_from_db(_get("importe")))
            except Exception:
                importe_val = 0.0

            # Formateos tal como los usa el resto de la app
            precio_lista_str = f"{precio_lista_val:.1f} $"
            precio_final_str = f"{precio_final_val:.3f} $"
            importe_str      = f"{importe_val:.2f} $"
            iva_str          = f"{iva_val} %"

            rows_to_insert.append((
                nuevo_cat,
                codigo,
                descripcion,
                cantidad_val,
                precio_lista_str,
                iva_str,
                bnf_val,
                precio_final_str,
                importe_str,
                retiro,
                next_orden
            ))
            next_orden += 1

        # Ejecutamos la inserción
        try:
            conn.executemany(insert_sql, rows_to_insert)
            conn.commit()
        except Exception as e:
            messagebox.showerror("Pegar selección", f"No se pudo pegar la selección:\n{e}")
            return

    # Refrescamos vista
    cargar_datos()

# — Función de cambio de categoría —
def switch_category(delta):
    global current_cat_idx
    current_cat_idx = (current_cat_idx + delta) % len(CATEGORIES)
    lbl_cat.config(text=CATEGORIES[current_cat_idx])
    cargar_datos()

# — Construcción del frame y botones — 
cat_select_frame = ttk.Frame(search_frame)
cat_select_frame.place(relx=0.5, rely=0.5, anchor="center")

btn_prev = ttk.Button(cat_select_frame, text="◀", width=2,
                      command=lambda: switch_category(-1))
lbl_cat  = ttk.Label(cat_select_frame, text=CATEGORIES[current_cat_idx],
                     width=12, anchor="center")
btn_next = ttk.Button(cat_select_frame, text="▶", width=2,
                      command=lambda: switch_category(+1))

btn_prev.pack(side="left")
lbl_cat.pack(side="left", padx=5)
btn_next.pack(side="left", padx=(0,10))

def move_selected(delta):
    sel = tree.selection()
    if not sel:
        return
    iid    = sel[0]
    idx    = tree.index(iid)
    newidx = max(0, min(len(tree.get_children())-1, idx+delta))
    if newidx == idx:
        return
    tree.move(iid, '', newidx)
    # Persistimos el nuevo orden en BD:
    with sqlite3.connect(DB_PATH) as conn:
        for pos, iid2 in enumerate(tree.get_children('')):
            conn.execute(f"UPDATE {TABLE_NAME} SET orden = ? WHERE id = ?", (pos, iid2))
    # Opcional: mantén la selección
    tree.selection_set(iid)

# Función única de cierre con backup
def on_closing():
    # Antes de salir, hacemos un backup automático
    backup_db()
    root.destroy()
# Asignamos esa función al evento de cierre
root.protocol("WM_DELETE_WINDOW", on_closing)

# FORMULARIO DE ENTRADA (row=0 en content)
frm = ttk.Frame(content, padding=(0,5))
frm.grid(row=0, column=0, sticky="ew", pady=(0,5))
# Configuramos cada columna interna para que crezca por igual
for i in range(len(VISIBLE_COLUMNS)):
    frm.columnconfigure(i, weight=1)
# Creamos un dict para guardar referencias a cada Entry
entries = {}
# Generamos labels y entradas expandibles
for idx, col in enumerate(VISIBLE_COLUMNS):
    ttk.Label(frm, text=COLUMN_LABELS[col]) \
       .grid(row=0, column=idx, sticky="ew", padx=2)
    # Usar tk.Entry en lugar de ttk.Entry para tener control completo del caret
    ent = tk.Entry(frm, relief="solid", borderwidth=1)
    ent.grid(row=1, column=idx, sticky="ew", padx=2, pady=2)
    entries[col] = ent


# Recuperamos los entry desde el dict
entry_cantidad     = entries["cantidad"]
entry_codigo       = entries["codigo"]
entry_descripcion  = entries["descripcion"]
entry_iva          = entries["iva"]
entry_precio_lista = entries["precio_lista"]
entry_bnf          = entries["bnf"]
entry_precio_final = entries["precio_final"]
entry_importe      = entries["importe"]
entry_retiro       = entries["fecha_retiro"]

# — Deseleccionar la fila al hacer clic en cualquier campo, sin interferir con el foco —
def _deselect_tree(e):
    tree.selection_remove(tree.selection())

for ent in entries.values():
    ent.bind("<Button-1>", _deselect_tree, add="+")

entry_search.bind("<Button-1>", _deselect_tree, add="+")

# Bindings y formato de precio
def al_salir_p_lista(e=None):
    texto = entry_precio_lista.get().strip().rstrip(" $").replace(",", ".")
    if not texto:
        return
    try:
        pl = float(texto)
    except ValueError:
        return
    # Formateamos con una decimal y añadimos el símbolo
    entry_precio_lista.delete(0, tk.END)
    entry_precio_lista.insert(0, f"{pl:.1f} $")
    # Y recalculamos el precio e importe
    calcular_importe()


# Reemplaza cualquier binding antiguo:
entry_precio_lista.bind("<FocusOut>", al_salir_p_lista)

# Cálculo en tiempo real mientras escribes en P. lista o IVA
entry_precio_lista.bind("<KeyRelease>", calcular_importe)
entry_iva.bind("<KeyRelease>", calcular_importe)

# También al perder foco en IVA (por si editan y salen rápido)
entry_iva.bind("<FocusOut>", calcular_importe)

# Botones de acción debajo del formulario (row=2 en frm)
btns = [
    ("Guardar",            guardar_producto),
    ("Editar",             editar_producto),
    ("Eliminar",           eliminar_producto)
]
# Botones de acción debajo del formulario (usar ttk.Button)
for i, (txt, cmd) in enumerate(btns):
    ttk.Button(frm, text=txt, command=cmd).grid(
        row=2, column=i, padx=5, pady=5
    )
# Botón Limpiar campos
ttk.Button(frm, text="Limpiar campos", command=limpiar_form) \
   .grid(row=2, column=3, padx=5, pady=5)

# --- Creación y configuración del Treeview ---
tree = ttk.Treeview(
    content,
    columns=VISIBLE_COLUMNS,
    show="headings"
)
# Configuración de columnas
for col in VISIBLE_COLUMNS:
    tree.heading(col, text=COLUMN_LABELS[col], command=lambda c=col: sort_column(c))
    tree.column(col, width=100, anchor="center", stretch=True)

# — Drag & Drop con feedback sólo en movimiento — 
_dragged_item     = None
_dragging_active  = False
_prev_cursor      = None

# tag para destacar la fila al arrastrar
tree.tag_configure("dragging", background="#cce6ff")

def on_drag_start(event):
    global _dragged_item, _dragging_active
    # ignoramos clicks en encabezado
    if tree.identify_region(event.x, event.y) == "heading":
        return
    iid = tree.identify_row(event.y)
    _dragged_item = iid if iid else None
    _dragging_active = False  # todavía no hemos movido

def on_drag_motion(event):
    global _dragging_active, _prev_cursor
    if not _dragged_item:
        return
    # primer movimiento: activamos el feedback
    if not _dragging_active:
        _prev_cursor = tree["cursor"]
        tree.configure(cursor="hand2")
        tree.item(_dragged_item, tags=("dragging",))
        _dragging_active = True
    # tree.yview_scroll(int((event.y - tree.winfo_height()/2)/20), "units")

def on_drag_drop(event):
    global _dragged_item, _dragging_active
    if _dragging_active and _dragged_item:
        target = tree.identify_row(event.y)
        if target and target != _dragged_item:
            new_index = tree.index(target)
            tree.move(_dragged_item, "", new_index)
            # persistimos el nuevo orden
            with sqlite3.connect(DB_PATH) as conn:
                for pos, iid in enumerate(tree.get_children("")):
                    conn.execute(
                        f"UPDATE {TABLE_NAME} SET orden = ? WHERE id = ?",
                        (pos, iid)
                    )
    # restauramos estado
    if _dragging_active:
        tree.configure(cursor=_prev_cursor or "")
        tree.item(_dragged_item, tags=())
    _dragged_item    = None
    _dragging_active = False

# Bindings sin sobreescribir otros
tree.bind("<ButtonPress-1>",   on_drag_start, add="+")
tree.bind("<B1-Motion>",       on_drag_motion, add="+")
tree.bind("<ButtonRelease-1>", on_drag_drop,  add="+")

tree.tag_configure("bajo_stock", background="#ffdddd")
tree.grid(row=2, column=0, sticky="nsew")

# Desseleccionar clics globalmente
def clear_all_selection(event):
    widget = event.widget
    # Si clicas dentro del Treeview...
    if widget is tree:
        # detecta la zona exacta del clic
        region = tree.identify('region', event.x, event.y)
        # solo si NO es sobre filas, encabezado, bordes, etc.
        if region == 'nothing':
            tree.selection_remove(tree.selection())
            root.focus_set()
        return

    # Si clicas dentro de un Entry o Button, no tocamos nada
    if isinstance(widget, (tk.Entry, ttk.Entry, tk.Button, ttk.Button)):
        return

    # Cualquier otro clic (en cualquier parte fuera del Treeview):
    tree.selection_remove(tree.selection())
    root.focus_set()

# — Barra de estado debajo del Treeview —
status_bar = ttk.Label(content, textvariable=status_var, anchor="w")
status_bar.grid(row=3, column=0, sticky="ew", pady=(5,0))

# Inicializamos su valor
update_status()

def _fetch_latest_release(repo, token=None, timeout=8):
    """
    Helper pequeño que consulta la API de GitHub y devuelve el JSON de la última release
    o None si hay cualquier error (silencioso).
    """
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    headers = {"User-Agent": f"Co-op_Stock_Manager/{VERSION} (+https://github.com/{repo})"}
    if token:
        headers["Authorization"] = f"token {token}"
    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return _json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        # Silencioso en el inicio: sólo registramos
        # if e.code == 403: print("GitHub rate limit or forbidden")
        # if e.code == 404: print("Repo/releases not found")
        return None
    except Exception:
        return None


def check_updates_on_startup(delay_ms=1500):
    """
    Comprobación silenciosa que se ejecuta 1 vez al iniciar (llamar con root.after).
    - No muestra mensajes si no hay actualizaciones o si ocurre un error.
    - Si encuentra una release con tag distinto a VERSION, muestra el diálogo modal personalizado
      (Sí / No / Cancelar) para que el usuario elija qué hacer.
    """
    try:
        cfg = load_config()
        default_repo = "SrBenja/Co-op_Stock_Manager"
        repo = cfg.get("github_repo", default_repo)
        token = cfg.get("github_token")

        j = _fetch_latest_release(repo, token=token)
        if not j:
            return  # silencioso si no hay datos

        latest_tag = (j.get("tag_name") or j.get("name") or "").strip()
        if not latest_tag:
            return

        if _norm_tag(latest_tag) == _norm_tag(VERSION):
            return  # estás en la misma versión -> silencio

        # hay una versión distinta: preparamos el mensaje resumido
        body = j.get("body", "") or ""
        html_url = j.get("html_url", "") or ""
        summary = (body[:800] + "...") if len(body) > 800 else body

        msg = (f"Versión instalada: {VERSION}\n"
               f"Versión disponible: {latest_tag}\n\n"
               f"{summary}\n\n"
               "¿Descargar e instalar automáticamente (Sí), abrir la release en el navegador (No) o cancelar?")

        # parent: intentamos usar root para que el diálogo sea modal sobre la app
        parent = globals().get("root", None)
        choice = _ask_update_custom(parent, "Actualización disponible", msg)

        if choice == "yes":
            # lanzar la descarga/instalador existente
            download_and_install_release_exe(repo, preferred_asset_name="Co-op_Stock_Manager.exe")
        elif choice == "no":
            if html_url:
                webbrowser.open(html_url)
            else:
                # si no hay url, mantenemos silencio mínimo
                try:
                    messagebox.showinfo("Actualizaciones", "No hay URL de la release para abrir.")
                except Exception:
                    pass
        else:
            # "cancel" (o cierre con X) -> no hacemos nada
            return

    except Exception:
        # Silencioso en startup: no queremos romper el arranque por un fallo en la comprobación
        return


# --- Llamada: poner esto justo después de cargar_datos() y antes de root.mainloop() ---
# programamos la comprobación una sola vez (no bloqueante)
try:
    root.after(1500, check_updates_on_startup)
except Exception:
    # si por alguna razón root no existe en este punto, no fallamos
    pass


# Atajos de teclado
root.bind_all("<Control-c>", copiar_seleccion, add="+")
root.bind_all("<Control-C>", copiar_seleccion, add="+")
root.bind_all("<Control-v>", pegar_seleccion, add="+")
root.bind_all("<Control-V>", pegar_seleccion, add="+")
root.bind("<Control-z>", deshacer)
root.bind("<Control-y>", rehacer)
root.bind_all("<Button-1>", clear_all_selection, add="+")
# Tema inicial y carga de datos
if current_theme == "dark":
    apply_dark()
else:
    apply_light()

cargar_datos()
root.mainloop()