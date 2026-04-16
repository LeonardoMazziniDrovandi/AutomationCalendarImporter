import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, ttk
import sys
import os

# ─────────────────────────────────────────────
#  FINESTRA ROOT NASCOSTA (serve per i dialog)
# ─────────────────────────────────────────────
root = tk.Tk()
root.withdraw()

# ─────────────────────────────────────────────
#  HELPER: dialog con radio button
# ─────────────────────────────────────────────
def chiedi_scelta(titolo, domanda, opzioni):
    """Mostra una finestra con radio button e restituisce la scelta."""
    win = tk.Toplevel(root)
    win.title(titolo)
    win.resizable(False, False)
    win.grab_set()

    tk.Label(win, text=domanda, font=("Segoe UI", 10), wraplength=380,
             justify="left", pady=10, padx=20).pack(anchor="w")

    scelta = tk.StringVar(value=opzioni[0][1])
    for label, valore in opzioni:
        tk.Radiobutton(win, text=label, variable=scelta, value=valore,
                       font=("Segoe UI", 10), padx=30).pack(anchor="w")

    risultato = [None]

    def conferma():
        risultato[0] = scelta.get()
        win.destroy()

    def annulla():
        risultato[0] = None
        win.destroy()

    tk.Button(win, text="Conferma", command=conferma,
              font=("Segoe UI", 10), width=12, pady=4).pack(pady=12)
    win.protocol("WM_DELETE_WINDOW", annulla)

    win.update_idletasks()
    w = win.winfo_reqwidth()
    h = win.winfo_reqheight()
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f"+{x}+{y}")

    win.wait_window()
    return risultato[0]


def chiedi_testo(titolo, domanda, default=""):
    """Dialog semplice per testo libero."""
    risposta = simpledialog.askstring(titolo, domanda, initialvalue=default, parent=root)
    return risposta


def chiedi_numero_colonna(titolo, domanda):
    """Chiede un numero colonna (1-based) e lo restituisce come indice 0-based."""
    while True:
        val = simpledialog.askstring(titolo, domanda, parent=root)
        if val is None:
            return None
        try:
            n = int(val.strip())
            if n >= 1:
                return n - 1   # converti in 0-based
        except ValueError:
            pass
        messagebox.showwarning("Valore non valido",
                               "Inserisci un numero intero maggiore o uguale a 1.", parent=root)


def chiedi_si_no(titolo, domanda):
    return messagebox.askyesno(titolo, domanda, parent=root)


# ═══════════════════════════════════════════════
#  STEP 0 — AVVISO INIZIALE
# ═══════════════════════════════════════════════
messagebox.showinfo(
    "⚠️  Prima di iniziare",
    "Assicurati che:\n\n"
    "• Il file Excel contenga UN SOLO FOGLIO con i dati.\n\n"
    "• Il file  credenziali.json  si trovi nella STESSA CARTELLA\n"
    "  di questo script Python.",
    parent=root
)

# ═══════════════════════════════════════════════
#  STEP 1 — SELEZIONA FILE EXCEL
# ═══════════════════════════════════════════════
messagebox.showinfo("Passo 1 / 6", "Seleziona il file Excel con le scadenze.", parent=root)
EXCEL_FILE = filedialog.askopenfilename(
    title="Seleziona il file Excel",
    filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")],
    parent=root
)
if not EXCEL_FILE:
    messagebox.showerror("Annullato", "Nessun file selezionato. Lo script verrà chiuso.")
    sys.exit()

# Verifica che credenziali.json esista vicino allo script
script_dir = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(script_dir, "credenziali.json")
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    messagebox.showerror(
        "File mancante",
        f"Non trovo  credenziali.json  in:\n{script_dir}\n\n"
        "Copia il file lì e riavvia lo script."
    )
    sys.exit()

# ═══════════════════════════════════════════════
#  STEP 1b — PRIMA RIGA: INTESTAZIONI O DATI?
# ═══════════════════════════════════════════════
try:
    import pandas as pd
    _df_preview = pd.read_excel(EXCEL_FILE, header=None, nrows=1)
except ImportError:
    messagebox.showerror("Modulo mancante",
                         "Il modulo 'pandas' non è installato.\n"
                         "Esegui:  pip install pandas openpyxl")
    sys.exit()
except Exception as e:
    messagebox.showerror("Errore lettura Excel", str(e))
    sys.exit()

# Mostra il contenuto della prima riga all'utente
prima_riga_vals = [str(v) for v in _df_preview.iloc[0].tolist()]
anteprima = "  ".join(
    f"[{i+1}] {v}" for i, v in enumerate(prima_riga_vals)
)
USA_HEADER = chiedi_si_no(
    "Prima riga del foglio",
    f"La prima riga del file Excel contiene i NOMI delle colonne (intestazioni),\n"
    f"oppure è già un dato da importare?\n\n"
    f"Contenuto attuale della riga 1:\n{anteprima}\n\n"
    f"✅ SÌ  →  è un'intestazione (verrà ignorata)\n"
    f"❌ NO  →  è un dato reale (verrà importata)"
)

# ═══════════════════════════════════════════════
#  STEP 2 — CALENDAR ID (con salvataggio opzionale)
# ═══════════════════════════════════════════════
CONFIG_FILE = os.path.join(script_dir, "calendario_config.json")
import json

def leggi_id_salvato():
    """Restituisce il Calendar ID salvato, o None se il file non esiste/è corrotto."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            dati = json.load(f)
            return dati.get("calendar_id", "").strip() or None
    except Exception:
        return None

def salva_id(calendar_id):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"calendar_id": calendar_id}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("Salvataggio fallito",
                               f"Non è stato possibile salvare l'ID:\n{e}", parent=root)

id_salvato = leggi_id_salvato()

if id_salvato:
    usa_salvato = chiedi_si_no(
        "Passo 2 / 6 — Calendar ID",
        f"È stato trovato un Calendar ID salvato:\n\n"
        f"  {id_salvato}\n\n"
        f"Vuoi usarlo?\n\n"
        f"✅ SÌ  →  usa questo ID\n"
        f"❌ NO  →  inseriscine uno nuovo"
    )
    if usa_salvato:
        CALENDAR_ID = id_salvato
    else:
        CALENDAR_ID = chiedi_testo(
            "Passo 2 / 6 — Calendar ID",
            "Incolla il nuovo ID del Google Calendar\n(es. xxxx@group.calendar.google.com):"
        )
        if not CALENDAR_ID or not CALENDAR_ID.strip():
            messagebox.showerror("Annullato", "Nessun Calendar ID inserito. Lo script verrà chiuso.")
            sys.exit()
        CALENDAR_ID = CALENDAR_ID.strip()
        if chiedi_si_no("Salva Calendar ID",
                         f"Vuoi salvare questo ID per le prossime volte?\n\n  {CALENDAR_ID}"):
            salva_id(CALENDAR_ID)
            messagebox.showinfo("Salvato",
                                f"ID salvato in:\n{CONFIG_FILE}", parent=root)
else:
    # Nessun file salvato: chiedi e offri di salvare
    CALENDAR_ID = chiedi_testo(
        "Passo 2 / 6 — Calendar ID",
        "Incolla l'ID del Google Calendar\n(es. xxxx@group.calendar.google.com):"
    )
    if not CALENDAR_ID or not CALENDAR_ID.strip():
        messagebox.showerror("Annullato", "Nessun Calendar ID inserito. Lo script verrà chiuso.")
        sys.exit()
    CALENDAR_ID = CALENDAR_ID.strip()
    if chiedi_si_no("Salva Calendar ID",
                     f"Vuoi salvare questo ID così la prossima volta non dovrai reinserirlo?\n\n"
                     f"  {CALENDAR_ID}"):
        salva_id(CALENDAR_ID)
        messagebox.showinfo("Salvato",
                            f"ID salvato in:\n{CONFIG_FILE}", parent=root)

# ─── Carica il file Excel tenendo conto della scelta sulla prima riga ───
try:
    if USA_HEADER:
        df = pd.read_excel(EXCEL_FILE, header=0)
    else:
        df = pd.read_excel(EXCEL_FILE, header=None)
except Exception as e:
    messagebox.showerror("Errore lettura Excel", str(e))
    sys.exit()

num_colonne = len(df.columns)
if USA_HEADER:
    # Mostra i nomi reali delle colonne
    elenco_colonne = "\n".join(
        f"  {i+1} → {col}" for i, col in enumerate(df.columns)
    )
else:
    # Senza intestazioni: mostra numero + valore della prima cella come anteprima
    elenco_colonne = "\n".join(
        f"  {i+1} → (es. {str(df.iloc[0, i])[:30]})" for i in range(num_colonne)
    )
info_colonne = f"Il foglio ha {num_colonne} colonne:\n\n{elenco_colonne}"

# ═══════════════════════════════════════════════
#  STEP 3 — COLONNA TITOLO EVENTO
# ═══════════════════════════════════════════════
messagebox.showinfo("Passo 3 / 6 — Colonne disponibili", info_colonne, parent=root)
col_titolo_idx = chiedi_numero_colonna(
    "Passo 3 / 6 — Titolo evento",
    f"Inserisci il NUMERO della colonna da usare come titolo dell'evento\n"
    f"(1 = prima colonna, {num_colonne} = ultima):"
)
if col_titolo_idx is None:
    messagebox.showerror("Annullato", "Nessuna colonna selezionata.")
    sys.exit()

# ═══════════════════════════════════════════════
#  STEP 4 — COLONNA DATA
# ═══════════════════════════════════════════════
col_data_idx = chiedi_numero_colonna(
    "Passo 4 / 6 — Data evento",
    f"Inserisci il NUMERO della colonna con la DATA dell'evento\n"
    f"(1 = prima colonna, {num_colonne} = ultima):"
)
if col_data_idx is None:
    messagebox.showerror("Annullato", "Nessuna colonna selezionata.")
    sys.exit()

# ═══════════════════════════════════════════════
#  STEP 5 — RICORRENZA
# ═══════════════════════════════════════════════
RICORRENZE = [
    ("1 settimana",  "RRULE:FREQ=WEEKLY"),
    ("1 mese",       "RRULE:FREQ=MONTHLY"),
    ("6 mesi",       "RRULE:FREQ=MONTHLY;INTERVAL=6"),
    ("1 anno",       "RRULE:FREQ=YEARLY"),
    ("2 anni",       "RRULE:FREQ=YEARLY;INTERVAL=2"),
    ("3 anni",       "RRULE:FREQ=YEARLY;INTERVAL=3"),
    ("Nessuna (evento singolo)", "NESSUNA"),
]
rrule = chiedi_scelta(
    "Passo 5 / 6 — Ricorrenza",
    "Con quale frequenza deve ripetersi l'evento?",
    RICORRENZE
)
if rrule is None:
    messagebox.showerror("Annullato", "Nessuna ricorrenza selezionata.")
    sys.exit()

# ═══════════════════════════════════════════════
#  STEP 6 — DESCRIZIONE
# ═══════════════════════════════════════════════
descrizione_fissa = ""
col_descrizione_idx = None

vuole_descrizione = chiedi_si_no(
    "Passo 6 / 6 — Descrizione",
    "Vuoi aggiungere una descrizione agli eventi?"
)
if vuole_descrizione:
    da_colonna = chiedi_si_no(
        "Fonte descrizione",
        "La descrizione deve essere presa da una colonna del file Excel?\n\n"
        "(Scegli NO per inserirla manualmente)"
    )
    if da_colonna:
        col_descrizione_idx = chiedi_numero_colonna(
            "Colonna descrizione",
            f"Inserisci il NUMERO della colonna con la descrizione\n"
            f"(1 = prima colonna, {num_colonne} = ultima):"
        )
        if col_descrizione_idx is None:
            messagebox.showwarning("Attenzione",
                                   "Nessuna colonna selezionata: la descrizione sarà vuota.")
    else:
        descrizione_fissa = chiedi_testo(
            "Descrizione manuale",
            "Inserisci il testo della descrizione (uguale per tutti gli eventi):"
        ) or ""

# ═══════════════════════════════════════════════
#  RIEPILOGO PRIMA DI PROCEDERE
# ═══════════════════════════════════════════════
col_titolo_nome = df.columns[col_titolo_idx] if USA_HEADER else f"colonna {col_titolo_idx+1}"
col_data_nome   = df.columns[col_data_idx]   if USA_HEADER else f"colonna {col_data_idx+1}"
col_desc_nome   = (
    (df.columns[col_descrizione_idx] if USA_HEADER else f"colonna {col_descrizione_idx+1}")
    if col_descrizione_idx is not None
    else ("Manuale" if descrizione_fissa else "—")
)

rrule_label = next(lab for lab, val in RICORRENZE if val == rrule)

riepilogo = (
    f"📄  File Excel:      {os.path.basename(EXCEL_FILE)}\n"
    f"📋  Prima riga:      {'Intestazioni (ignorata)' if USA_HEADER else 'Dato reale (importata)'}\n"
    f"📅  Calendar ID:     {CALENDAR_ID}\n"
    f"🏷️   Titolo evento:   colonna {col_titolo_idx+1} ({col_titolo_nome})\n"
    f"📆  Data evento:     colonna {col_data_idx+1} ({col_data_nome})\n"
    f"🔁  Ricorrenza:      {rrule_label}\n"
    f"📝  Descrizione:     {col_desc_nome}\n\n"
    f"Vuoi procedere con l'importazione?"
)
if not chiedi_si_no("Riepilogo — Conferma", riepilogo):
    messagebox.showinfo("Annullato", "Importazione annullata dall'utente.")
    sys.exit()

# ═══════════════════════════════════════════════
#  IMPORTAZIONE
# ═══════════════════════════════════════════════
try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
except ImportError:
    messagebox.showerror("Modulo mancante",
                         "Il modulo 'google-api-python-client' non è installato.\n"
                         "Esegui:  pip install google-api-python-client google-auth")
    sys.exit()

SCOPES = ['https://www.googleapis.com/auth/calendar']
try:
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('calendar', 'v3', credentials=creds)
except Exception as e:
    messagebox.showerror("Errore autenticazione",
                         f"Impossibile autenticarsi con Google:\n{e}")
    sys.exit()


def evento_esiste(summary, data_str):
    start_day = data_str + "T00:00:00Z"
    end_day   = data_str + "T23:59:59Z"
    eventi = service.events().list(
        calendarId=CALENDAR_ID,
        timeMin=start_day,
        timeMax=end_day,
        q=summary,
        singleEvents=True
    ).execute()
    return len(eventi.get('items', [])) > 0


# ── Finestra di log ────────────────────────────
log_win = tk.Toplevel(root)
log_win.title("📋  Log importazione")
log_win.geometry("600x420")

frame = tk.Frame(log_win)
frame.pack(fill="both", expand=True, padx=10, pady=10)

scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side="right", fill="y")

log_box = tk.Text(frame, yscrollcommand=scrollbar.set,
                  font=("Consolas", 9), state="disabled", wrap="word")
log_box.pack(fill="both", expand=True)
scrollbar.config(command=log_box.yview)


def log(msg):
    log_box.config(state="normal")
    log_box.insert("end", msg + "\n")
    log_box.see("end")
    log_box.config(state="disabled")
    log_win.update()


creati = 0
saltati = 0
errori = 0

for i, riga in df.iterrows():
    try:
        titolo_val = str(riga.iloc[col_titolo_idx]).strip()
        data_excel = riga.iloc[col_data_idx]

        if titolo_val == "nan" or pd.isna(data_excel):
            continue

        data_evento = pd.to_datetime(data_excel, dayfirst=True)
        data_str = data_evento.strftime('%Y-%m-%d')

        summary = titolo_val

        if evento_esiste(summary, data_str):
            log(f"⏭️  Già esiste: {summary} ({data_str})")
            saltati += 1
            continue

        # Descrizione
        if col_descrizione_idx is not None:
            desc = str(riga.iloc[col_descrizione_idx]).strip()
            desc = "" if desc == "nan" else desc
        else:
            desc = descrizione_fissa

        evento = {
            'summary': summary,
            'description': desc,
            'start': {'date': data_str},
            'end':   {'date': data_str},
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'popup', 'minutes': 20160}  # 14 giorni prima
                ]
            }
        }
        if rrule != "NESSUNA":
            evento['recurrence'] = [rrule]

        service.events().insert(calendarId=CALENDAR_ID, body=evento).execute()
        log(f"✅  Creato: {summary} ({data_str})")
        creati += 1

    except Exception as e:
        log(f"❌  Errore riga {i}: {e}")
        errori += 1

log(f"\n{'─'*50}")
log(f"🎉  IMPORTAZIONE COMPLETATA")
log(f"   ✅  Creati:   {creati}")
log(f"   ⏭️   Saltati:  {saltati}")
log(f"   ❌  Errori:   {errori}")

tk.Button(log_win, text="Chiudi", command=root.destroy,
          font=("Segoe UI", 10), width=12, pady=4).pack(pady=8)

root.mainloop()
