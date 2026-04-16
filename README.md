# AutomationCalendarImporter

# Importatore da Excel a Google Calendar

**Riepilogo**
* Importa in modo efficiente eventi da un file Excel a un Google Calendar.
* Guida l'utente attraverso un'intuitiva interfaccia grafica (GUI) in 6 passaggi.
* Controlla automaticamente i duplicati e supporta ricorrenze e promemoria automatici.
* Necessario Python

---

## 🛠️ Configurazione Iniziale (Step-by-Step)

Per far funzionare lo script, è necessario configurare l'accesso alle API di Google. Segui questi passaggi dettagliati.

### 1. Come ottenere le credenziali (`credenziali.json`)
Il file `credenziali.json` permette allo script di "parlare" con Google come se fosse un utente autorizzato (Service Account).

1.  Vai sulla [Google Cloud Console](https://console.cloud.google.com/).
2.  **Crea un Progetto**: Clicca in alto a sinistra su "Seleziona un progetto" > "Nuovo progetto" e dagli un nome (es. "Excel-Calendar-Importer").
3.  **Abilita le API**: Nel menu a sinistra, vai su **API e servizi > Dashboard**. Clicca su **"+ ABILITA API E SERVIZI"**, cerca "Google Calendar API" e clicca su **Abilita**.
4.  **Crea un Service Account**:
    * Vai su **API e servizi > Credenziali**.
    * Clicca su **"CREATE CREDENTIALS"** in alto e scegli **"Service Account"**.
    * Inserisci un nome (es. "calendar-bot") e clicca su "Crea e continua". Salta i passaggi opzionali e clicca su **Fine**.
5.  **Scarica la Chiave JSON**:
    * Nella lista delle credenziali, clicca sull'indirizzo email del Service Account appena creato.
    * Vai nella scheda **"Keys" (Chiavi)**.
    * Clicca su **"Add Key" > "Create new key"**.
    * Scegli il formato **JSON** e clicca su Crea.
    * Verrà scaricato un file sul tuo PC. **Rinominalo in `credenziali.json`** e spostalo nella stessa cartella dello script Python.

### 2. Come ottenere l'ID Calendario
Lo script deve sapere in quale calendario scrivere.

1.  Apri [Google Calendar](https://calendar.google.com/) dal tuo browser.
2.  Sulla sinistra, sotto "I miei calendari", passa il mouse sul calendario che vuoi usare.
3.  Clicca sui tre puntini verticali e seleziona **"Impostazioni e condivisione"**.
4.  Scorri verso il basso fino alla sezione **"Integra calendario"**.
5.  Copia il testo sotto **"ID calendario"** (sarà qualcosa di simile a `nome@group.calendar.google.com` o la tua email principale).

### 3. Autorizzare lo script (Condivisione e Permessi)
Senza questo passaggio, Google bloccherà l'accesso allo script anche se le chiavi sono corrette.

1.  Copia l'indirizzo email del **Service Account** che hai creato nel punto 1 (lo trovi in `credenziali.json` alla voce `client_email` o nella console Google). Ha un formato tipo: `calendar-bot@progetto.iam.gserviceaccount.com`.
2.  Torna nelle **"Impostazioni e condivisione"** del tuo calendario.
3.  Trova la sezione **"Condividi con persone o gruppi specifici"**.
4.  Clicca su **"+ Aggiungi persone e gruppi"**.
5.  Incolla l'email del Service Account.
6.  **Fondamentale**: Nel menu a discesa dei permessi, seleziona **"Apportare modifiche agli eventi"**.
7.  Clicca su **Invia**.

---

## 🚀 Installazione e Utilizzo

**Installa Python**
- Vai sul sito ufficiale Python e installa il [Python install manager]([https://calendar.google.com/](https://www.python.org/downloads/)).

**Apri il prompt dei comandi (CMD)**
- Esegui questo comando all'interno del CMD per installare le librerie necessarie:
```bash
pip3 install pandas openpyxl google-api-python-client google-auth
```

**Avvio dello Script**
- Assicurati di avere `prova.py` e `credenziali.json` nella stessa cartella:
```bash
python3 prova.py
```

---

## 📈 Log e Output
Lo strumento restituisce i dati visivamente tramite una finestra di log:
* ✅ **Creato**: L'evento è stato caricato con successo.
* ⏭️ **Già esiste**: L'evento è stato ignorato per evitare duplicati.
* ❌ **Errore**: Problema di formattazione o di rete.

**Note sulla Licenza:** GNU GPL v3.0. Realizzato per scopi di produttività. Non condividere mai pubblicamente il file `credenziali.json`.
