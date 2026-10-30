# Passaggio di consegne — Fatture Grimaldi App

Documento per ripartire in una **sessione Claude Code in locale sul Mac**, dopo una prima sessione
nel cloud (ottobre 2026). La sessione cloud non poteva raggiungere le applicazioni installate sul Mac
(accesso a Intermodale, portale SMET, automazioni), quindi il lavoro continua in locale.

---

## 0. Come ripartire sul Mac

1. **Installa Claude Code** (se non c'è già): `curl -fsSL https://claude.ai/install.sh | bash`,
   oppure usa Claude Code dentro l'app desktop di Claude.
2. **Clona il repository e passa al branch di lavoro:**
   ```bash
   cd ~/Developer            # o la cartella dove tieni i progetti
   git clone https://github.com/Daniele027/Fatture-Grimaldi-App.git
   cd Fatture-Grimaldi-App
   git checkout claude/peaceful-johnson-fskah7
   ```
3. **Metti i dati fuori dal repository**, una cartella per mese, ad esempio:
   ```
   ~/SMET/Grimaldi-dati/2026-09/
       fattura_grimaldi.xls        (export Grimaldi così com'è)
       imbarchi_intermodale.xlsx   (fogli ESP e IT)
       tariffe.xlsx                (listino Grimaldi)
       sql_collega.xlsx            (output del collega, solo per il confronto finale)
   ```
4. **Avvia Claude dando accesso anche alle altre app** (sostituisci i percorsi con quelli veri):
   ```bash
   claude --add-dir ~/SMET/Grimaldi-dati --add-dir ~/PERCORSO/app-intermodale --add-dir ~/PERCORSO/portale-smet
   ```
   Le cartelle si possono aggiungere anche dopo, dentro la sessione, con `/add-dir PERCORSO`.
   **Non** aggiungere la cartella con gli script del collega.
5. **Primo messaggio da incollare:**
   > Leggi CLAUDE.md e docs/HANDOFF.md. Stiamo continuando il progetto Fatture Grimaldi iniziato in
   > una sessione cloud. I dati di settembre sono in ~/SMET/Grimaldi-dati/2026-09. Prima trova sul Mac
   > le applicazioni che accedono a Intermodale, Showtrip e al portale SMET (sezione 6), dimmi cosa
   > hai trovato e come ci si autentica, senza modificarle. Non aprire il codice del collega. Poi
   > riparti dalla fase 1 del piano.

---

## 1. Obiettivo

Ogni mese Grimaldi manda una fattura con tutti gli imbarchi di SMET (rimorchi) tra Spagna e Italia.

- **A — Imputazione costi:** collegare ogni riga della fattura alla nostra posizione (numero di
  viaggio) in **Showtrip**, per conoscere il costo reale di ogni viaggio.
- **B — Controllo tariffe:** confrontare il costo fatturato con la tariffa concordata e reclamare le
  differenze **solo quando sono a nostro favore** (fatturato > tariffato).

**Obiettivo dell'utente:** ogni mese **si importa solo la fattura Grimaldi**. Tutto il resto lo fa
l'app: scarica gli imbarchi da Intermodale, corregge e completa i dati, calcola le differenze e
mostra **solo le righe che richiedono un intervento umano**. Poi l'app sarà integrata nel portale SMET.

## 2. Processo attuale (manuale)

Fonte: documento interno "Especificación de procesos – Factura Grimaldi 2026" (in spagnolo).

1. **Preparazione:** dal portale **Intermodale**, scheda *imbarchi*, si scaricano gli imbarchi del mese
   ESP→IT e IT→ESP. Si correggono targhe sbagliate e si riempiono a mano targhe e commodity mancanti.
   Si aggiorna l'Excel delle tariffe. Nel file Grimaldi si toglie la prima riga vuota e lo si salva
   da `.xls` a `.xlsx`.
2. **Script:** tre script Python del collega (`embarques_esp` → `embarques_it` → `factura`), con
   percorsi modificati a mano ogni mese. Producono un Excel chiamato **"SQL"**.
3. **Revisione (la fase più lunga):** nell'Excel SQL mancano numeri di viaggio e prezzi, per targhe
   sbagliate o per date di imbarco in settimane diverse (gli script abbinano anche per settimana).
   Si correggono a mano.
4. Si sostituisce la data dentro **`ref1`** (chiave univoca nostra) con la **data di carico di
   Showtrip** (da un export mensile di Showtrip) e si aggiunge la colonna **`Ncrt`**, tutta a 0.
5. Si carica l'Excel nel **database SQL** tramite **Access**.
6. **Reclamo:** Excel per Grimaldi con le sole righe dove `Dif > 0`, senza dati interni (n° viaggio,
   `ref1`…).
7. **Chiusura:** se la differenza è riconosciuta, `Dif` → 0 e l'importo va in `Ncrt` (nota di
   credito). Se l'errore era nostro, `Dif` → 0 e l'importo si aggiunge a `coste`.
   **Quadratura:** `coste + dif + ncrt = Amount Per Unit`.

## 3. Fonti dati (struttura verificata su settembre 2026)

### Fattura Grimaldi (`.xls`, un solo foglio)
- La riga 1 è vuota, le intestazioni sono nella riga 2 → `pd.read_excel(path, header=1)`.
  Serve `xlrd` per leggere il `.xls`.
- **Una riga per rimorchio.** `Shipment No.` raggruppa più rimorchi (fino a 6 per spedizione).
- Colonne utili: `Vessel`, `Voyage`, `Laden on Board Date` (data/ora di imbarco), `POL`, `POD`
  (porti in chiaro: BARCELONA, VALENCIA, SAGUNTO, SAVONA, LIVORNO, SALERNO, CIVITAVECCHIA,
  CAGLIARI, PORTO TORRES), `Shipment No.`, `Status`, `Chassis/Plate/Cntr No.` (targa, sempre
  uguale a `VIN`), `Equipment Type` (Trailer-Full / Trailer-Empty), `Commodity`, `Length`
  (13.6, 10.5…), `Amount Per Unit` (nolo), `ETS Surcharge` (sovrapprezzo emissioni),
  `Charge Per Shipment` (raro, costo fisso per spedizione), `Invoice Number`, `Invoice Date`.
- **`Invoice Date` è scritta male:** "26/09/2007" significa **2026-09-07** (yy/mm/dd mostrato come
  dd/mm/yyyy). In un mese ci sono più fatture con date diverse.
- Colonne sempre vuote o costanti: `Payor remarks`, `Shipper Ref No`, `Height`=0,
  `Number Of Drivers`=0, `Driver Chg`=0, `Vat Code`=78, `Client`/`Payor` = SMET SA.

### Export Intermodale (`.xlsx`, fogli `ESP` e `IT`)
- `ESP` = imbarchi Spagna→Italia, `IT` = Italia→Spagna. Le stesse colonne in entrambi.
- **Targa del rimorchio = colonna `S/R`**, con suffissi da togliere: `XA693EP/SAT`, `/M`, `/HS`, `/PC`,
  `/V`, `/CS13`, `/ISO`… Due formati: `AA999AA` e `AA99999`. Le colonne `Truck`, `Driver`,
  `Shipment` sono vuote.
- `Prot.` (es. `SAVA/030431`): molto probabilmente è la **posizione Showtrip** → da confermare.
- `Linea Porti`: codici porto `POL POD` → SAG, SAL, BCN, SAV, VAL, LIV, CIV, CAG, PTS.
- `Dt.Imbarco` e `Dt.Sbarco` sono testo `dd/mm/yyyy  HH:MM` (due spazi). `Dt. Carico` e
  `Dt. Scarico` sono date vere.
- Altre colonne: `Tipo` (MEGA, AU, GL, GLL, V03, PIVO, M120…), `PesoN`, `Cliente`, `commodity`
  (nomi diversi da Grimaldi), `N° Ordine`, `Fil.` (filiale), `NAVE`, `ZG Part.`, `ZG. Arr.`,
  `Corr_ok`, `BlType`.
- Alcune righe **non hanno la targa** (96 a settembre).

### Mancano ancora
- **Listino tariffe** Grimaldi.
- **Export Showtrip** con la data di carico per posizione, oppure un accesso diretto ai dati Showtrip.
- **Schema della tabella SQL** di destinazione e definizione di `ref1`.

## 4. Risultati dell'analisi su settembre 2026

Fattura: 3.406 righe (1.859 spedizioni, 15 date fattura). Intermodale: 3.414 righe (ESP 1.704, IT 1.710).

Con il prototipo (`prototype/match_prototype.py`):

| Livello | Significato | Righe |
|---|---|---|
| `A_esatto` | stessa targa, stessi porti, imbarco entro ±24 ore | 3.238 (95%) |
| `B_data_spostata` | stessa targa e stessi porti, data diversa (fino a 7 giorni) | 63 |
| `C_un_porto_diverso` | stessa targa, POL o POD diverso (es. Cagliari vs Salerno) | 2 |
| `S_proposta_senza_targa` | riga Intermodale **senza targa**, stessa tratta, data vicina, merce compatibile | 96 |
| `X_non_abbinata` | nessun candidato | 7 |

- **32 righe abbinate hanno una commodity diversa** da quella di Intermodale (per esempio fatturate
  come pieno o come merce più cara, mentre in Intermodale risultano vuote): sono **possibili reclami**
  anche senza listino.
- 15 righe Intermodale non vengono usate (imbarchi senza fattura o fatturati in un altro mese?).
- **Importi anomali visibili anche senza listino:** un importo negativo, prezzi isolati molto diversi
  dal resto per la stessa tratta + tipo (es. rimorchi vuoti a un prezzo molto più alto del solito),
  una tratta con POL = POD. Sulla stessa tratta e merce ci sono anche più livelli di prezzo
  (sconti? promozioni?) da spiegare con il listino.
- L'`ETS Surcharge` dipende dalla tratta e sembra più basso per i rimorchi da 10,5 m.

## 5. Algoritmo di abbinamento (prototipo)

1. Targa normalizzata: maiuscole, solo lettere e numeri, si toglie tutto dopo `/`.
2. Candidati: stessa targa e `|Dt.Imbarco − Laden on Board Date| ≤ 7 giorni`.
3. Punteggio = giorni di distanza + 3 se il POL è diverso + 2 se il POD è diverso. Assegnazione 1:1
   greedy dal punteggio migliore. Possibile miglioramento: assegnazione ottima (algoritmo ungherese).
4. Le righe rimaste si abbinano alle righe Intermodale **senza targa** sulla stessa tratta, entro ±48 ore,
   con penalità se la merce non è compatibile (tabella `COMMODITY_OK`, ricavata dai dati e **da far
   validare** all'ufficio). Sono solo **proposte da confermare**.
5. Ogni riga ha un livello, che decide se va in revisione: A = automatica, B/C/S = da confermare,
   X = da gestire a mano.

Possibili miglioramenti: correzione di targhe con un carattere sbagliato (distanza di Levenshtein 1
con lo stesso imbarco), riconoscimento dei trasbordi e delle soste nei porti intermedi, uso di
`Vessel`/`NAVE` come conferma.

## 6. Altre applicazioni e sessioni sul Mac (da individuare all'inizio)

L'utente ha già sviluppato altre applicazioni, e altre **sessioni Claude** le stanno usando o
sviluppando. Alcune hanno **già l'accesso a Intermodale**. Cosa sappiamo:

- **Accesso a Intermodale:** esiste un'app/script sul Mac che si collega al portale Intermodale.
  Va trovata e capito **come si autentica e scarica** la scheda *imbarchi* (API, scraping con
  browser, export). L'obiettivo è riusare lo stesso meccanismo per scaricare ESP e IT del mese in
  automatico, senza duplicare le credenziali.
- **Plugin / server MCP `smet-automation`:** già configurato per l'utente. Scarica gli export di
  **Ingetime** (controllo orari delle 3 sedi: Barcelona, Valencia, Madrid), rigenera Excel e invia
  mail con Gmail; lo usano attività programmate settimanali e mensili. Probabilmente vive in una
  cartella sul Mac (cerca la configurazione MCP: `~/.claude.json`, `.mcp.json` nei progetti, oppure
  `~/Library/Application Support/Claude/claude_desktop_config.json`). Può essere un **modello** per
  esporre le funzioni della nuova app (es. `scarica_imbarchi_intermodale`).
- **Portale SMET:** la destinazione finale. Bisogna capire stack (framework, linguaggio), login,
  database e come si aggiunge un modulo, così la web app nasce già compatibile.
- **Showtrip:** il gestionale trasporti (posizioni = numeri di viaggio). Da capire se un'app esistente
  ci accede già (database, API o export).
- **Database SQL** in cui oggi si carica l'Excel con Access: da trovare stringa di connessione e schema.

Regole: queste app **si leggono ma non si modificano** senza chiedere, perché altre sessioni ci
lavorano. Le credenziali restano dove sono già (keychain, `.env` delle altre app) oppure in un `.env`
locale ignorato da git: **mai nel repository**.

## 7. Piano di lavoro

1. **Fase 1 — Motore di calcolo** (Python, con test):
   lettura diretta del `.xls` Grimaldi; download o lettura degli imbarchi Intermodale; abbinamento
   con livelli di affidabilità; elenco eccezioni. Validazione su settembre 2026 e **confronto solo
   degli output** con l'Excel "SQL" del collega.
2. **Fase 2 — Tariffe:** tabella delle tariffe con validità nel tempo (aumenti, sconti, prezzi
   speciali); calcolo di `Dif`; controllo commodity e vuoto/pieno; generazione dell'Excel reclamo
   per Grimaldi (solo `Dif > 0`, senza dati interni).
3. **Fase 3 — Web app:** si carica la fattura → download Intermodale automatico → schermata delle
   sole eccezioni, con conferma in un clic delle proposte → ciclo dei reclami
   (aperto / riconosciuto → `Ncrt` / respinto / errore nostro → `coste`) con controllo della
   quadratura → scrittura sul database SQL al posto di Access.
4. **Fase 4 — Integrazione nel portale SMET** con lo stesso stack e lo stesso login.

## 8. Domande aperte per l'utente

1. Percorsi sul Mac delle app che accedono a Intermodale e del portale SMET.
2. Listino tariffe attuale: da cosa dipende il prezzo (tratta, pieno/vuoto, commodity, lunghezza,
   `Tipo` mezzo, periodo)? Gli sconti come funzionano?
3. `Prot.` è il numero di posizione Showtrip? Come si ottiene la data di carico? Come si costruisce
   `ref1`?
4. Database: tipo (SQL Server?), tabella e colonne di destinazione.
5. L'`ETS Surcharge` va controllato rispetto al listino o solo il nolo?
6. Lingua dell'interfaccia (spagnolo?) e chi userà l'app (una persona o più persone/sedi).
7. Tabella `COMMODITY_OK` (merce Grimaldi ↔ merce Intermodale): è corretta?

## 9. Stato del repository

- `CLAUDE.md` — regole per Claude (lette automaticamente).
- `docs/HANDOFF.md` — questo documento.
- `prototype/match_prototype.py` — prototipo dell'abbinamento (non definitivo).
- `requirements.txt`, `.gitignore` (blocca Excel, CSV, Access, `.env`, `data/`).
- Nessun dato aziendale nel repository.
