# Fatture Grimaldi App — istruzioni per Claude

Web app per controllare ogni mese la fattura del fornitore **Grimaldi** (traghetti Spagna ↔ Italia)
di SMET. Andrà integrata nel **portale SMET**.

**Prima di fare qualsiasi cosa leggi `docs/HANDOFF.md`**: contiene il processo, l'analisi già fatta
sui dati di settembre 2026, le decisioni prese e il piano di lavoro.

## Regole

- **Non leggere il codice del collega.** Gli script `embarques_esp`, `embarques_it` e `factura`
  (Python, con il loro Excel "SQL") non vanno aperti né cercati: il proprietario vuole una soluzione
  non influenzata. Si confrontano solo gli **output** (l'Excel "SQL"), e solo quando lui lo chiede.
- **Non mettere mai nel repository** fatture, export Intermodale/Showtrip, listini, credenziali,
  password o dati di clienti. Anche se il repository è privato. I dati stanno in una cartella
  esterna al repo (vedi `docs/HANDOFF.md`). `.gitignore` blocca già `*.xls`, `*.xlsx`, `.env`, `data/`.
- Sul Mac ci sono **altre applicazioni e sessioni Claude** che usano già Intermodale, Showtrip,
  Gmail e il portale SMET: riusa il loro modo di accedere (vedi `docs/HANDOFF.md` §6). Non
  modificarle senza chiedere: altre sessioni ci stanno lavorando.
- Lingua: si parla con l'utente in **italiano**; i documenti sorgente sono in spagnolo;
  l'interfaccia dell'app sarà probabilmente in spagnolo (da confermare).
- Lavora sul branch `claude/peaceful-johnson-fskah7` finché l'utente non dice diversamente.

## Comandi

```bash
pip install -r requirements.txt
python prototype/match_prototype.py PERCORSO/FATTURA.xls PERCORSO/IMBARCHI.xlsx --out /tmp/risultato.xlsx
```
