# MarginRail v2

Web app batch per Commercial Margin Governance, riorganizzata per uso interno / pilot.

## Cosa fa
Il cliente:
1. scarica il template Excel
2. compila i fogli richiesti
3. carica il file nella web app
4. clicca **Esegui analisi**
5. ottiene dashboard, casi prioritari e file scaricabili
6. trova anche una copia persistente della run in `runs/`

## Struttura progetto
- `app.py` → web app Streamlit
- `engine/main_engine.py` → motore analitico
- `config/rules_config.json` → soglie e parametri
- `templates/MarginRail_Input_Template_v1.xlsx` → template ufficiale input
- `runs/` → storico locale di input, output, metadata, log e indice run
- `docs/` → note operative V2

## Avvio locale
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Requisiti input
La V2 accetta file Excel **standardizzati** con questi fogli obbligatori:
- `Vendite_2025`
- `Workflow_Deroghe_2025`
- `Clienti`
- `Prodotti`
- `Accordi_Commerciali`
- `Promo_2025`

## Dove vengono salvate le run
Ogni esecuzione riuscita o fallita viene salvata in:
```text
runs/run_YYYYMMDD_HHMMSS_xxxxxx/
```

Ogni run contiene:
- `input/`
- `output/`
- `metadata.json`
- `stdout.log`
- `stderr.log`
- zip output della run

Nella cartella `runs/` vengono inoltre mantenuti:
- `_runs_index.json` → indice locale sintetico delle run
- `_retention_policy.json` → policy retention locale, disattivata di default

## Cosa salva metadata.json
Per ogni run vengono salvati anche:
- `started_at`
- `finished_at`
- `duration_seconds`
- hash SHA256 del file input
- dimensione file input
- righe rilevate / analizzate se disponibili
- conteggio file output
- manifest sintetico degli artefatti generati
- return code esecuzione

## Retention policy
Non usa database e non usa sistemi esterni.

La retention è solo filesystem locale e parte **disattivata**.  
Quando `enabled=true` nella policy, la modalità attuale elimina le run più vecchie oltre `keep_last_n`.

## Limiti noti
- legge solo Excel conformi al template
- non è ancora multiutente
- non ha ancora login, ruoli o servizi esterni
