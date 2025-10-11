# BariSircle

Tool per mappare dati da file XML a fogli Excel.

## Cosa fa

Prende file XML (tipo quelli medici/radiologici) e copia automaticamente i dati in un foglio Excel esistente, trovando la riga giusta per ogni paziente.

## Come usare

1. Apri l'app
2. Carica un XML di esempio per vedere la struttura
3. Clicca sugli elementi che ti servono e mappali alle colonne Excel
4. Salva la configurazione
5. Vai in "Elaborazione Batch" e scegli la cartella con tutti gli XML
6. Parti - elabora tutto in automatico

## Requisiti

- Python 3.x
- openpyxl

```bash
pip install -r requirements.txt
```

## Esecuzione

```bash
python xml_to_excel_mapper_gui.py
```

Oppure usa l'exe se l'hai compilato.

## Note

- I nomi pazienti vengono estratti dal nome del file (formato: `COGNOME_NOME_data_...xml`)
- Controlla sempre i match parziali/ambigui nel log
- Fai backup del file Excel prima di elaborare

