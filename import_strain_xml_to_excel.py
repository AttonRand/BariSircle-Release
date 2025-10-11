#!/usr/bin/env python3
"""
Script per importare dati strain da file XML in Excel.

Il script legge file XML con pattern COGNOME_NOME_* dalla cartella specificata,
estrae i valori di strain usando le mappature configurate nel file JSON,
e li inserisce nel file Excel nelle colonne appropriate.

Le mappature XPath → Colonna Excel sono configurabili tramite la GUI
xml_to_excel_mapper_gui.py e salvate in file .json
"""

import os
import sys
import xml.etree.ElementTree as ET
import pandas as pd
import re
import json
from pathlib import Path
import logging
from typing import Dict, Optional, Tuple, List


def setup_logging():
    """Configura il logging per il debug"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('strain_import.log'),
            logging.StreamHandler()
        ]
    )


def parse_filename(filename: str) -> Optional[Tuple[str, str]]:
    """
    Estrae cognome e nome dal nome del file XML.
    
    Pattern: COGNOME_NOME_DATA_etc.xml
    
    Args:
        filename: Nome del file XML
        
    Returns:
        Tupla (cognome, nome) o None se il pattern non corrisponde
    """
    # Rimuovi estensione .xml
    basename = filename.replace('.xml', '')
    
    # Split per underscore
    parts = basename.split('_')
    if len(parts) < 2:
        return None

    # Gestisci cognomi con particelle (DE, DEL, DEGL, D', etc.)
    cognome = parts[0].strip()
    nome = parts[1].strip()
    
    # Se il cognome inizia con particelle comuni, includi la parte successiva
    if cognome.upper() in ['DE', 'DEL', 'DELL', 'DEGL', 'DEGLI', 'DI', 'DA', 'DAL', 'DALLA', 'LE', 'LA', 'LO']:
        if len(parts) >= 3:
            cognome = f"{cognome} {parts[1]}"
            nome = parts[2].strip()
        else:
            return None
    
    # Gestisci apostrofi sostituiti con underscore (es. D'ARCHI diventa D_ARCHI)
    if cognome.endswith('D') and len(parts) >= 3 and parts[1].startswith('ARCHI'):
        cognome = f"D'{parts[1]}"
        nome = parts[2].strip() if len(parts) >= 3 else nome
    
    return (cognome, nome)


def load_mapping_config(config_file: str) -> List[Dict]:
    """
    Carica la configurazione delle mappature dal file JSON.

    Args:
        config_file: Percorso del file JSON con le configurazioni

    Returns:
        Lista di dizionari con le mappature
    """
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('mappings', [])
    except FileNotFoundError:
        logging.error(f"File di configurazione non trovato: {config_file}")
        raise
    except json.JSONDecodeError as e:
        logging.error(f"Errore nel parsing del file JSON: {e}")
        raise


def excel_column_to_index(col: str) -> int:
    """
    Converte una colonna Excel (es. 'AB', 'AC') in indice numerico (0-based).

    Args:
        col: Nome colonna Excel (es. 'AB')

    Returns:
        Indice numerico (es. 'AB' → 27)
    """
    result = 0
    col = col.upper()
    for char in col:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 0-based


def extract_strain_values(xml_file: str, mappings: List[Dict]) -> Dict[str, Optional[float]]:
    """
    Estrae i valori strain dal file XML usando le mappature configurate.

    Args:
        xml_file: Percorso del file XML
        mappings: Lista di mappature XPath → Colonna Excel

    Returns:
        Dizionario con i valori strain estratti {mapping_name: value}
    """
    strain_data = {}

    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Per ogni mapping configurato
        for mapping in mappings:
            name = mapping.get('name')
            xpath = mapping.get('xpath')
            attribute = mapping.get('attribute')
            data_type = mapping.get('type', 'text')

            # Cerca l'elemento usando l'XPath configurato
            elements = root.findall(f".//{xpath}")

            if elements:
                # Prendi il primo elemento trovato
                element = elements[0]

                # Estrai il valore dall'attributo specificato
                value = element.get(attribute)

                if value:
                    # Converti il valore nel tipo appropriato
                    if data_type == 'number':
                        try:
                            strain_data[name] = float(value)
                        except ValueError:
                            logging.warning(f"Impossibile convertire '{value}' in numero per {name}")
                            strain_data[name] = None
                    else:
                        strain_data[name] = value
                else:
                    strain_data[name] = None
            else:
                logging.debug(f"Nessun elemento trovato per XPath: {xpath}")
                strain_data[name] = None

    except Exception as e:
        logging.error(f"Errore nel parsing di {xml_file}: {e}")

    return strain_data


def find_patient_row(df: pd.DataFrame, cognome: str, nome: str) -> Optional[int]:
    """
    Trova la riga del paziente nel DataFrame Excel.
    
    Args:
        df: DataFrame Excel
        cognome: Cognome del paziente
        nome: Nome del paziente
        
    Returns:
        Indice della riga o None se non trovato
    """
    # Normalizza cognome e nome per il confronto
    def normalize_name(name):
        if pd.isna(name):
            return ""
        return str(name).upper().strip()
    
    cognome_norm = normalize_name(cognome)
    nome_norm = normalize_name(nome)
    
    # Cerca per match esatto (case-insensitive)
    mask = (df['Cognome'].apply(normalize_name) == cognome_norm) & \
           (df['Nome'].apply(normalize_name) == nome_norm)
    matching_rows = df.index[mask].tolist()
    
    if matching_rows:
        return matching_rows[0]
    
    # Se non trovato match esatto, prova varianti del cognome
    cognome_variations = [cognome_norm]
    
    # Aggiungi varianti per cognomi con particelle
    if ' ' in cognome_norm:
        # Prova senza spazi (es. "DE LUCA" -> "DELUCA")
        cognome_variations.append(cognome_norm.replace(' ', ''))
        # Prova solo la parte finale (es. "DE LUCA" -> "LUCA")
        parts = cognome_norm.split()
        if len(parts) > 1:
            cognome_variations.append(parts[-1])
    else:
        # Prova ad aggiungere particelle comuni
        for particella in ['DE ', 'DEL ', 'DELL\'', 'DEGL\'']:
            cognome_variations.append(particella + cognome_norm)
    
    # Prova le variazioni
    for var_cognome in cognome_variations:
        mask = (df['Cognome'].apply(normalize_name) == var_cognome) & \
               (df['Nome'].apply(normalize_name) == nome_norm)
        matching_rows = df.index[mask].tolist()
        if matching_rows:
            logging.info(f"Match trovato con variante {var_cognome} per {cognome} {nome}")
            return matching_rows[0]
    
    # Se non trovato match esatto, cerca per match parziale
    mask_partial = (df['Cognome'].apply(normalize_name).str.contains(cognome_norm, na=False)) & \
                   (df['Nome'].apply(normalize_name).str.contains(nome_norm, na=False))
    matching_rows_partial = df.index[mask_partial].tolist()
    
    if matching_rows_partial:
        logging.warning(f"Match parziale trovato per {cognome} {nome} alla riga {matching_rows_partial[0]}")
        return matching_rows_partial[0]
    
    return None


def update_excel_data(df: pd.DataFrame, row_idx: int, strain_data: Dict[str, Optional[float]],
                      mappings: List[Dict]) -> None:
    """
    Aggiorna i dati Excel con i valori strain usando le mappature configurate.

    Args:
        df: DataFrame Excel
        row_idx: Indice della riga da aggiornare
        strain_data: Dati strain da inserire {mapping_name: value}
        mappings: Lista di mappature con le colonne Excel target
    """
    # Crea un dizionario mapping_name → excel_column
    name_to_column = {m['name']: m['excel_column'] for m in mappings}

    for name, value in strain_data.items():
        if value is not None and name in name_to_column:
            excel_col = name_to_column[name]
            col_idx = excel_column_to_index(excel_col)

            if col_idx < len(df.columns):
                df.iloc[row_idx, col_idx] = value
                logging.info(f"Inserito {name} = {value} alla riga {row_idx}, colonna {excel_col} (idx {col_idx})")
            else:
                logging.warning(f"Colonna {excel_col} (idx {col_idx}) fuori range per {name}")


def process_xml_files(xml_folder: str, excel_file: str, config_file: str, output_file: str = None) -> None:
    """
    Processa tutti i file XML e aggiorna il file Excel usando le mappature configurate.

    Args:
        xml_folder: Cartella contenente i file XML
        excel_file: File Excel di destinazione
        config_file: File JSON con le configurazioni delle mappature
        output_file: File Excel di output (opzionale, default: sovrascrivi il file originale)
    """
    # Carica le mappature dal file JSON
    logging.info(f"Caricamento configurazione da: {config_file}")
    mappings = load_mapping_config(config_file)
    logging.info(f"Caricate {len(mappings)} mappature dal file di configurazione")
    if not os.path.exists(xml_folder):
        raise FileNotFoundError(f"Cartella XML non trovata: {xml_folder}")
    
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"File Excel non trovato: {excel_file}")
    
    # Leggi il file Excel
    logging.info(f"Caricamento file Excel: {excel_file}")
    try:
        # Determina il motore basato sull'estensione del file
        if excel_file.lower().endswith('.xlsx'):
            engine = 'openpyxl'
        elif excel_file.lower().endswith('.xls'):
            engine = 'xlrd'
        else:
            engine = None
        
        df = pd.read_excel(excel_file, sheet_name='LGENENICM-PazientiTrieste_DATA_', engine=engine)
    except Exception as e:
        logging.error(f"Errore nel caricamento del file Excel: {e}")
        # Prova con il motore alternativo se il primo fallisce
        try:
            if engine == 'openpyxl':
                logging.info("Tentativo con motore xlrd...")
                df = pd.read_excel(excel_file, sheet_name='LGENENICM-PazientiTrieste_DATA_', engine='xlrd')
            elif engine == 'xlrd':
                logging.info("Tentativo con motore openpyxl...")
                df = pd.read_excel(excel_file, sheet_name='LGENENICM-PazientiTrieste_DATA_', engine='openpyxl')
            else:
                raise
        except Exception as e2:
            logging.error(f"Errore anche con il motore alternativo: {e2}")
            raise e
    
    # Statistiche
    files_processed = 0
    patients_found = 0
    patients_not_found = 0
    
    # Processa tutti i file XML nella cartella
    xml_files = [f for f in os.listdir(xml_folder) if f.endswith('.xml')]
    logging.info(f"Trovati {len(xml_files)} file XML da processare")
    
    for xml_filename in xml_files:
        xml_path = os.path.join(xml_folder, xml_filename)
        
        # Estrai cognome e nome dal nome del file
        patient_info = parse_filename(xml_filename)
        if not patient_info:
            logging.warning(f"Impossibile estrarre cognome/nome da: {xml_filename}")
            continue
        
        cognome, nome = patient_info
        logging.info(f"Processando: {cognome} {nome} ({xml_filename})")

        # Estrai i dati strain usando le mappature configurate
        strain_data = extract_strain_values(xml_path, mappings)

        # Verifica se sono stati estratti dati
        has_data = any(v is not None for v in strain_data.values())
        if not has_data:
            logging.warning(f"Nessun dato strain trovato in {xml_filename}")
            continue

        # Trova la riga del paziente nell'Excel
        row_idx = find_patient_row(df, cognome, nome)
        if row_idx is not None:
            update_excel_data(df, row_idx, strain_data, mappings)
            patients_found += 1
            logging.info(f"Paziente {cognome} {nome} aggiornato alla riga {row_idx + 1}")
        else:
            patients_not_found += 1
            logging.warning(f"Paziente {cognome} {nome} non trovato nell'Excel")
        
        files_processed += 1
    
    # Salva il file Excel aggiornato
    output_path = output_file if output_file else excel_file
    logging.info(f"Salvataggio file Excel aggiornato: {output_path}")
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='LGENENICM-PazientiTrieste_DATA_', index=False)
    except Exception as e:
        logging.error(f"Errore nel salvataggio del file Excel: {e}")
        raise
    
    # Riepilogo finale
    logging.info(f"\n=== RIEPILOGO ===")
    logging.info(f"File XML processati: {files_processed}")
    logging.info(f"Pazienti trovati e aggiornati: {patients_found}")
    logging.info(f"Pazienti non trovati: {patients_not_found}")
    logging.info(f"File Excel aggiornato salvato in: {output_path}")


def main():
    """Funzione principale che gestisce l'input dell'utente"""
    setup_logging()

    print("=== IMPORT STRAIN XML TO EXCEL ===")
    print("Questo script importa dati strain da file XML in un file Excel")
    print("usando le mappature configurate nella GUI.\n")

    try:
        # Richiedi il file di configurazione JSON
        config_file = input("Inserisci il percorso del file di configurazione JSON: ").strip()
        if not config_file:
            print("Errore: Deve essere specificato un file di configurazione JSON")
            return

        config_file = config_file.strip('"\'')  # Rimuovi virgolette se presenti

        if not os.path.exists(config_file):
            print(f"Errore: File di configurazione non trovato: {config_file}")
            return

        # Richiedi la cartella dei file XML
        xml_folder = input("Inserisci il percorso della cartella contenente i file XML: ").strip()
        if not xml_folder:
            print("Errore: Deve essere specificata una cartella XML")
            return

        xml_folder = xml_folder.strip('"\'')  # Rimuovi virgolette se presenti

        # Richiedi il file Excel
        excel_file = input("Inserisci il percorso del file Excel di destinazione: ").strip()
        if not excel_file:
            print("Errore: Deve essere specificato un file Excel")
            return

        excel_file = excel_file.strip('"\'')  # Rimuovi virgolette se presenti

        # Chiedi se vuole creare un nuovo file o sovrascrivere
        create_copy = input("Vuoi creare una copia del file Excel originale? (s/n): ").strip().lower()

        output_file = None
        if create_copy in ['s', 'si', 'y', 'yes']:
            base_name = os.path.splitext(excel_file)[0]
            output_file = f"{base_name}_strain_updated.xlsx"
            print(f"Verrà creato il file: {output_file}")
        else:
            print("Il file Excel originale verrà aggiornato")

        # Conferma prima di procedere
        print(f"\nFile configurazione: {config_file}")
        print(f"Cartella XML: {xml_folder}")
        print(f"File Excel: {excel_file}")
        if output_file:
            print(f"File output: {output_file}")

        confirm = input("\nProcedere con l'importazione? (s/n): ").strip().lower()
        if confirm not in ['s', 'si', 'y', 'yes']:
            print("Operazione annullata")
            return

        # Esegui l'importazione
        process_xml_files(xml_folder, excel_file, config_file, output_file)
        print("\nImportazione completata con successo!")

    except KeyboardInterrupt:
        print("\nOperazione interrotta dall'utente")
    except Exception as e:
        logging.error(f"Errore durante l'esecuzione: {e}")
        print(f"Errore: {e}")
        print("Controlla il file strain_import.log per maggiori dettagli")


if __name__ == "__main__":
    main()