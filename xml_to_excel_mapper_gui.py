#!/usr/bin/env python3
"""
GUI Wizard per mappare elementi XML a colonne Excel.

Permette agli utenti di:
1. Caricare un file XML di esempio
2. Esplorare visualmente la struttura XML
3. Selezionare elementi e attributi
4. Mappare alle colonne Excel
5. Salvare/caricare configurazioni
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import xml.etree.ElementTree as ET
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import openpyxl
from datetime import datetime
import re
from difflib import SequenceMatcher


class XMLTreeViewer(ttk.Frame):
    """Widget per visualizzare e navigare l'albero XML"""

    def __init__(self, parent, on_select_callback=None):
        super().__init__(parent)
        self.on_select_callback = on_select_callback
        self.xml_root = None
        self.all_items = []  # Lista di tutti gli item nel tree per la ricerca
        self.matched_items = []  # Item che matchano la ricerca corrente
        self.current_match_index = 0
        self.is_expanded = False  # Stato di espansione dell'albero

        # Frame per la ricerca
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(search_frame, text="🔍 Cerca:").pack(side=tk.LEFT, padx=(0, 5))

        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_changed)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.search_count_label = ttk.Label(search_frame, text="0 risultati", foreground='gray')
        self.search_count_label.pack(side=tk.LEFT, padx=(0, 5))

        # Pulsanti navigazione
        ttk.Button(search_frame, text="◀ Prec", command=self._previous_match, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(search_frame, text="Succ ▶", command=self._next_match, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(search_frame, text="✕ Pulisci", command=self._clear_search, width=8).pack(side=tk.LEFT, padx=2)

        # Separatore
        ttk.Separator(search_frame, orient='vertical').pack(side=tk.LEFT, fill='y', padx=5)

        # Pulsante toggle espansione
        self.toggle_expand_btn = ttk.Button(
            search_frame,
            text="⊞ Espandi tutto",
            command=self._toggle_expand,
            width=15
        )
        self.toggle_expand_btn.pack(side=tk.LEFT, padx=2)

        # Crea il treeview con scrollbar
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal")

        # Treeview
        self.tree = ttk.Treeview(
            self.tree_frame,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode='browse'
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Colonne
        self.tree['columns'] = ('xpath', 'attributes')
        self.tree.column('#0', width=300, minwidth=200)
        self.tree.column('xpath', width=400, minwidth=200)
        self.tree.column('attributes', width=400, minwidth=200)

        self.tree.heading('#0', text='Elemento XML', anchor=tk.W)
        self.tree.heading('xpath', text='XPath', anchor=tk.W)
        self.tree.heading('attributes', text='Attributi', anchor=tk.W)

        # Bind eventi
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)

    def load_xml(self, xml_file: str):
        """Carica e visualizza un file XML"""
        try:
            tree = ET.parse(xml_file)
            self.xml_root = tree.getroot()

            # Pulisci il tree
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Reset lista items per ricerca
            self.all_items = []
            self.matched_items = []

            # Popola il tree
            self._populate_tree('', self.xml_root, '')

            return True
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare XML: {e}")
            return False

    def _populate_tree(self, parent: str, element: ET.Element, xpath: str):
        """Popola ricorsivamente il treeview con gli elementi XML"""
        # Costruisci XPath con predicati per questo elemento
        # Aggiungi predicati per attributi chiave come 'nb', 'id', 'name'
        predicates = []
        for key in ['nb', 'id', 'name', 'type']:
            if key in element.attrib:
                predicates.append(f"[@{key}='{element.attrib[key]}']")

        predicate_str = ''.join(predicates)
        tag_with_predicate = f"{element.tag}{predicate_str}"
        current_xpath = f"{xpath}/{tag_with_predicate}" if xpath else element.tag

        # Formatta gli attributi
        attrs_str = ", ".join([f"{k}={v}" for k, v in element.attrib.items()])

        # Crea il display text con attributi chiave evidenziati
        display_text = element.tag
        if predicates:
            # Mostra gli attributi chiave nel nome
            key_attrs = []
            for key in ['nb', 'id', 'name', 'type']:
                if key in element.attrib:
                    key_attrs.append(f"{key}={element.attrib[key]}")
            if key_attrs:
                display_text += f" [{', '.join(key_attrs)}]"

        if element.text and element.text.strip():
            display_text += f" = '{element.text.strip()}'"

        # Inserisci nel tree
        node = self.tree.insert(
            parent,
            'end',
            text=display_text,
            values=(current_xpath, attrs_str),
            open=False
        )

        # Salva item per la ricerca (text, xpath, attrs)
        self.all_items.append({
            'id': node,
            'text': display_text.lower(),
            'xpath': current_xpath.lower(),
            'attrs': attrs_str.lower()
        })

        # Aggiungi i figli ricorsivamente
        for child in element:
            self._populate_tree(node, child, current_xpath)

    def _on_tree_select(self, event):
        """Gestisce la selezione di un elemento nel tree"""
        selection = self.tree.selection()
        if selection and self.on_select_callback:
            item = selection[0]
            values = self.tree.item(item)
            xpath = values['values'][0] if values['values'] else ''
            attrs = values['values'][1] if len(values['values']) > 1 else ''
            self.on_select_callback(xpath, attrs)

    def get_selected_xpath(self) -> Optional[str]:
        """Restituisce l'XPath dell'elemento selezionato"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            values = self.tree.item(item)
            return values['values'][0] if values['values'] else None
        return None

    def _on_search_changed(self, *args):
        """Gestisce il cambiamento del testo di ricerca"""
        search_text = self.search_var.get().lower().strip()

        # Rimuovi evidenziazione precedente
        for item_data in self.matched_items:
            self.tree.item(item_data['id'], tags=())

        self.matched_items = []
        self.current_match_index = 0

        if not search_text:
            self.search_count_label.config(text="0 risultati", foreground='gray')
            return

        # Cerca in tutti gli item
        for item_data in self.all_items:
            if (search_text in item_data['text'] or
                search_text in item_data['xpath'] or
                search_text in item_data['attrs']):
                self.matched_items.append(item_data)

        # Aggiorna contatore
        count = len(self.matched_items)
        if count > 0:
            self.search_count_label.config(
                text=f"{count} risultat{'o' if count == 1 else 'i'}",
                foreground='green'
            )
            # Evidenzia tutti i match
            for item_data in self.matched_items:
                self.tree.item(item_data['id'], tags=('match',))

            # Configura il tag per evidenziare
            self.tree.tag_configure('match', background='yellow')
            self.tree.tag_configure('current_match', background='orange')

            # Seleziona e mostra il primo risultato
            self._highlight_current_match()
        else:
            self.search_count_label.config(text="0 risultati", foreground='red')

    def _highlight_current_match(self):
        """Evidenzia il match corrente e fa lo scroll"""
        if not self.matched_items:
            return

        # Reset tag precedente
        for item_data in self.matched_items:
            self.tree.item(item_data['id'], tags=('match',))

        # Evidenzia il match corrente
        current_item = self.matched_items[self.current_match_index]
        self.tree.item(current_item['id'], tags=('current_match',))

        # Espandi i genitori per mostrare l'item
        self._expand_to_item(current_item['id'])

        # Scroll all'item
        self.tree.see(current_item['id'])

        # Seleziona l'item
        self.tree.selection_set(current_item['id'])

        # Aggiorna contatore
        count = len(self.matched_items)
        self.search_count_label.config(
            text=f"{self.current_match_index + 1}/{count} risultati",
            foreground='green'
        )

    def _expand_to_item(self, item_id):
        """Espande tutti i genitori di un item per renderlo visibile"""
        parent = self.tree.parent(item_id)
        while parent:
            self.tree.item(parent, open=True)
            parent = self.tree.parent(parent)

    def _next_match(self):
        """Passa al prossimo risultato"""
        if not self.matched_items:
            return

        self.current_match_index = (self.current_match_index + 1) % len(self.matched_items)
        self._highlight_current_match()

    def _previous_match(self):
        """Passa al risultato precedente"""
        if not self.matched_items:
            return

        self.current_match_index = (self.current_match_index - 1) % len(self.matched_items)
        self._highlight_current_match()

    def _clear_search(self):
        """Pulisce la ricerca"""
        self.search_var.set('')
        for item_data in self.matched_items:
            self.tree.item(item_data['id'], tags=())
        self.matched_items = []
        self.current_match_index = 0
        self.search_count_label.config(text="0 risultati", foreground='gray')

    def _toggle_expand(self):
        """Toggle espansione/compressione dell'albero"""
        if self.is_expanded:
            # Comprimi tutto
            def collapse_recursive(item):
                children = self.tree.get_children(item)
                for child in children:
                    collapse_recursive(child)
                self.tree.item(item, open=False)

            for item in self.tree.get_children():
                collapse_recursive(item)

            self.is_expanded = False
            self.toggle_expand_btn.config(text="⊞ Espandi tutto")
        else:
            # Espandi tutto
            def expand_recursive(item):
                self.tree.item(item, open=True)
                children = self.tree.get_children(item)
                for child in children:
                    expand_recursive(child)

            for item in self.tree.get_children():
                expand_recursive(item)

            self.is_expanded = True
            self.toggle_expand_btn.config(text="⊟ Comprimi tutto")


class MappingEditorFrame(ttk.Frame):
    """Frame per creare/modificare un mapping XML -> Excel"""

    def __init__(self, parent, on_add_callback=None, xml_root=None):
        super().__init__(parent)
        self.on_add_callback = on_add_callback
        self.xml_root = xml_root  # Reference all'XML caricato per validazione

        # Variabili
        self.xpath_var = tk.StringVar()
        self.attribute_var = tk.StringVar()
        self.excel_column_var = tk.StringVar()
        self.mapping_name_var = tk.StringVar()
        self.data_type_var = tk.StringVar(value='float')

        self._create_widgets()

    def set_xml_root(self, xml_root):
        """Imposta il reference all'XML root per la validazione"""
        self.xml_root = xml_root

    def _create_widgets(self):
        """Crea i widget del form"""
        # Nome mapping
        ttk.Label(self, text="Nome mapping:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(self, textvariable=self.mapping_name_var, width=40).grid(
            row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=5
        )

        # XPath
        ttk.Label(self, text="XPath elemento:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        xpath_entry = ttk.Entry(self, textvariable=self.xpath_var, width=40)
        xpath_entry.grid(row=1, column=1, columnspan=2, sticky='ew', padx=5, pady=5)

        # Nota per XPath
        ttk.Label(
            self,
            text="(Modificabile: puoi aggiungere predicati personalizzati come [@attr='value'])",
            font=('TkDefaultFont', 8)
        ).grid(row=2, column=1, columnspan=2, sticky='w', padx=5)

        # Attributo
        ttk.Label(self, text="Nome attributo:").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.attribute_combo = ttk.Combobox(self, textvariable=self.attribute_var, width=38)
        self.attribute_combo.grid(row=3, column=1, columnspan=2, sticky='ew', padx=5, pady=5)

        # Colonna Excel
        ttk.Label(self, text="Colonna Excel:").grid(row=4, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(self, textvariable=self.excel_column_var, width=40).grid(
            row=4, column=1, columnspan=2, sticky='ew', padx=5, pady=5
        )
        ttk.Label(self, text="(es: AB, AC, o nome colonna)", font=('TkDefaultFont', 8)).grid(
            row=5, column=1, columnspan=2, sticky='w', padx=5
        )

        # Tipo di dato
        ttk.Label(self, text="Tipo dato:").grid(row=6, column=0, sticky='w', padx=5, pady=5)
        type_combo = ttk.Combobox(
            self,
            textvariable=self.data_type_var,
            values=['float', 'int', 'string'],
            state='readonly',
            width=38
        )
        type_combo.grid(row=6, column=1, columnspan=2, sticky='ew', padx=5, pady=5)

        # Pulsante aggiungi
        ttk.Button(self, text="Aggiungi Mapping", command=self._add_mapping).grid(
            row=7, column=0, columnspan=3, pady=10
        )

        self.grid_columnconfigure(1, weight=1)

    def set_xpath_and_attributes(self, xpath: str, attributes_str: str):
        """Imposta XPath e popola la combo degli attributi"""
        self.xpath_var.set(xpath)

        # Parse gli attributi
        if attributes_str:
            attrs = [attr.split('=')[0].strip() for attr in attributes_str.split(',')]
            self.attribute_combo['values'] = attrs
            if attrs:
                self.attribute_combo.current(0)
        else:
            self.attribute_combo['values'] = []
            self.attribute_var.set('')

    def _validate_xpath(self, xpath: str, attribute: str) -> Optional[str]:
        """
        Valida l'XPath sull'XML di esempio.
        Restituisce il valore trovato se univoco, None altrimenti.
        Mostra warning/errori all'utente.
        """
        if not self.xml_root:
            # Nessun XML caricato, skippa validazione
            return None

        try:
            # Rimuovi "Report/" iniziale se presente, perché xml_root è già Report
            xpath_to_use = xpath
            if xpath.startswith('Report/'):
                xpath_to_use = xpath[7:]  # Rimuovi 'Report/'

            # ElementTree ha limitazioni con XPath: supporta solo un predicato per elemento
            # Esempio: ROI[@id='1'][@name='ROI 1'] NON funziona
            # Soluzione: trova elementi e filtra manualmente per predicati multipli

            # Separa il path dall'ultimo elemento con predicati
            import re

            # Trova tutti gli elementi che matchano l'XPath base
            # Prima proviamo con l'XPath completo (potrebbe funzionare se ha un solo predicato)
            try:
                if xpath_to_use.startswith('/'):
                    elements = self.xml_root.findall(f".{xpath_to_use}")
                else:
                    elements = self.xml_root.findall(f".//{xpath_to_use}")
            except SyntaxError:
                # Se l'XPath ha predicati multipli, ElementTree darà SyntaxError
                # In questo caso, rimuoviamo i predicati e filtriamo manualmente

                # Rimuovi tutti i predicati [@...] dall'ultimo elemento
                xpath_without_last_predicates = re.sub(r'(@\[[^\]]+\])+$', '', xpath_to_use)
                # Estrai l'ultimo elemento
                parts = xpath_to_use.split('/')
                last_part = parts[-1]

                # Trova il nome del tag (prima dei predicati)
                tag_match = re.match(r'^([^[]+)', last_part)
                if tag_match:
                    tag_name = tag_match.group(1)

                    # Estrai tutti i predicati
                    predicates = re.findall(r'@(\w+)=[\'"]([^\'"]+)[\'"]', last_part)

                    # Cerca usando solo il path senza predicati
                    base_path = '/'.join(parts[:-1] + [tag_name]) if len(parts) > 1 else tag_name

                    if base_path.startswith('/'):
                        candidates = self.xml_root.findall(f".{base_path}")
                    else:
                        candidates = self.xml_root.findall(f".//{base_path}")

                    # Filtra manualmente per tutti i predicati
                    elements = []
                    for elem in candidates:
                        matches = True
                        for attr_name, attr_value in predicates:
                            if elem.get(attr_name) != attr_value:
                                matches = False
                                break
                        if matches:
                            elements.append(elem)
                else:
                    elements = []

            if len(elements) == 0:
                response = messagebox.askyesno(
                    "XPath non trovato",
                    f"⚠️ ATTENZIONE: L'XPath non trova nessun elemento nell'XML di esempio!\n\n"
                    f"XPath: {xpath}\n\n"
                    f"Possibili cause:\n"
                    f"- XPath errato o modificato manualmente\n"
                    f"- Predicati troppo restrittivi\n"
                    f"- Elemento non presente in questo XML\n\n"
                    f"Vuoi comunque aggiungere questo mapping?",
                    icon='warning'
                )
                return "SKIP_VALIDATION" if response else None

            elif len(elements) > 1:
                # Controlla se hanno tutti lo stesso valore
                values = []
                for elem in elements:
                    val = elem.get(attribute)
                    if val:
                        values.append(val)

                unique_values = set(values)

                msg = (
                    f"⚠️ ATTENZIONE: L'XPath trova {len(elements)} elementi!\n\n"
                    f"XPath: {xpath}\n"
                    f"Attributo: {attribute}\n\n"
                )

                if len(unique_values) == 1:
                    msg += f"Tutti gli elementi hanno lo stesso valore: '{values[0]}'\n\n"
                    msg += "Vuoi comunque aggiungere questo mapping?"
                else:
                    msg += f"Gli elementi hanno valori DIVERSI:\n"
                    for i, val in enumerate(values[:5], 1):  # Mostra max 5 esempi
                        msg += f"  {i}. {val}\n"
                    if len(values) > 5:
                        msg += f"  ... e altri {len(values) - 5}\n"
                    msg += "\n⚠️ Il risultato sarà AMBIGUO!\n\n"
                    msg += "Suggerimento: Aggiungi predicati [@attr='value'] per rendere l'XPath univoco.\n\n"
                    msg += "Vuoi comunque aggiungere questo mapping?"

                response = messagebox.askyesno("XPath ambiguo", msg, icon='warning')
                return "SKIP_VALIDATION" if response else None

            else:
                # Esattamente 1 elemento trovato ✓
                element = elements[0]
                value = element.get(attribute)

                if value:
                    messagebox.showinfo(
                        "✓ XPath valido",
                        f"✓ XPath univoco! Trovato 1 elemento.\n\n"
                        f"XPath: {xpath}\n"
                        f"Attributo: {attribute}\n"
                        f"Valore estratto: '{value}'\n\n"
                        f"Il mapping verrà aggiunto.",
                        icon='info'
                    )
                    return value
                else:
                    response = messagebox.askyesno(
                        "Attributo non trovato",
                        f"⚠️ Elemento trovato ma l'attributo '{attribute}' non esiste!\n\n"
                        f"XPath: {xpath}\n\n"
                        f"Vuoi comunque aggiungere questo mapping?",
                        icon='warning'
                    )
                    return "SKIP_VALIDATION" if response else None

        except Exception as e:
            response = messagebox.askyesno(
                "Errore validazione XPath",
                f"❌ Errore durante la validazione dell'XPath:\n\n{str(e)}\n\n"
                f"XPath potrebbe avere sintassi non valida.\n\n"
                f"Vuoi comunque aggiungere questo mapping?",
                icon='error'
            )
            return "SKIP_VALIDATION" if response else None

    def _add_mapping(self):
        """Aggiunge il mapping corrente"""
        # Validazione campi obbligatori
        if not self.mapping_name_var.get().strip():
            messagebox.showwarning("Attenzione", "Inserisci un nome per il mapping")
            return

        if not self.xpath_var.get().strip():
            messagebox.showwarning("Attenzione", "Seleziona un elemento XML")
            return

        if not self.attribute_var.get().strip():
            messagebox.showwarning("Attenzione", "Seleziona un attributo XML")
            return

        if not self.excel_column_var.get().strip():
            messagebox.showwarning("Attenzione", "Inserisci la colonna Excel di destinazione")
            return

        # Validazione XPath sull'XML di esempio
        xpath = self.xpath_var.get().strip()
        attribute = self.attribute_var.get().strip()

        validation_result = self._validate_xpath(xpath, attribute)
        if validation_result is None:
            # Utente ha annullato
            return

        # Crea il mapping
        mapping = {
            'name': self.mapping_name_var.get().strip(),
            'xml_path': xpath,
            'xml_attribute': attribute,
            'excel_column': self.excel_column_var.get().strip().upper(),
            'data_type': self.data_type_var.get()
        }

        if self.on_add_callback:
            self.on_add_callback(mapping)

        # Reset form
        self.mapping_name_var.set('')
        self.excel_column_var.set('')

    def clear(self):
        """Pulisce il form"""
        self.mapping_name_var.set('')
        self.xpath_var.set('')
        self.attribute_var.set('')
        self.excel_column_var.set('')
        self.data_type_var.set('float')
        self.attribute_combo['values'] = []


class MappingsListFrame(ttk.Frame):
    """Frame per visualizzare la lista dei mapping creati"""

    def __init__(self, parent, on_delete_callback=None):
        super().__init__(parent)
        self.on_delete_callback = on_delete_callback
        self.mappings = []

        self._create_widgets()

    def _create_widgets(self):
        """Crea i widget"""
        ttk.Label(self, text="Mapping configurati:", font=('TkDefaultFont', 10, 'bold')).pack(
            anchor='w', padx=5, pady=5
        )

        # Treeview per i mapping
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")

        self.tree = ttk.Treeview(
            tree_frame,
            yscrollcommand=vsb.set,
            selectmode='browse',
            height=10
        )
        vsb.config(command=self.tree.yview)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Colonne
        self.tree['columns'] = ('xpath', 'attribute', 'excel_col', 'type')
        self.tree.column('#0', width=200, minwidth=150)
        self.tree.column('xpath', width=250, minwidth=150)
        self.tree.column('attribute', width=150, minwidth=100)
        self.tree.column('excel_col', width=100, minwidth=80)
        self.tree.column('type', width=80, minwidth=60)

        self.tree.heading('#0', text='Nome', anchor=tk.W)
        self.tree.heading('xpath', text='XPath', anchor=tk.W)
        self.tree.heading('attribute', text='Attributo', anchor=tk.W)
        self.tree.heading('excel_col', text='Colonna Excel', anchor=tk.W)
        self.tree.heading('type', text='Tipo', anchor=tk.W)

        # Pulsanti
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(btn_frame, text="Elimina Selezionato", command=self._delete_selected).pack(
            side=tk.LEFT, padx=5
        )

    def add_mapping(self, mapping: Dict):
        """Aggiunge un mapping alla lista"""
        self.mappings.append(mapping)
        self.tree.insert(
            '',
            'end',
            text=mapping['name'],
            values=(
                mapping['xml_path'],
                mapping['xml_attribute'],
                mapping['excel_column'],
                mapping['data_type']
            )
        )

    def _delete_selected(self):
        """Elimina il mapping selezionato"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Attenzione", "Seleziona un mapping da eliminare")
            return

        item = selection[0]
        index = self.tree.index(item)

        # Rimuovi dalla lista e dal tree
        del self.mappings[index]
        self.tree.delete(item)

        if self.on_delete_callback:
            self.on_delete_callback(index)

    def clear(self):
        """Pulisce la lista"""
        self.mappings.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

    def get_mappings(self) -> List[Dict]:
        """Restituisce la lista dei mapping"""
        return self.mappings.copy()

    def load_mappings(self, mappings: List[Dict]):
        """Carica una lista di mapping"""
        self.clear()
        for mapping in mappings:
            self.add_mapping(mapping)


class BatchProcessorFrame(ttk.Frame):
    """Frame per il processing batch di XML -> Excel"""

    def __init__(self, parent, root):
        super().__init__(parent)
        self.root = root  # Reference alla root window per update_idletasks
        self.config_file = None
        self.xml_folder = None
        self.excel_file = None
        self.mappings = []
        self.log_messages = []

        self._create_widgets()

    def _create_widgets(self):
        """Crea i widget del batch processor"""

        # Sezione configurazione
        config_frame = ttk.LabelFrame(self, text="Configurazione", padding=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)

        # File configurazione JSON
        ttk.Label(config_frame, text="File configurazione:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.config_label = ttk.Label(config_frame, text="Nessun file caricato", foreground='gray')
        self.config_label.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        ttk.Button(config_frame, text="Carica JSON", command=self._load_config).grid(
            row=0, column=2, padx=5, pady=5
        )

        # Cartella XML
        ttk.Label(config_frame, text="Cartella XML:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.xml_folder_label = ttk.Label(config_frame, text="Nessuna cartella selezionata", foreground='gray')
        self.xml_folder_label.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        ttk.Button(config_frame, text="Sfoglia", command=self._select_xml_folder).grid(
            row=1, column=2, padx=5, pady=5
        )

        # File Excel
        ttk.Label(config_frame, text="File Excel:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.excel_label = ttk.Label(config_frame, text="Nessun file selezionato", foreground='gray')
        self.excel_label.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        ttk.Button(config_frame, text="Sfoglia", command=self._select_excel_file).grid(
            row=2, column=2, padx=5, pady=5
        )

        # Colonne Nome e Cognome
        match_frame = ttk.LabelFrame(self, text="Matching Paziente", padding=10)
        match_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(match_frame, text="Colonna Cognome Excel:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.lastname_col_var = tk.StringVar()
        ttk.Entry(match_frame, textvariable=self.lastname_col_var, width=10).grid(
            row=0, column=1, sticky='w', padx=5, pady=5
        )
        ttk.Label(match_frame, text="(es: A)").grid(row=0, column=2, sticky='w', padx=5, pady=5)

        ttk.Label(match_frame, text="Colonna Nome Excel:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.firstname_col_var = tk.StringVar()
        ttk.Entry(match_frame, textvariable=self.firstname_col_var, width=10).grid(
            row=1, column=1, sticky='w', padx=5, pady=5
        )
        ttk.Label(match_frame, text="(es: B)").grid(row=1, column=2, sticky='w', padx=5, pady=5)

        # Informazione sul formato del nome file
        info_label = ttk.Label(
            match_frame,
            text="ℹ Il nome paziente viene estratto dal nome del file XML\n"
                 "Formato atteso: COGNOME_NOME_data_...xml",
            foreground='blue',
            font=('TkDefaultFont', 9, 'italic')
        )
        info_label.grid(row=2, column=0, columnspan=3, sticky='w', padx=5, pady=10)

        # Pulsante elabora
        process_frame = ttk.Frame(self)
        process_frame.pack(fill=tk.X, padx=10, pady=10)

        self.process_btn = ttk.Button(
            process_frame,
            text="▶ Avvia Elaborazione",
            command=self._process_batch,
            style='Accent.TButton'
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(process_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Area log
        log_frame = ttk.LabelFrame(self, text="Log Elaborazione", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Pulsanti log
        log_btn_frame = ttk.Frame(log_frame)
        log_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(log_btn_frame, text="Pulisci Log", command=self._clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_btn_frame, text="Salva Log", command=self._save_log).pack(side=tk.LEFT, padx=5)

    def _load_config(self):
        """Carica il file di configurazione JSON"""
        filename = filedialog.askopenfilename(
            title="Seleziona configurazione JSON",
            filetypes=[("File JSON", "*.json"), ("Tutti i file", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                if 'mappings' not in config:
                    messagebox.showerror("Errore", "File di configurazione non valido")
                    return

                self.config_file = filename
                self.mappings = config['mappings']
                self.config_label.config(text=Path(filename).name, foreground='black')
                self._log(f"✓ Configurazione caricata: {len(self.mappings)} mapping")
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile caricare la configurazione:\n{e}")

    def _select_xml_folder(self):
        """Seleziona la cartella contenente i file XML"""
        folder = filedialog.askdirectory(title="Seleziona cartella XML")
        if folder:
            self.xml_folder = folder
            xml_files = list(Path(folder).glob("*.xml"))
            self.xml_folder_label.config(text=f"{Path(folder).name} ({len(xml_files)} file XML)", foreground='black')
            self._log(f"✓ Cartella XML selezionata: {len(xml_files)} file trovati")

    def _select_excel_file(self):
        """Seleziona il file Excel"""
        filename = filedialog.askopenfilename(
            title="Seleziona file Excel",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")]
        )
        if filename:
            self.excel_file = filename
            self.excel_label.config(text=Path(filename).name, foreground='black')
            self._log(f"✓ File Excel selezionato: {Path(filename).name}")

    def _log(self, message: str, level: str = "INFO"):
        """Aggiunge un messaggio al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_messages.append(log_entry)
        self.root.update_idletasks()

    def _clear_log(self):
        """Pulisce il log"""
        self.log_text.delete(1.0, tk.END)
        self.log_messages.clear()

    def _save_log(self):
        """Salva il log su file"""
        if not self.log_messages:
            messagebox.showwarning("Attenzione", "Nessun log da salvare")
            return

        filename = filedialog.asksaveasfilename(
            title="Salva log",
            defaultextension=".txt",
            filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.writelines(self.log_messages)
                messagebox.showinfo("Successo", f"Log salvato in:\n{filename}")
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile salvare il log:\n{e}")

    def _normalize_name(self, name: str) -> str:
        """Normalizza un nome per il matching"""
        if not name:
            return ""
        # Rimuovi spazi extra, underscore, converti in maiuscolo
        return " ".join(name.replace("_", " ").upper().split())

    def _extract_patient_name_from_filename(self, xml_file: str) -> Optional[Tuple[str, str]]:
        """
        Estrae cognome e nome dal nome del file XML.
        Formato atteso: COGNOME_NOME_data_...xml
        Restituisce (cognome, nome) o None
        """
        try:
            filename = Path(xml_file).stem  # Nome file senza estensione
            # Dividi per underscore
            parts = filename.split('_')

            if len(parts) >= 2:
                lastname = parts[0].strip()
                firstname = parts[1].strip()
                return (lastname, firstname)

            return None
        except Exception as e:
            self._log(f"⚠ Errore parsing nome file {Path(xml_file).name}: {e}", "WARNING")
            return None

    def _find_matching_row(self, wb, lastname: str, firstname: str) -> Optional[Tuple[int, str, str]]:
        """
        Trova la riga nell'Excel che corrisponde al nome del paziente.
        Restituisce (row_number, match_status, message) o None
        match_status può essere: 'exact', 'partial', 'none'
        """
        ws = wb.active
        lastname_col = self.lastname_col_var.get().strip().upper()
        firstname_col = self.firstname_col_var.get().strip().upper()

        if not lastname_col or not firstname_col:
            return None

        # Normalizza i nomi dal filename
        lastname_norm = self._normalize_name(lastname)
        firstname_norm = self._normalize_name(firstname)

        # Cerca match esatto
        for row in range(2, ws.max_row + 1):
            lastname_cell = ws[f"{lastname_col}{row}"]
            firstname_cell = ws[f"{firstname_col}{row}"]

            excel_lastname = self._normalize_name(str(lastname_cell.value or "").strip())
            excel_firstname = self._normalize_name(str(firstname_cell.value or "").strip())

            if not excel_lastname and not excel_firstname:
                continue

            # Match esatto: cognome E nome corrispondono
            if excel_lastname == lastname_norm and excel_firstname == firstname_norm:
                return (row, 'exact', f"{lastname} {firstname}")

        # Cerca match parziale (solo cognome)
        partial_matches = []
        for row in range(2, ws.max_row + 1):
            lastname_cell = ws[f"{lastname_col}{row}"]
            firstname_cell = ws[f"{firstname_col}{row}"]

            excel_lastname = self._normalize_name(str(lastname_cell.value or "").strip())
            excel_firstname = self._normalize_name(str(firstname_cell.value or "").strip())

            if excel_lastname == lastname_norm:
                partial_matches.append((row, excel_lastname, excel_firstname))

        if len(partial_matches) == 1:
            # Un solo match parziale
            row, excel_ln, excel_fn = partial_matches[0]
            return (row, 'partial', f"{excel_ln} {excel_fn} (solo cognome match)")
        elif len(partial_matches) > 1:
            # Match ambiguo
            return (None, 'ambiguous', f"Trovati {len(partial_matches)} pazienti con cognome '{lastname}'")

        return None

    def _extract_value_from_xml(self, xml_file: str, mapping: Dict) -> Optional[str]:
        """Estrae un valore dall'XML usando un mapping"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            xpath = mapping['xml_path']
            attribute = mapping['xml_attribute']

            # Rimuovi "Report/" se presente
            if xpath.startswith('Report/'):
                xpath = xpath[7:]

            # Gestione predicati multipli (come nel _validate_xpath)
            try:
                elements = root.findall(f".//{xpath}")
            except SyntaxError:
                # Fallback per predicati multipli
                parts = xpath.split('/')
                last_part = parts[-1]
                tag_match = re.match(r'^([^[]+)', last_part)

                if tag_match:
                    tag_name = tag_match.group(1)
                    predicates = re.findall(r'@(\w+)=[\'"]([^\'"]+)[\'"]', last_part)
                    base_path = '/'.join(parts[:-1] + [tag_name]) if len(parts) > 1 else tag_name

                    candidates = root.findall(f".//{base_path}")
                    elements = []
                    for elem in candidates:
                        matches = True
                        for attr_name, attr_value in predicates:
                            if elem.get(attr_name) != attr_value:
                                matches = False
                                break
                        if matches:
                            elements.append(elem)
                else:
                    elements = []

            if elements:
                value = elements[0].get(attribute)
                return value
            return None
        except Exception as e:
            self._log(f"⚠ Errore estrazione valore da {Path(xml_file).name}: {e}", "ERROR")
            return None

    def _process_batch(self):
        """Elabora tutti i file XML"""
        # Validazione
        if not self.config_file or not self.mappings:
            messagebox.showwarning("Attenzione", "Carica prima la configurazione JSON")
            return

        if not self.xml_folder:
            messagebox.showwarning("Attenzione", "Seleziona la cartella XML")
            return

        if not self.excel_file:
            messagebox.showwarning("Attenzione", "Seleziona il file Excel")
            return

        if not self.lastname_col_var.get() or not self.firstname_col_var.get():
            messagebox.showwarning("Attenzione", "Specifica le colonne Cognome e Nome")
            return

        try:
            # Carica Excel
            self._log("=" * 60)
            self._log("Avvio elaborazione batch")
            self._log("=" * 60)
            self.progress.start()
            self.process_btn.config(state='disabled')

            wb = openpyxl.load_workbook(self.excel_file)
            ws = wb.active

            xml_files = list(Path(self.xml_folder).glob("*.xml"))
            self._log(f"File XML trovati: {len(xml_files)}")
            self._log(f"Mapping da applicare: {len(self.mappings)}")
            self._log("")

            stats = {
                'processed': 0,
                'exact_match': 0,
                'partial_match': 0,
                'not_matched': 0,
                'ambiguous': 0,
                'errors': 0
            }

            # Processa ogni XML
            for xml_file in xml_files:
                self._log(f"Elaborazione: {xml_file.name}")

                try:
                    # Estrai cognome e nome dal filename
                    name_parts = self._extract_patient_name_from_filename(str(xml_file))

                    if not name_parts:
                        self._log(f"  ✗ Impossibile estrarre nome dal filename", "WARNING")
                        stats['not_matched'] += 1
                        continue

                    lastname, firstname = name_parts
                    self._log(f"  Nome paziente: {lastname} {firstname}")

                    # Trova riga corrispondente
                    match_result = self._find_matching_row(wb, lastname, firstname)

                    if not match_result:
                        self._log(f"  ✗ Nessuna corrispondenza trovata nell'Excel", "WARNING")
                        stats['not_matched'] += 1
                        continue

                    row_num, match_status, message = match_result

                    if match_status == 'exact':
                        self._log(f"  ✓ Match ESATTO trovato: '{message}' (riga {row_num})")
                        stats['exact_match'] += 1
                    elif match_status == 'partial':
                        self._log(f"  ⚠ Match PARZIALE: '{message}' (riga {row_num})", "WARNING")
                        stats['partial_match'] += 1
                    elif match_status == 'ambiguous':
                        self._log(f"  ✗ Match AMBIGUO: {message}", "ERROR")
                        stats['ambiguous'] += 1
                        continue
                    else:
                        self._log(f"  ✗ Nessuna corrispondenza trovata", "WARNING")
                        stats['not_matched'] += 1
                        continue

                    # Applica i mapping
                    for mapping in self.mappings:
                        value = self._extract_value_from_xml(str(xml_file), mapping)

                        if value is not None:
                            col = mapping['excel_column']
                            cell = ws[f"{col}{row_num}"]

                            # Converti tipo
                            if mapping['data_type'] == 'float':
                                try:
                                    cell.value = float(value)
                                except ValueError:
                                    cell.value = value
                            elif mapping['data_type'] == 'int':
                                try:
                                    cell.value = int(float(value))
                                except ValueError:
                                    cell.value = value
                            else:
                                cell.value = value

                            self._log(f"    • {mapping['name']}: {value} → {col}{row_num}")
                        else:
                            self._log(f"    ⚠ {mapping['name']}: valore non trovato", "WARNING")

                    stats['processed'] += 1
                    self._log("")

                except Exception as e:
                    self._log(f"  ✗ ERRORE: {e}", "ERROR")
                    stats['errors'] += 1
                    continue

            # Salva Excel
            wb.save(self.excel_file)
            self._log("=" * 60)
            self._log("✓ Excel salvato con successo")
            self._log("")
            self._log("STATISTICHE:")
            self._log(f"  File processati: {stats['processed']}/{len(xml_files)}")
            self._log(f"  Match ESATTI: {stats['exact_match']}")
            self._log(f"  Match PARZIALI (solo cognome): {stats['partial_match']}")
            self._log(f"  Match AMBIGUI: {stats['ambiguous']}")
            self._log(f"  Non trovati: {stats['not_matched']}")
            self._log(f"  Errori: {stats['errors']}")
            self._log("=" * 60)

            messagebox.showinfo(
                "Completato",
                f"Elaborazione completata!\n\n"
                f"File processati: {stats['processed']}/{len(xml_files)}\n"
                f"Match ESATTI: {stats['exact_match']}\n"
                f"Match PARZIALI: {stats['partial_match']}\n"
                f"Match AMBIGUI: {stats['ambiguous']}\n"
                f"Non trovati: {stats['not_matched']}\n"
                f"Errori: {stats['errors']}"
            )

        except Exception as e:
            self._log(f"✗ ERRORE CRITICO: {e}", "ERROR")
            messagebox.showerror("Errore", f"Errore durante l'elaborazione:\n{e}")
        finally:
            self.progress.stop()
            self.process_btn.config(state='normal')


class XMLToExcelMapperApp:
    """Applicazione principale"""

    def __init__(self, root):
        self.root = root
        self.root.title("BariSircle v1.0 - Creato da Alessandro Mingardo per Andrea Barison")
        self.root.geometry("1200x800")

        # Imposta l'icona dell'applicazione
        try:
            icon_path = Path(__file__).parent / "BariSircle.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception as e:
            print(f"Impossibile caricare l'icona: {e}")

        self.current_xml_file = None
        self.current_config_file = None

        self._create_menu()
        self._create_widgets()

    def _create_menu(self):
        """Crea il menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Menu File
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Carica XML esempio", command=self._load_xml)
        file_menu.add_separator()
        file_menu.add_command(label="Salva configurazione", command=self._save_config)
        file_menu.add_command(label="Carica configurazione", command=self._load_config)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.quit)

        # Menu Info
        info_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Info", menu=info_menu)
        info_menu.add_command(label="Informazioni", command=self._show_info)

    def _create_widgets(self):
        """Crea i widget principali"""
        # Notebook con due tab
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # TAB 1: Configurazione Mapping
        config_tab = ttk.Frame(notebook)
        notebook.add(config_tab, text="Configurazione Mapping")

        # Frame superiore: info file XML
        top_frame = ttk.Frame(config_tab)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top_frame, text="File XML esempio:").pack(side=tk.LEFT, padx=5)
        self.xml_file_label = ttk.Label(top_frame, text="Nessun file caricato", foreground='gray')
        self.xml_file_label.pack(side=tk.LEFT, padx=5)

        ttk.Button(top_frame, text="Carica XML", command=self._load_xml).pack(side=tk.RIGHT, padx=5)

        # PanedWindow principale
        paned = ttk.PanedWindow(config_tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Frame sinistro: albero XML
        left_frame = ttk.LabelFrame(paned, text="Struttura XML", padding=10)
        self.xml_viewer = XMLTreeViewer(left_frame, on_select_callback=self._on_xml_select)
        self.xml_viewer.pack(fill=tk.BOTH, expand=True)
        paned.add(left_frame, weight=2)

        # Frame destro: editor mapping + lista
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)

        # Editor mapping
        editor_frame = ttk.LabelFrame(right_frame, text="Crea nuovo mapping", padding=10)
        editor_frame.pack(fill=tk.X, padx=5, pady=5)

        self.mapping_editor = MappingEditorFrame(editor_frame, on_add_callback=self._on_add_mapping)
        self.mapping_editor.pack(fill=tk.X)

        # Lista mapping
        list_frame = ttk.LabelFrame(right_frame, text="Mapping configurati", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.mappings_list = MappingsListFrame(list_frame)
        self.mappings_list.pack(fill=tk.BOTH, expand=True)

        # Frame inferiore: pulsanti azione
        bottom_frame = ttk.Frame(config_tab)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(
            bottom_frame,
            text="Salva Configurazione",
            command=self._save_config,
            style='Accent.TButton'
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            bottom_frame,
            text="Carica Configurazione",
            command=self._load_config
        ).pack(side=tk.RIGHT, padx=5)

        # TAB 2: Elaborazione Batch
        batch_tab = ttk.Frame(notebook)
        notebook.add(batch_tab, text="▶ Elaborazione Batch")

        self.batch_processor = BatchProcessorFrame(batch_tab, self.root)
        self.batch_processor.pack(fill=tk.BOTH, expand=True)

    def _load_xml(self):
        """Carica un file XML di esempio"""
        filename = filedialog.askopenfilename(
            title="Seleziona un file XML di esempio",
            filetypes=[("File XML", "*.xml"), ("Tutti i file", "*.*")]
        )

        if filename:
            if self.xml_viewer.load_xml(filename):
                self.current_xml_file = filename
                self.xml_file_label.config(text=Path(filename).name, foreground='black')
                # Passa l'XML root al mapping editor per la validazione
                self.mapping_editor.set_xml_root(self.xml_viewer.xml_root)

    def _on_xml_select(self, xpath: str, attributes: str):
        """Gestisce la selezione di un elemento XML"""
        self.mapping_editor.set_xpath_and_attributes(xpath, attributes)

    def _on_add_mapping(self, mapping: Dict):
        """Gestisce l'aggiunta di un nuovo mapping"""
        self.mappings_list.add_mapping(mapping)

    def _save_config(self):
        """Salva la configurazione corrente"""
        mappings = self.mappings_list.get_mappings()

        if not mappings:
            messagebox.showwarning("Attenzione", "Nessun mapping da salvare")
            return

        filename = filedialog.asksaveasfilename(
            title="Salva configurazione",
            defaultextension=".json",
            filetypes=[("File JSON", "*.json"), ("Tutti i file", "*.*")]
        )

        if filename:
            config = {
                'version': '1.0',
                'mappings': mappings
            }

            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)

                self.current_config_file = filename
                # Rimozione nag screen - salvataggio silenzioso
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile salvare la configurazione:\n{e}")

    def _load_config(self):
        """Carica una configurazione esistente"""
        filename = filedialog.askopenfilename(
            title="Carica configurazione",
            filetypes=[("File JSON", "*.json"), ("Tutti i file", "*.*")]
        )

        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                if 'mappings' not in config:
                    messagebox.showerror("Errore", "File di configurazione non valido")
                    return

                self.mappings_list.load_mappings(config['mappings'])
                self.current_config_file = filename
                # Rimozione nag screen - caricamento silenzioso
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile caricare la configurazione:\n{e}")

    def _show_info(self):
        """Mostra le informazioni sull'applicazione"""
        # Crea una finestra toplevel personalizzata
        info_window = tk.Toplevel(self.root)
        info_window.title("Informazioni")
        info_window.geometry("500x300")
        info_window.resizable(False, False)

        # Imposta l'icona anche per questa finestra
        try:
            icon_path = Path(__file__).parent / "BariSircle.ico"
            if icon_path.exists():
                info_window.iconbitmap(str(icon_path))
        except:
            pass

        # Centra la finestra rispetto alla finestra principale
        info_window.transient(self.root)
        info_window.grab_set()

        # Calcola posizione centrale
        info_window.update_idletasks()
        x = (info_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (info_window.winfo_screenheight() // 2) - (300 // 2)
        info_window.geometry(f"500x300+{x}+{y}")

        # Contenuto
        frame = ttk.Frame(info_window, padding=30)
        frame.pack(fill=tk.BOTH, expand=True)

        # Titolo
        title_label = ttk.Label(
            frame,
            text="BariSircle v1.0",
            font=('TkDefaultFont', 18, 'bold')
        )
        title_label.pack(pady=(10, 20))

        # Crediti
        credits_label = ttk.Label(
            frame,
            text="Core scritto da Alessandro Mingardo\nGUI scritta da Claude Code\n\nDonato ad Andrea Barison",
            font=('TkDefaultFont', 12),
            justify=tk.CENTER
        )
        credits_label.pack(pady=15)

        # Separatore per evidenziare il pulsante
        separator = ttk.Separator(frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=20)

        # Pulsante OK con sfondo evidenziato - usa tk.Button invece di ttk per avere colori
        button_frame = tk.Frame(frame, bg='#f0f0f0')
        button_frame.pack(pady=10)

        ok_button = tk.Button(
            button_frame,
            text="CHIUDI",
            command=info_window.destroy,
            width=20,
            height=2,
            bg='#0078D4',  # Blu Windows
            fg='white',
            font=('TkDefaultFont', 11, 'bold'),
            relief=tk.RAISED,
            bd=3,
            cursor='hand2',
            activebackground='#005A9E',
            activeforeground='white'
        )
        ok_button.pack(padx=20, pady=10)

        # Focus sul pulsante OK e binding Enter
        ok_button.focus_set()
        info_window.bind('<Return>', lambda e: info_window.destroy())
        info_window.bind('<Escape>', lambda e: info_window.destroy())


def main():
    """Avvia l'applicazione"""
    root = tk.Tk()
    app = XMLToExcelMapperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
