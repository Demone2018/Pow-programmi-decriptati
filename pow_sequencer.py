#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
POW Program Sequencer
=====================
Genera file MDB combinando i programmi 30, 31, 32, 33 nella sequenza desiderata.

Uso:
    python pow_sequencer.py sequenza.txt output.mdb

Oppure:
    python pow_sequencer.py --interactive

Il file sequenza.txt deve contenere un numero di programma per riga:
    30
    32
    33
"""

import os
import sys
import shutil
import struct
import re
from pathlib import Path

# Informazioni sui programmi
PROGRAMS = {
    30: {"name": "30IGNIT", "functions": 12, "max_line": 11, "desc": "Accensione"},
    31: {"name": "31NOWELD", "functions": 39, "max_line": 38, "desc": "No saldatura"},
    32: {"name": "32WELD", "functions": 49, "max_line": 48, "desc": "Saldatura"},
    33: {"name": "33DWNSLP", "functions": 49, "max_line": 48, "desc": "Downslope"},
}

def find_source_files(base_path):
    """Trova i file MDB sorgente."""
    sources = {}
    for prog_num, info in PROGRAMS.items():
        mdb_file = base_path / f"{info['name']}.mdb"
        xml_file = base_path / f"{info['name']}.xml"

        if mdb_file.exists():
            sources[prog_num] = {"mdb": mdb_file, "xml": xml_file if xml_file.exists() else None}
        else:
            print(f"ATTENZIONE: File {mdb_file} non trovato")

    return sources

def read_sequence(sequence_file):
    """Legge la sequenza da file."""
    sequence = []

    with open(sequence_file, 'r') as f:
        for line_num, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            try:
                prog_num = int(line)
                if prog_num not in PROGRAMS:
                    print(f"Errore riga {line_num}: programma {prog_num} non valido (usa 30, 31, 32, 33)")
                    return None
                sequence.append(prog_num)
            except ValueError:
                print(f"Errore riga {line_num}: '{line}' non è un numero valido")
                return None

    return sequence

def interactive_sequence():
    """Modo interattivo per definire la sequenza."""
    print("\n=== POW Program Sequencer - Modo Interattivo ===\n")
    print("Programmi disponibili:")
    for num, info in PROGRAMS.items():
        print(f"  {num} = {info['name']} ({info['desc']}) - {info['functions']} funzioni")

    print("\nInserisci la sequenza dei programmi (uno per riga)")
    print("Digita 'fine' per terminare, 'cancella' per ricominciare\n")

    sequence = []
    while True:
        try:
            user_input = input(f"Programma {len(sequence) + 1}: ").strip().lower()

            if user_input == 'fine':
                if not sequence:
                    print("Sequenza vuota. Inserisci almeno un programma.")
                    continue
                break
            elif user_input == 'cancella':
                sequence = []
                print("Sequenza cancellata. Ricomincia.\n")
                continue
            elif user_input == 'aiuto':
                print("Comandi: 30, 31, 32, 33, fine, cancella, aiuto")
                continue

            prog_num = int(user_input)
            if prog_num not in PROGRAMS:
                print(f"Programma {prog_num} non valido. Usa: 30, 31, 32, 33")
                continue

            sequence.append(prog_num)
            print(f"  Aggiunto: {PROGRAMS[prog_num]['name']}")

        except ValueError:
            print("Inserisci un numero valido (30, 31, 32, 33) o un comando")
        except KeyboardInterrupt:
            print("\n\nOperazione annullata.")
            return None

    return sequence

def generate_mdb_simple(sequence, source_files, output_path):
    """
    Genera il file MDB copiando il primo e appendendo info sulla sequenza.

    NOTA: Questa è una versione semplificata.
    Per la manipolazione completa del database Access,
    usare la versione VBA con DAO/ADO su Windows.
    """

    if not sequence:
        print("Errore: sequenza vuota")
        return False

    # Usa il primo programma come base
    first_prog = sequence[0]
    source_mdb = source_files[first_prog]["mdb"]

    # Copia il file base
    shutil.copy2(source_mdb, output_path)

    # Se c'è solo un programma, abbiamo finito
    if len(sequence) == 1:
        print(f"File generato (programma singolo): {output_path}")
        return True

    # Per sequenze multiple, dobbiamo modificare il database
    # Questa operazione richiede accesso alle tabelle Access

    print("\n" + "="*50)
    print("ATTENZIONE: Generazione sequenza multipla")
    print("="*50)
    print("\nPer combinare più programmi in un unico MDB,")
    print("è necessario usare la macro VBA su Windows con Access.")
    print("\nFile copiato come base:", output_path)
    print("\nSequenza richiesta:")

    total_functions = 0
    for i, prog in enumerate(sequence, 1):
        info = PROGRAMS[prog]
        print(f"  {i}. {info['name']} - {info['functions']} funzioni")
        total_functions += info['functions']

    print(f"\nTotale funzioni: {total_functions}")

    # Genera anche un file di istruzioni
    instructions_file = output_path.with_suffix('.txt')
    with open(instructions_file, 'w') as f:
        f.write("POW Program Sequencer - Istruzioni\n")
        f.write("="*40 + "\n\n")
        f.write("Sequenza richiesta:\n")
        for i, prog in enumerate(sequence, 1):
            info = PROGRAMS[prog]
            f.write(f"  {i}. Programma {prog} ({info['name']})\n")
        f.write(f"\nTotale funzioni: {total_functions}\n\n")
        f.write("Per completare la generazione:\n")
        f.write("1. Apri il file Excel POW_Sequencer.xlsm\n")
        f.write("2. Inserisci la sequenza nel foglio\n")
        f.write("3. Esegui la macro 'GenerateMDB'\n")

    print(f"\nIstruzioni salvate in: {instructions_file}")

    return True

def generate_xml_sequence(sequence, source_files, output_path):
    """
    Genera un file XML combinato con la sequenza di programmi.
    Questo può essere più utile per analisi e debug.
    """

    output_xml = output_path.with_suffix('.xml')

    combined_functions = []
    combined_configs = []
    current_line_offset = 0

    for prog_num in sequence:
        xml_file = source_files[prog_num].get("xml")
        if not xml_file or not xml_file.exists():
            print(f"File XML non trovato per programma {prog_num}")
            continue

        with open(xml_file, 'r', encoding='iso-8859-1') as f:
            content = f.read()

        # Estrai CONFIG e FUNCTION
        configs = re.findall(r'<CONFIG.*?</CONFIG>', content, re.DOTALL)
        functions = re.findall(r'<FUNCTION.*?</FUNCTION>', content, re.DOTALL)

        # Aggiungi CONFIG (senza modifiche)
        combined_configs.extend(configs)

        # Aggiungi FUNCTION con lineNumber aggiornato
        for func in functions:
            # Aggiorna lineNumber
            def update_line(match):
                old_line = int(match.group(1))
                new_line = old_line + current_line_offset
                return f'lineNumber="{new_line}"'

            updated_func = re.sub(r'lineNumber="(\d+)"', update_line, func)
            combined_functions.append(updated_func)

        # Aggiorna offset
        current_line_offset += PROGRAMS[prog_num]["max_line"]

    # Genera XML combinato
    with open(output_xml, 'w', encoding='iso-8859-1') as f:
        f.write('<?xml version="1.0" encoding="ISO-8859-1"?>\n')
        f.write(f'<COMBINED_PROGRAM sequence="{",".join(map(str, sequence))}" ')
        f.write(f'total_functions="{len(combined_functions)}">\n')

        f.write('  <!-- CONFIGURAZIONI -->\n')
        for config in combined_configs:
            f.write('  ' + config + '\n')

        f.write('  <!-- FUNZIONI -->\n')
        for func in combined_functions:
            f.write('  ' + func + '\n')

        f.write('</COMBINED_PROGRAM>\n')

    print(f"File XML combinato generato: {output_xml}")
    return True

def main():
    # Determina il percorso base (directory dello script)
    base_path = Path(__file__).parent

    # Trova i file sorgente
    source_files = find_source_files(base_path)

    if not source_files:
        print("Errore: nessun file MDB sorgente trovato")
        print(f"Cercato in: {base_path}")
        sys.exit(1)

    # Determina la modalità
    if len(sys.argv) == 1 or sys.argv[1] == '--interactive':
        # Modo interattivo
        sequence = interactive_sequence()
        if not sequence:
            sys.exit(1)

        # Chiedi nome output
        output_name = input("\nNome file output (default: sequenza.mdb): ").strip()
        if not output_name:
            output_name = "sequenza.mdb"
        if not output_name.endswith('.mdb'):
            output_name += '.mdb'

        output_path = base_path / output_name

    elif len(sys.argv) >= 3:
        # Modo da riga di comando
        sequence_file = sys.argv[1]
        output_path = Path(sys.argv[2])

        sequence = read_sequence(sequence_file)
        if not sequence:
            sys.exit(1)

    else:
        print("Uso: python pow_sequencer.py <sequenza.txt> <output.mdb>")
        print("     python pow_sequencer.py --interactive")
        sys.exit(1)

    # Mostra riepilogo
    print("\n=== Riepilogo Sequenza ===")
    total = 0
    for i, prog in enumerate(sequence, 1):
        info = PROGRAMS[prog]
        print(f"  {i}. {info['name']} - {info['functions']} funzioni")
        total += info['functions']
    print(f"  Totale: {total} funzioni\n")

    # Genera i file
    generate_mdb_simple(sequence, source_files, output_path)
    generate_xml_sequence(sequence, source_files, output_path)

    print("\n=== Completato ===")

if __name__ == "__main__":
    main()
