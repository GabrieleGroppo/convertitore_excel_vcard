#!/usr/bin/env python3
"""
Script per convertire un file Excel in file vCard
Separa automaticamente nome e cognome dalla colonna "Nome"
"""

import pandas as pd
import argparse
import os
import sys
from pathlib import Path

def clean_phone_number(phone):
    """
    Pulisce il numero di telefono rimuovendo caratteri non numerici
    e mantenendo il prefisso internazionale
    """
    if pd.isna(phone):
        return ""
    
    phone_str = str(phone).strip()
    # Rimuove spazi, trattini e altri caratteri non numerici eccetto il +
    cleaned = ''.join(c for c in phone_str if c.isdigit() or c == '+')
    
    return cleaned

def split_name(full_name):
    """
    Separa nome e cognome dalla stringa completa
    La prima parola è il cognome, il resto è il nome
    """
    if pd.isna(full_name):
        return "", ""
    
    name_parts = str(full_name).strip().split()
    
    if len(name_parts) == 0:
        return "", ""
    elif len(name_parts) == 1:
        return name_parts[0], ""  # Solo cognome
    else:
        cognome = name_parts[0]
        nome = " ".join(name_parts[1:])
        return nome, cognome

def create_vcard(nome, cognome, telefono, nome_negozio):
    """
    Crea una vCard nel formato standard
    """
    vcard = []
    vcard.append("BEGIN:VCARD")
    vcard.append("VERSION:3.0")
    
    # Nome completo per la visualizzazione
    full_name = f"{nome} {cognome}".strip()
    if full_name:
        vcard.append(f"FN:{full_name}")
        
    # Nome strutturato (Cognome;Nome;;;;;)
    vcard.append(f"N:{cognome};{nome};;;")
    
    # Telefono
    if telefono:
        vcard.append(f"TEL;TYPE=CELL:{telefono}")
    
    # Organizzazione (nome negozio)
    if nome_negozio:
        vcard.append(f"ORG:{nome_negozio}")
    
    vcard.append("END:VCARD")
    
    return "\n".join(vcard)

def excel_to_vcard(excel_file, nome_negozio, output_file=None):
    """
    Converte un file Excel in un unico file vCard
    """
    try:
        # Legge il file Excel
        print(f"Leggendo il file Excel: {excel_file}")
        df = pd.read_excel(excel_file)
        
        # Verifica che le colonne necessarie esistano
        required_columns = ['Nome', 'telefono']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Errore: Colonne mancanti nel file Excel: {missing_columns}")
            print(f"Colonne disponibili: {list(df.columns)}")
            return False
        
        # Determina il file di output
        if output_file is None:
            excel_path = Path(excel_file)
            output_file = excel_path.parent / f"{excel_path.stem}_contatti.vcf"
        else:
            output_file = Path(output_file)
        
        # Contatori per statistiche
        total_rows = len(df)
        processed_rows = 0
        skipped_rows = 0
        
        print(f"Elaborando {total_rows} righe...")
        
        # Lista per raccogliere tutte le vCard
        all_vcards = []
        
        # Processa ogni riga
        for index, row in df.iterrows():
            full_name = row['Nome']
            telefono = row['telefono']
            
            # Salta righe vuote o senza nome
            if pd.isna(full_name) or str(full_name).strip() == "":
                skipped_rows += 1
                continue
            
            # Separa nome e cognome
            nome, cognome = split_name(full_name)
            
            # Pulisce il numero di telefono
            telefono_pulito = clean_phone_number(telefono)
            
            # Crea il contenuto vCard
            vcard_content = create_vcard(nome, cognome, telefono_pulito, nome_negozio)
            
            # Aggiunge la vCard alla lista
            all_vcards.append(vcard_content)
            
            processed_rows += 1
            
            # Mostra progresso ogni 50 contatti
            if processed_rows % 50 == 0:
                print(f"Processati {processed_rows}/{total_rows} contatti...")
        
        # Salva tutte le vCard in un unico file
        if all_vcards:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(all_vcards))
            
            print(f"\nCompleto!")
            print(f"- Contatti processati: {processed_rows}")
            print(f"- Contatti saltati: {skipped_rows}")
            print(f"- File vCard salvato: {output_file}")
        else:
            print("Nessun contatto valido trovato nel file Excel")
            return False
        
        return True
        
    except FileNotFoundError:
        print(f"Errore: File non trovato: {excel_file}")
        return False
    except Exception as e:
        print(f"Errore durante l'elaborazione: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Converte un file Excel in file vCard",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi di utilizzo:
  python excel_to_vcard.py clienti.xlsx "Negozio ABC"
  python excel_to_vcard.py clienti.xlsx "Negozio ABC" --output contatti_negozio.vcf
        """
    )
    
    parser.add_argument('excel_file', 
                       help='Percorso del file Excel da convertire')
    
    parser.add_argument('nome_negozio', 
                       help='Nome del negozio da inserire nei contatti')
    
    parser.add_argument('--output', '-o',
                       help='Percorso del file vCard di output (default: nome_excel_contatti.vcf)')
    
    args = parser.parse_args()
    
    # Verifica che il file Excel esista
    if not os.path.exists(args.excel_file):
        print(f"Errore: Il file {args.excel_file} non esiste")
        sys.exit(1)
    
    # Esegue la conversione
    success = excel_to_vcard(args.excel_file, args.nome_negozio, args.output)
    
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
