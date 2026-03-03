import csv
import os
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List
import re

def clean_text_for_ics(text: str) -> str:
    """Nettoie un champ TEXT pour l'inclure dans un ICS (RFC 5545 §3.3.11)"""
    if not text:
        return ""

    text = str(text)

    # Normaliser les fins de ligne (toujours \n)
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Échapper les caractères spéciaux autorisés
    text = (
        text.replace("\\", "\\\\")  # backslash
            .replace("\n", "\\n")   # retour ligne
            .replace(";", "\\;")    # point-virgule
            .replace(",", "\\,")    # virgule
    )

    return text.strip()

def load_arena_addresses(filename: str) -> Dict[str, str]:
    """Charge le mapping arena -> adresse depuis le fichier CSV"""
    arena_map = {}
    
    try:
        df = pd.read_csv(filename, encoding='utf-8', sep=None, engine='python')
        df.columns = df.columns.str.strip()
        df = df.fillna('').astype(str)
        for col in df.columns:
            df[col] = df[col].str.strip()
        
        arena_col = None
        address_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if col_lower in ['arena', 'arene']:
                arena_col = col
            elif col_lower in ['adresse', 'address']:
                address_col = col
        
        if not arena_col or not address_col:
            print(f"📋 Colonnes disponibles: {list(df.columns)}")
            if not arena_col:
                print("⚠️ Colonne 'arena/arene' non trouvée")
            if not address_col:
                print("⚠️ Colonne 'adresse/address' non trouvée")
        
        if arena_col and address_col:
            for index, row in df.iterrows():
                arena_name = row[arena_col]
                address = row[address_col]
                if arena_name and address:
                    arena_map[arena_name] = address
        
        print(f"✅ {len(arena_map)} adresses chargées:")
        for arena, addr in arena_map.items():
            print(f"  📍 {arena} → {addr}")
            
    except Exception as e:
        print(f"❌ Erreur chargement adresses: {e}")
    
    return arena_map

def parse_date_iso(date_str):
    """Parse une date au format ISO YYYY-MM-DD ou DD/MM/YYYY"""
    try:
        if '-' in date_str and len(date_str.strip()) == 10:
            year, month, day = map(int, date_str.strip().split('-'))
            return datetime(year, month, day)
        elif '/' in date_str:
            day, month, year = map(int, date_str.strip().split('/'))
            return datetime(year, month, day)
    except (ValueError, IndexError):
        pass
    
    print(f"Attention: impossible de parser la date '{date_str}'")
    return None

def parse_time(time_str: str) -> tuple:
    """Parse l'heure et retourne (heures, minutes)"""
    if not time_str:
        return (0, 0)
    
    time_str = str(time_str).strip()
    
    try:
        if ':' in time_str:
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
            return (hours, minutes)
        elif 'h' in time_str.lower():
            parts = time_str.lower().split('h')
            hours = int(parts[0])
            minutes = int(parts[1]) if parts[1] else 0
            return (hours, minutes)
        else:
            return (int(time_str), 0)
    except:
        print(f"❌ Heure non reconnue: '{time_str}'")
        return (0, 0)
    
def sort_events_by_datetime(events: List[str]) -> List[str]:
    """Trie les événements ICS par date/heure de début"""
    def get_start_datetime(event: str) -> datetime:
        for line in event.split('\n'):
            if line.startswith('DTSTART'):
                dt_str = line.split(':', 1)[1].strip()
                return datetime.strptime(dt_str, '%Y%m%dT%H%M%S')
        return datetime.min
    
    return sorted(events, key=get_start_datetime)    

def get_group_name(affectation: str) -> str:
    """Retourne le nom du groupe selon l'affectation"""
    if not affectation or affectation.strip() == '':
        return "Non affecté"
    return affectation.strip()

def create_ics_event(row: Dict, arena_addresses: Dict[str, str]) -> str:
    """Crée un événement ICS à partir d'une ligne du CSV"""
    
    date_obj = parse_date_iso(row.get('Date', ''))
    if not date_obj:
        return ""
    
    start_h, start_m = parse_time(row.get('Heure_debut', ''))
    end_h, end_m = parse_time(row.get('Heure_fin', ''))
    
    start_dt = date_obj.replace(hour=start_h, minute=start_m)
    end_dt = date_obj.replace(hour=end_h, minute=end_m)
    
    start_ics = start_dt.strftime('%Y%m%dT%H%M%S')
    end_ics = end_dt.strftime('%Y%m%dT%H%M%S')
    
    affectation = row.get('Affectation', '').strip()
    arena = row.get('Arena', '')
    
    affectation_for_uid = affectation if affectation else "non_affecte"
    safe_affectation = affectation_for_uid.replace('&', '_')
    safe_affectation = re.sub(r'[^\w-]', '_', safe_affectation.lower())
    arena_for_uid = arena.replace('&', '_') if arena else "no_arena"
    arena_for_uid = re.sub(r'[^\w-]', '_', arena_for_uid.lower())
    
    uid = f"{safe_affectation}-{arena_for_uid}-{start_ics}@calendar.local"
    
    if not affectation:
        title = "Non affecté"
    else:
        title = clean_text_for_ics(affectation)
    
    if arena:
        title += f" ({arena})"
    
    location_parts = []
    if arena:
        location_parts.append(arena)
        if arena in arena_addresses:
            location_parts.append(arena_addresses[arena])
        else:
            for arena_key, address in arena_addresses.items():
                if arena.lower() in arena_key.lower() or arena_key.lower() in arena.lower():
                    location_parts.append(address)
                    print(f"🎯 Correspondance partielle: {arena} → {arena_key}")
                    break
            else:
                print(f"⚠️  Pas d'adresse pour: {arena}")
    
    location = clean_text_for_ics(", ".join(location_parts))
    
    desc_parts = []
    no_match = row.get('No_match', '').strip()
    if no_match:
        try:
            no_match_int = int(float(no_match))
            desc_parts.append(f"Match #{no_match_int}")
        except (ValueError, TypeError):
            desc_parts.append(f"Match #{no_match}")
    
    commentaire = row.get('Commentaire', '').strip()
    if commentaire:
        desc_parts.append(commentaire)
    
    description = clean_text_for_ics(" ".join(desc_parts)) if desc_parts else ""
    
    event = f"""BEGIN:VEVENT
UID:{uid}
DTSTART;TZID=America/Montreal:{start_ics}
DTEND;TZID=America/Montreal:{end_ics}
SUMMARY:{title}
LOCATION:{location}
DESCRIPTION:{description}
STATUS:CONFIRMED
END:VEVENT"""
    
    return event

def create_ics_file(events: List[str], calendar_name: str) -> str:
    """Crée le contenu complet d'un fichier ICS"""
    sorted_events = sort_events_by_datetime(events)
    
    header = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Calendrier Generator//FR
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-CALNAME:{clean_text_for_ics(calendar_name)}
X-WR-TIMEZONE:America/Montreal"""
    
    footer = "END:VCALENDAR"
    return header + "\n" + "\n".join(sorted_events) + "\n" + footer

def get_safe_filename(group_name: str) -> str:
    """Génère un nom de fichier sécurisé selon le groupe"""
    if group_name == "Non affecté":
        return "non_affecte"
    safe_name = group_name.replace('&', '_')
    safe_name = re.sub(r'[^\w\s-]', '', safe_name).strip()
    safe_name = re.sub(r'\s+', '_', safe_name)
    return safe_name

def process_csv(csv_file: str, arenas_file: str, include_non_affecte: bool = False):
    """Traite le fichier principal et génère les calendriers"""
    
    print("🏒 Chargement des adresses des arenas...")
    arena_addresses = load_arena_addresses(arenas_file)
    
    print(f"\n📖 Lecture de {csv_file}...")
    print(f"📋 Créneaux non affectés: {'✅ Inclus' if include_non_affecte else '❌ Exclus'}")
    
    events_by_group = {}
    events_by_arena = {}
    total_processed = 0
    non_affecte_count = 0
    excluded_count = 0

    try:
        df = pd.read_csv(csv_file, encoding='utf-8', sep=None, engine='python')
        
        print(f"📋 Délimiteur détecté automatiquement")
        print(f"📋 Colonnes: {list(df.columns)}")
        
        df.columns = df.columns.str.strip()
        df = df.fillna('').astype(str)
        for col in df.columns:
            df[col] = df[col].str.strip()
        
        total_processed = len(df)
        tous_rows = []

        for index, row in df.iterrows():
            clean_row = row.to_dict()
            affectation = clean_row.get('Affectation', '').strip()
            arena = clean_row.get('Arena', '').strip()
            group = get_group_name(affectation)

            if group == "Non affecté":
                non_affecte_count += 1
                if not include_non_affecte:
                    excluded_count += 1
                    continue

            if re.match(r'^M\d+-TOUS$', affectation, re.IGNORECASE):
                tous_rows.append(clean_row)
                continue

            if group not in events_by_group:
                events_by_group[group] = []

            if affectation and arena:
                if arena not in events_by_arena:
                    events_by_arena[arena] = []
                events_by_arena[arena].append(clean_row)

            event_ics = create_ics_event(clean_row, arena_addresses)
            if event_ics:
                events_by_group[group].append(event_ics)

        for clean_row in tous_rows:
            affectation = clean_row.get('Affectation', '').strip()
            arena = clean_row.get('Arena', '').strip()
            prefix = re.match(r'^(M\d+)-TOUS$', affectation, re.IGNORECASE).group(1)

            event_ics = create_ics_event(clean_row, arena_addresses)
            if not event_ics:
                continue

            for existing_group in events_by_group:
                if re.match(rf'^{re.escape(prefix)}-[A-C]', existing_group, re.IGNORECASE):
                    events_by_group[existing_group].append(event_ics)

            if arena:
                if arena not in events_by_arena:
                    events_by_arena[arena] = []
                events_by_arena[arena].append(clean_row)
    
    except Exception as e:
        print(f"❌ Erreur lecture CSV: {e}")
        return
    
    print(f"\n📊 {total_processed} événements traités")
    if non_affecte_count > 0:
        print(f"⚠️  {non_affecte_count} créneaux non affectés détectés")
        if excluded_count > 0:
            print(f"🚫 {excluded_count} créneaux non affectés exclus")
    print(f"👥 Groupes: {list(events_by_group.keys())}")
    print(f"🏟️  Arenas: {list(events_by_arena.keys())}")
    
    output_dir = "ics_groupes"
    os.makedirs(output_dir, exist_ok=True)
    
    for group, events in events_by_group.items():
        if events:
            calendar_content = create_ics_file(events, f"Calendrier {group}")
            safe_name = get_safe_filename(group)
            filename = f"{output_dir}/{safe_name}.ics"
            
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(calendar_content)
                print(f"✅ {filename} créé ({len(events)} événements)")
            except Exception as e:
                print(f"❌ Erreur création {filename}: {e}")
    
    arena_output_dir = "ics_arenas"
    os.makedirs(arena_output_dir, exist_ok=True)
    
    for arena, rows in events_by_arena.items():
        if rows:
            arena_events = []
            for row in rows:
                event_ics = create_ics_event(row, arena_addresses)
                if event_ics:
                    arena_events.append(event_ics)
            
            if arena_events:
                calendar_content = create_ics_file(arena_events, f"Arena {arena}")
                safe_arena_name = get_safe_filename(arena)
                filename = f"{arena_output_dir}/{safe_arena_name}.ics"
                
                try:
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(calendar_content)
                    print(f"🏟️  {filename} créé ({len(arena_events)} événements)")
                except Exception as e:
                    print(f"❌ Erreur création {filename}: {e}")
    
    print(f"\n📁 Fichiers groupes dans: {output_dir}/")
    print(f"📁 Fichiers arenas dans: {arena_output_dir}/")

def main():
    print("🗓️  GÉNÉRATEUR CALENDRIERS ICS")
    print("=" * 40)
    
    INCLUDE_NON_AFFECTE = False
    
    csv_file = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTko9LGXacAQt9cVw019lXwRThQoUiQugy8yzoxSqIdt-ZjL-x8EuE2CK31AaXxyRcFnuaqtB7V_jY2/pub?gid=1017445105&single=true&output=csv"
    arenas_file = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTko9LGXacAQt9cVw019lXwRThQoUiQugy8yzoxSqIdt-ZjL-x8EuE2CK31AaXxyRcFnuaqtB7V_jY2/pub?gid=1484034317&single=true&output=csv"
        
    process_csv(csv_file, arenas_file, INCLUDE_NON_AFFECTE)
    print("\n🎉 Terminé!")

if __name__ == "__main__":
    main()
