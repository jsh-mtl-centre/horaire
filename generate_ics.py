import csv
import json
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

def load_team_names(filename: str) -> Dict[str, str]:
    """Charge le nom affiché de chaque équipe (colonnes id et nom de l'onglet des équipes)"""
    try:
        df = pd.read_csv(filename, encoding='utf-8', dtype=str, keep_default_na=False)
        names = {}
        for _, row in df.iterrows():
            if row['id'].strip() and row['nom'].strip():
                names[row['id'].strip()] = row['nom'].strip()
        print(f"✅ {len(names)} noms d'équipes chargés")
        return names
    except Exception as e:
        print(f"❌ Erreur chargement noms d'équipes: {e}")
        return {}

def load_published_arenas(filename: str):
    """Noms de fichier des arénas publiées (liste « arenas » de calendars.json), ou None si illisible"""
    try:
        with open(filename, encoding='utf-8') as f:
            published = {arena['fileName'] for arena in json.load(f).get('arenas', [])}
        print(f"✅ {len(published)} arénas publiées dans {filename}")
        return published
    except Exception as e:
        print(f"❌ Erreur lecture {filename}, tous les calendriers d'arénas seront créés: {e}")
        return None

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

def create_ics_event(row: Dict, arena_addresses: Dict[str, str], arena_mode: bool = False,
                     team_names: Dict[str, str] = None) -> str:
    """Crée un événement ICS à partir d'une ligne du CSV.

    Le titre d'une joute (Commentaire « vs … » ou « @ … ») commence par 🏒, celui d'un tournoi
    (colonne Tournoi cochée) par 🏆 et celui du power skating par ⛸️. Un événement sans description
    est décrit « Pratique » (« Power skating » au besoin), et une demi-glace « Pratique (demi-glace) ».

    arena_mode : calendrier d'une aréna. Le titre donne le nom de l'équipe (team_names, sinon le code
    d'affectation), sans l'aréna, et pour une joute l'adversaire, qui n'est alors plus dans la
    description. Calendrier d'équipe : le titre garde le code d'affectation et l'aréna, et la
    description d'une joute donne « équipe vs adversaire · Match #… ».

    _groupes (calendrier d'aréna) : activité commune à plusieurs groupes, voir merge_shared_activities.
    Le titre est le commentaire et la description la liste des groupes."""

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

    commentaire = row.get('Commentaire', '').strip()
    joute = re.match(r'(vs|@) ', commentaire) is not None
    power_skating = not joute and (affectation.upper().startswith('POWERSKATING')
                                   or re.search(r'power\s*skating', commentaire, re.IGNORECASE) is not None)
    team_name = (team_names or {}).get(affectation, affectation)

    if not affectation:
        title = "Non affecté"
    elif arena_mode:
        title = clean_text_for_ics(team_name)
    else:
        title = clean_text_for_ics(affectation)

    if joute and arena_mode:
        title += f" {clean_text_for_ics(commentaire)}"
    elif arena and not arena_mode:
        title += f" ({arena})"

    groupes = row.get('_groupes')
    if groupes:
        activity_for_uid = re.sub(r'[^\w-]', '_', commentaire.lower())
        uid = f"groupes-{activity_for_uid}-{arena_for_uid}-{start_ics}@calendar.local"
        title = clean_text_for_ics(commentaire)

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

    if commentaire and not joute:
        desc_parts.append("Pratique (demi-glace)" if commentaire.lower() == 'demi-glace' else commentaire)
    if joute and not arena_mode:
        desc_parts = [" · ".join([f"{team_name} {commentaire}"] + desc_parts)]
    if not desc_parts and affectation and not joute:
        desc_parts.append("Power skating" if power_skating else "Pratique")
    if groupes:
        names = sorted(((team_names or {}).get(g, g) for g in groupes), key=category_sort_key)
        desc_parts = ["Groupes : " + ", ".join(names)]
    
    description = clean_text_for_ics(" ".join(desc_parts)) if desc_parts else ""

    if row.get('Tournoi', '').strip().upper() == 'TRUE':
        title = f"🏆 {title}"
    elif joute:
        title = f"🏒 {title}"
    elif power_skating:
        title = f"⛸️ {title}"
    
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

def category_sort_key(name: str) -> List:
    """Tri naturel sur la catégorie, le premier mot du nom (M5 < M11 < M13-A < M13-AA). Le tri
    étant stable, les groupes d'une même catégorie gardent l'ordre de la feuille."""
    return [int(part) if part.isdigit() else part for part in re.split(r'(\d+)', name.split(' ', 1)[0].lower())]

def strip_category(name: str, other: str) -> str:
    """Retire de name la catégorie qui commence aussi other (« M13-A Faucons JSH » → « Faucons JSH »)"""
    category = other.split(' ', 1)[0] + ' '
    return name[len(category):] if name.startswith(category) else name

def name_internal_opponents(rows: List[Dict], team_names: Dict[str, str]) -> List[Dict]:
    """Une joute JSH contre JSH occupe deux lignes à la même heure dans la même aréna, « vs … »
    pour l'équipe receveuse et « @ … » pour la visiteuse. Chacune reçoit en commentaire le nom de
    l'autre équipe, sans la catégorie commune (« vs Faucons JSH », « @ Phénix JSH »). La ligne de
    la visiteuse est marquée _visiteur_interne : le calendrier de l'aréna n'affiche la joute qu'une fois."""
    games = {}
    for i, row in enumerate(rows):
        if re.match(r'(vs|@) ', row.get('Commentaire', '')):
            slot = (row.get('Date'), row.get('Arena'), row.get('Heure_debut'), row.get('Heure_fin'),
                    row.get('No_match', ''))
            games.setdefault(slot, []).append(i)

    rows = list(rows)
    for indices in games.values():
        if len(indices) != 2:
            continue
        home, away = sorted(indices, key=lambda i: rows[i]['Commentaire'].startswith('@'))
        if rows[home]['Commentaire'].startswith('vs') and rows[away]['Commentaire'].startswith('@'):
            home_name = team_names.get(rows[home]['Affectation'], rows[home]['Affectation'])
            away_name = team_names.get(rows[away]['Affectation'], rows[away]['Affectation'])
            rows[home] = {**rows[home], 'Commentaire': f"vs {strip_category(away_name, home_name)}"}
            rows[away] = {**rows[away], 'Commentaire': f"@ {strip_category(home_name, away_name)}",
                          '_visiteur_interne': True}
    return rows

def merge_shared_activities(rows: List[Dict]) -> List[Dict]:
    """Calendrier d'une aréna : les groupes à la même heure avec le même commentaire, qui n'est ni une
    joute ni une pratique (power skating, hockey libre, patin libre…), ne forment qu'un événement.
    La première ligne reçoit la liste des affectations (_groupes), les autres sont retirées."""
    shared = {}
    for i, row in enumerate(rows):
        commentaire = row.get('Commentaire', '').strip()
        if not commentaire or re.match(r'(vs |@ |pratique|demi-glace)', commentaire, re.IGNORECASE):
            continue
        slot = (row.get('Date'), row.get('Heure_debut'), row.get('Heure_fin'), commentaire.lower())
        shared.setdefault(slot, []).append(i)

    rows = list(rows)
    for indices in shared.values():
        if len(indices) > 1:
            rows[indices[0]] = {**rows[indices[0]], '_groupes': [rows[i]['Affectation'] for i in indices]}
            for i in indices[1:]:
                rows[i] = None
    return [row for row in rows if row is not None]

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

def process_csv(csv_file: str, arenas_file: str, include_non_affecte: bool = False, teams_file: str = None,
                calendars_file: str = None):
    """Traite le fichier principal et génère les calendriers.

    calendars_file : calendars.json. Seules les arénas de sa liste « arenas » ont un calendrier."""

    print("🏒 Chargement des adresses des arenas...")
    arena_addresses = load_arena_addresses(arenas_file)
    team_names = load_team_names(teams_file) if teams_file else {}
    published_arenas = load_published_arenas(calendars_file) if calendars_file else None

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

        for clean_row in name_internal_opponents([row.to_dict() for _, row in df.iterrows()], team_names):
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

            event_ics = create_ics_event(clean_row, arena_addresses, team_names=team_names)
            if event_ics:
                events_by_group[group].append(event_ics)

        for clean_row in tous_rows:
            affectation = clean_row.get('Affectation', '').strip()
            arena = clean_row.get('Arena', '').strip()
            prefix = re.match(r'^(M\d+)-TOUS$', affectation, re.IGNORECASE).group(1)

            event_ics = create_ics_event(clean_row, arena_addresses, team_names=team_names)
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
                with open(filename, 'w', encoding='utf-8', newline='\n') as f:
                    f.write(calendar_content)
                print(f"✅ {filename} créé ({len(events)} événements)")
            except Exception as e:
                print(f"❌ Erreur création {filename}: {e}")
    
    arena_output_dir = "ics_arenas"
    os.makedirs(arena_output_dir, exist_ok=True)
    
    skipped_arenas = []
    for arena, rows in events_by_arena.items():
        safe_arena_name = get_safe_filename(arena)
        if published_arenas is not None and f"{safe_arena_name}.ics" not in published_arenas:
            skipped_arenas.append(arena)
            continue

        if rows:
            arena_events = []
            for row in merge_shared_activities([r for r in rows if not r.get('_visiteur_interne')]):
                event_ics = create_ics_event(row, arena_addresses, arena_mode=True, team_names=team_names)
                if event_ics:
                    arena_events.append(event_ics)
            
            if arena_events:
                calendar_content = create_ics_file(arena_events, f"Arena {arena}")
                filename = f"{arena_output_dir}/{safe_arena_name}.ics"
                
                try:
                    with open(filename, 'w', encoding='utf-8', newline='\n') as f:
                        f.write(calendar_content)
                    print(f"🏟️  {filename} créé ({len(arena_events)} événements)")
                except Exception as e:
                    print(f"❌ Erreur création {filename}: {e}")
    
    if skipped_arenas:
        print(f"⏭️  {len(skipped_arenas)} arénas absentes de calendars.json, sans calendrier: {skipped_arenas}")
    
    print(f"\n📁 Fichiers groupes dans: {output_dir}/")
    print(f"📁 Fichiers arenas dans: {arena_output_dir}/")

def main():
    print("🗓️  GÉNÉRATEUR CALENDRIERS ICS")
    print("=" * 40)
    
    INCLUDE_NON_AFFECTE = False
    
    csv_file = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTko9LGXacAQt9cVw019lXwRThQoUiQugy8yzoxSqIdt-ZjL-x8EuE2CK31AaXxyRcFnuaqtB7V_jY2/pub?gid=1017445105&single=true&output=csv"
    arenas_file = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTko9LGXacAQt9cVw019lXwRThQoUiQugy8yzoxSqIdt-ZjL-x8EuE2CK31AaXxyRcFnuaqtB7V_jY2/pub?gid=1484034317&single=true&output=csv"
    teams_file = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTko9LGXacAQt9cVw019lXwRThQoUiQugy8yzoxSqIdt-ZjL-x8EuE2CK31AaXxyRcFnuaqtB7V_jY2/pub?gid=2055974623&single=true&output=csv"
    calendars_file = "calendars.json"

    process_csv(csv_file, arenas_file, INCLUDE_NON_AFFECTE, teams_file, calendars_file)
    print("\n🎉 Terminé!")

if __name__ == "__main__":
    main()
