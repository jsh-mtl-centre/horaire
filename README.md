# horaire

Versionnage et affichage des calendriers ICS pour Jeunes Sportifs Hochelaga.

## Structure

```
ics_groupes/     # ICS des groupes
ics_arenas/      # ICS des arénas
calendars.json   # Index auto-généré
```

## Usage

Drop `.ics` files in `ics_groupes/` or `ics_arenas/` → GitHub Actions builds → GitHub Pages serves.

## Viewers

**Interactive calendar**: https://jsh-mtl-centre.github.io/horaire/calendrier.html  
**Simple table view**: https://jsh-mtl-centre.github.io/horaire/simple.html

Both read from the same ICS sources. Cache: 10min.

## Tech

- GitHub Actions for automation
- GitHub Pages for hosting
- Vanilla JS for frontend
- Standard iCalendar format
