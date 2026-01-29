# ConverterToPdf - Guide pour Claude

## Vue d'ensemble

Convertisseur de documents en PDF avec architecture modulaire Python 3.10+.

**Repository**: https://github.com/Pat2805/ConverterToPdf.git
**Version actuelle**: 1.0 (commit f2323f5)

## Architecture

```
converter_pdf/
├── __init__.py              # Version, exports
├── __main__.py              # Point d'entrée: python -m converter_pdf
├── cli.py                   # Parsing arguments (argparse)
├── config.py                # Configuration (dataclass + YAML)
├── logger.py                # Logging structuré (module logging)
├── processor.py             # Orchestration du traitement fichiers
├── report.py                # Rapport de session (statistiques, erreurs)
├── com_utils.py             # Utilitaires COM robustes (Word/Excel)
└── converters/
    ├── __init__.py          # Registry des convertisseurs (get_converter_chain)
    ├── base.py              # Classe abstraite BaseConverter
    ├── office.py            # Word/Excel/PowerPoint via COM (DispatchEx)
    ├── libreoffice.py       # Conversion LibreOffice headless
    ├── image.py             # Images (PIL)
    ├── html.py              # HTML via Chrome/Edge headless
    ├── text.py              # TXT/LOG via ReportLab
    ├── xml_converter.py     # XML via ReportLab
    ├── msg.py               # Outlook MSG (extract_msg)
    ├── archive.py           # ZIP, RAR, 7Z, TAR.GZ
    └── reportlab_fallback.py # Fallback Word/Excel
```

## Configuration

Fichier `.converterrc` (YAML) chargé automatiquement depuis:
1. Répertoire de travail courant
2. Répertoire du package
3. Répertoire utilisateur (~/)

### Options principales

| Option | Type | Défaut | Description |
|--------|------|--------|-------------|
| `method` | str | "auto" | auto, office, libreoffice, reportlab |
| `keep_extension` | bool | true | doc.docx -> doc.docx.pdf |
| `recursive` | bool | true | Parcourir sous-dossiers |
| `delete_source` | bool | false | Supprimer originaux après conversion |
| `report_enabled` | bool | true | Générer rapport de session |
| `log_level` | str | "INFO" | DEBUG, INFO, WARNING, ERROR |
| `office_timeout` | int | 60 | Timeout COM (secondes) |

## Convertisseurs

### Office COM (`office.py`)
- **Important**: Toujours utiliser `DispatchEx` (pas `Dispatch`) pour créer une nouvelle instance
- Évite les conflits quand Word/Excel est déjà ouvert
- Gestion des fichiers protégés par mot de passe (status `skipped_password`)

### MSG (`msg.py`)
- Utilise `extract_msg` pour parser les fichiers Outlook
- **Filtrage des petites images**: logos, signatures, tracking pixels
  - Seuils: < 10KB ou < 100x100 pixels avec nom suspect
  - Très petites: < 5KB ET < 50x50 pixels (filtrées même sans nom suspect)
- **Gestion des doublons**: `image.jpg`, `image (1).jpg`, `image (2).jpg`
- **Dossier de sortie**: `message.msg/` (pas `.pdf` car c'est un dossier)

### Archive (`archive.py`)
- Formats: ZIP (natif), TAR/TAR.GZ/TAR.BZ2 (natif), RAR (rarfile), 7Z (py7zr)
- **Nommage du dossier**: `archive.zip` -> `archive/` (sans extension)
- **Anti-duplication**: Si l'archive contient uniquement un dossier du même nom, on évite `test/test/`
- **Préservation des originaux**: Les fichiers extraits sont conservés sauf si `delete_source: true`

## Rapport de session (`report.py`)

Génère `conversion_report_YYYYMMDD_HHMMSS.txt` avec:
- Statistiques globales (fichiers, tailles, durées)
- Détail par type de fichier
- **Liste des conversions réussies** (source, dest, méthode, durée)
- **Échecs détaillés** (chemin, raison, exception)
- Fichiers protégés par mot de passe

## Utilisation CLI

```bash
# Convertir un répertoire
python -m converter_pdf /chemin/vers/dossier

# Options courantes
python -m converter_pdf /chemin -r              # Récursif
python -m converter_pdf /chemin -f              # Force reconversion
python -m converter_pdf /chemin -d              # Supprimer originaux
python -m converter_pdf /chemin --no-report     # Sans rapport
python -m converter_pdf /chemin --log-level DEBUG
python -m converter_pdf /chemin --method office # Forcer méthode
```

## Points d'attention pour le développement

### COM/Office
- Toujours `DispatchEx` pour nouvelle instance
- Context managers pour garantir `Quit()` et `CoUninitialize()`
- Timeout sur les opérations (évite blocages)

### Nommage des dossiers
- MSG: `source.msg/` (nom du fichier source)
- Archive: `source/` (sans extension .zip/.rar/.7z)
- Éviter les doubles dossiers (test.zip/test/ -> test/)

### Fichiers extraits
- Conserver les originaux par défaut
- Supprimer uniquement si `config.delete_source = True`

### Rapport
- Collecter les conversions réussies (pas seulement les erreurs)
- Ne pas mentionner les skips dans le rapport détaillé

## Dépendances

**Requises**:
- Python 3.10+
- pywin32 (COM Windows)
- Pillow (images)
- ReportLab (PDF)

**Optionnelles**:
- pyyaml (config YAML)
- extract_msg (fichiers MSG)
- rarfile + unrar (archives RAR)
- py7zr (archives 7Z)

## Tests manuels suggérés

1. Conversion avec Office fermé
2. Conversion avec Word/Excel ouvert (autre document)
3. Fichier protégé par mot de passe
4. Archive avec dossier du même nom (test.zip/test/)
5. MSG avec pièces jointes dupliquées
6. Vérification du rapport de session
