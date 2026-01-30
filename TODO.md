# TODO - ConverterToPdf

## Priorité Haute

### Gestion de la longueur des chemins (NAS compatibility)

**Problème**: Les NAS et certains systèmes Windows ont une limite de 260 caractères pour les chemins complets. Les archives imbriquées et les noms longs peuvent facilement dépasser cette limite.

**Solution proposée**:

1. **Raccourcissement intelligent des noms**
   - Priorité aux répertoires (impact sur tous les sous-fichiers)
   - Puis aux noms de fichiers si nécessaire
   - Conserver l'extension originale
   - Stratégies de raccourcissement :
     - Supprimer les mots redondants (copy, copie, final, v1, etc.)
     - Abréger les mots courants (document -> doc, rapport -> rpt)
     - Tronquer en gardant début + fin (ex: `très_long_nom_de_fichier.pdf` -> `très_lon...hier.pdf`)
     - Hash court si collision de noms

2. **Fichier de correspondance (mapping log)**
   - Créer `_path_mapping.json` ou `_path_mapping.csv` dans le dossier de sortie
   - Format: `{"chemin_court": "chemin_original_complet", ...}`
   - Permettre la restauration des noms originaux si besoin

3. **Options CLI**
   - `--max-path-length N` : Limite maximale (défaut: 260 pour Windows, 4096 pour Linux)
   - `--shorten-paths` : Activer le raccourcissement automatique
   - `--path-mapping-file` : Chemin du fichier de mapping

4. **Implémentation**
   - Calculer la longueur totale avant création
   - Si dépassement : raccourcir d'abord le répertoire parent, puis le fichier
   - Logger chaque renommage dans le fichier de mapping
   - Avertissement dans le rapport de session

**Fichiers à modifier**:
- `config.py` : Nouvelles options
- `cli.py` : Arguments CLI
- `processor.py` : Logique de vérification/raccourcissement
- `archive.py` / `msg.py` : Appliquer lors de l'extraction

---

## Priorité Moyenne

### Support de nouveaux formats

- [ ] **EML** - Emails standards (pas seulement MSG Outlook)
- [ ] **Markdown** (.md) - Via conversion HTML intermédiaire
- [ ] **CSV** - Tableaux en PDF
- [ ] **JSON** - Formaté avec coloration syntaxique

### Amélioration du rapport

- [ ] Export HTML avec liens cliquables
- [ ] Export JSON pour intégration avec d'autres outils
- [ ] Graphiques de statistiques (optionnel, nécessite matplotlib)

### PDF protégés par mot de passe

- [ ] Option `--pdf-password` pour fournir un mot de passe
- [ ] Fichier de mots de passe (un par ligne, essayés séquentiellement)
- [ ] Déverrouillage et reconversion

---

## Priorité Basse

### Performance

- [ ] Parallélisation des conversions (multithread/multiprocess)
- [ ] Cache des instances COM (éviter création/destruction répétées)
- [ ] Mode batch pour Office (ouvrir une fois, convertir plusieurs)

### Fonctionnalités avancées

- [ ] **Mode watch** - Surveillance continue d'un dossier
- [ ] **OCR** - Reconnaissance de texte sur images/PDF scannés (tesseract)
- [ ] **Interface graphique** - GUI simple (tkinter ou PyQt)
- [ ] **Fusion PDF** - Combiner plusieurs fichiers en un seul

### Intégrations

- [ ] Plugin pour Explorateur Windows (menu contextuel)
- [ ] Script PowerShell pour intégration NAS Synology/QNAP
- [ ] GitHub Actions pour CI/CD

---

## Bugs connus

*(Aucun bug connu actuellement)*

---

## Notes techniques

### Limites de chemins par système
| Système | Limite par défaut | Limite étendue |
|---------|-------------------|----------------|
| Windows | 260 caractères | 32767 (avec préfixe `\\?\`) |
| Linux | 4096 caractères | - |
| macOS | 1024 caractères | - |
| NAS (SMB) | Variable | Souvent 255 par composant |

### Références
- [Windows Long Paths](https://docs.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation)
- [Python pathlib et chemins longs](https://docs.python.org/3/library/pathlib.html)
