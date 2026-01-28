
# 📄 convertir_pdf – Documentation complète

## 1. Objectif général

Ce programme est un **utilitaire Python sous Windows** destiné à :

* parcourir **récursivement** un dossier de documents hétérogènes,
* convertir chaque fichier supporté en **PDF de qualité**,
* **sans OCR par défaut**,
* avec **traçabilité complète** via un journal CSV,
* en vue d’un **OCR ultérieur réalisé avec Adobe Acrobat Pro**.

Il est conçu pour des **cas sérieux** (juridique, notarial, bancaire, archives), où :

* un mauvais PDF est pire qu’un fichier ignoré,
* la reproductibilité et l’audit sont essentiels.

---

## 2. OCR : présent dans le code, mais volontairement non utilisé

### Capacités OCR existantes

Le code **intègre déjà** :

* détection des moteurs OCR :

  * Tesseract
  * EasyOCR
  * PaddleOCR
* support du français
* logique pour créer des PDF avec couche texte

### Décision volontaire actuelle

⚠️ **L’OCR n’est PAS utilisé dans le workflow actuel**.

* ❌ Pas d’OCR automatique
* ❌ Pas d’OCR sur images
* ❌ Pas d’OCR sur PDF existants

👉 **L’OCR est délégué à Adobe Acrobat Pro**, car :

* meilleure qualité globale
* “Améliorer le document” plus performant
* meilleure conformité juridique
* meilleure gestion des tableaux, en-têtes, structures

👉 Le script doit donc produire :

* des **PDF image-only propres**, ou
* des **PDF texte natifs** (Word/Excel/HTML),
  et **ne jamais lancer d’OCR implicitement**.

---

## 3. Formats pris en charge

### 3.1 Formats convertis en PDF

#### Documents Office

* `.doc`, `.docx`
* `.rtf`
* `.odt`
* `.xls`, `.xlsx`

#### Images

* `.jpg`, `.jpeg`
* `.png`
* `.webp`
* `.tif`, `.tiff`

#### Texte brut

* `.txt`
* `.log`

#### HTML

* `.htm`
* `.html`

#### Emails Outlook

* `.msg`

#### Données

* `.xml`

---

### 3.2 Formats explicitement ignorés

* `.pdf` (déjà PDF, pas d’OCR ici)
* `.mp4`, `.m4a`
* tout type inconnu

---

## 4. Règles fonctionnelles essentielles

### 4.1 Nommage des fichiers PDF (par défaut)

Le PDF généré **conserve l’extension d’origine** :

```
document.docx → document.docx.pdf
image.jpg     → image.jpg.pdf
email.msg     → email.msg.pdf
```

Avantages :

* traçabilité parfaite
* aucun conflit de noms
* audit facile

Une option permet de revenir à `document.pdf`, mais **ce n’est pas le comportement par défaut**.

---

### 4.2 Journal / Log (critique)

* ✅ **Activé par défaut**
* Format : **CSV**
* Emplacement : **dossier racine traité**
* Nom :

  ```
  conversion_log_YYYYMMDD_HHMMSS.csv
  ```

#### Colonnes typiques

* `timestamp`
* `status`

  * `success`
  * `skipped_pdf`
  * `skipped_password`
  * `skipped_type`
  * `error`
* `source`
* `output_pdf`
* `duration`
* `detail`

👉 Le journal est la **clé de confiance** du pipeline.

---

## 5. Word / Excel : stratégie et contraintes

### 5.1 Moteur principal

* **Microsoft Office COM**
* Utilisation obligatoire de :

  * `DispatchEx("Word.Application")` → instance dédiée
  * jamais l’instance GUI de l’utilisateur
* Paramètres :

  * `DisplayAlerts = 0`
  * export PDF via :

    * `ExportAsFixedFormat`
    * fallback `SaveAs2(FileFormat=17)`

👉 Avoir Word déjà ouvert **peut casser l’automatisation**
→ le script doit **toujours créer sa propre instance**.

---

### 5.2 Documents protégés par mot de passe (point critique)

Si un document Word / Excel est protégé :

* ❌ ne pas convertir
* ❌ ne pas tenter de fallback (LibreOffice / ReportLab)
* ❌ ne pas produire de PDF partiel
* ✅ **SKIP propre**
* ✅ journaliser `skipped_password`
* ✅ continuer le batch

Détection :

* message d’erreur contenant :

  ```
  password / mot de passe / protected / protégé / encrypt
  ```

👉 **Skip passwords est le comportement par défaut.**

---

## 6. Fallbacks autorisés / interdits

### Autorisés

* **LibreOffice** :

  * fallback acceptable si Office COM échoue
  * uniquement hors cas “password”

* **ReportLab** :

  * `.txt`, `.log`, `.xml`
  * jamais pour Word / Excel en échec

### Interdits

* ❌ ReportLab comme fallback pour Word protégé
* ❌ OCR implicite
* ❌ PDF généré malgré erreur bloquante

---

## 7. HTML

* Conversion via **Edge / Chrome headless**
* Méthode : print-to-PDF
* Objectif : rendu fidèle (CSS, tableaux, mise en page)
* ❌ pas via ReportLab

---

## 8. Images

* Conversion image → PDF simple
* Pas de recompression agressive
* Pas d’OCR
* Orientation et dimensions conservées autant que possible

---

## 9. PDF existants

* Toujours **ignorés**
* Jamais supprimés
* Jamais retraités sans OCR explicite

---

## 10. Options de ligne de commande

### Syntaxe générale

```bash
python convertir_pdf.py <repertoire> [options]
```

### Options principales

| Option              | Description                                |
| ------------------- | ------------------------------------------ |
| `<repertoire>`      | Dossier racine à traiter                   |
| `-r`, `--recursive` | Parcours récursif des sous-dossiers        |
| `--no-keep-ext`     | Désactive le nommage `x.ext.pdf` → `x.pdf` |
| `--no-journal`      | Désactive la création du journal           |
| `--delete`          | Supprime le fichier source après succès    |
| `--images-only`     | Traite uniquement les images               |
| `--word-only`       | Traite uniquement Word / Excel             |
| `--force`           | Force reconversion même si PDF existe      |

*(les options exactes peuvent varier légèrement selon la version, mais ces intentions doivent être respectées)*

---

## 11. Robustesse attendue

* `Ctrl+C` :

  * arrêt propre
  * journal fermé correctement
  * aucune instance Office laissée ouverte

* Vérification obligatoire :

```bash
python -m py_compile convertir_pdf.py
```

Aucune erreur Python (notamment **indentation**).

---

## 12. Problèmes rencontrés (à ne pas reproduire)

* erreurs d’indentation dans blocs `if`
* fallback ReportLab sur Word protégé
* faux “success” malgré PDF invalide
* réutilisation d’une instance Word déjà ouverte
* OCR lancé implicitement

---

## 13. Philosophie générale

* **Mieux vaut SKIP qu’un mauvais PDF**
* Conversion et OCR sont **deux étapes séparées**
* Traçabilité > automatisme aveugle
* Prévisibilité > magie
* Auditabilité > rapidité

---

Ce document décrit **l’état cible fonctionnel** du programme.
Tout développement ultérieur doit **respecter ces règles**, même si le code est refactoré.
