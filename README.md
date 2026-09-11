# Ecosys PPTX Toolkit

Toolkit de génération de présentations Ecosys au format PowerPoint. Utilisé par la skill Vibe `presentation-pptx-ecosys`.

## Structure

```
ecosys-pptx-toolkit/
  scripts/
    aide_pptx.py          # Utilitaires de remplissage (forme, ecrire, ph, motifs visuels...)
    controler_deck.py     # Contrôle du deck généré (débordements, illustrations, oublis)
    verifier_template.py  # Cohérence catalogue ↔ template (275 contrôles)
    add_slide.py          # Duplication de slides (python-pptx n'a pas d'API native)
    clean.py              # Purge des médias orphelins après suppressions
    validate.py           # Validation du deck contre le template original
    soffice.py            # Wrapper LibreOffice headless (conversion PDF)
  references/
    TEMPLATE_ECOSYS_ANNOTE.pptx   # Template maître 81 slides — JAMAIS modifié en place
    catalogue_da_ecosys.json      # 24 types, 166 champs, règles transversales
    structure_propale_ecosys.json # Analyse de 5 decks clients, banque de 25 références
  assets/
    icones/                         # 12 icônes Ecosys au format SVG
      analyse.svg  brevet.svg  budget.svg  delai.svg  donnees.svg  equipe.svg
      feuille_de_route.svg  livrable.svg  perimetre.svg  portefeuille.svg  resultat.svg  risque.svg
```

## Installation

### Prérequis

- Python 3.8+
- `python-pptx` : `pip install python-pptx`
- LibreOffice (pour le contrôle visuel via `soffice.py`)
- `pdftoppm` (pour les aperçus PNG, optionnel)

### Cloner le repo

```bash
git clone https://github.com/EquipeTechEG/ecosys-pptx-toolkit.git
cd ecosys-pptx-toolkit
pip install python-pptx
```

### Template PowerPoint

Le fichier `references/TEMPLATE_ECOSYS_ANNOTE.pptx` étant binaire, il doit être
téléversé manuellement via l'interface GitHub :

1. Aller sur https://github.com/EquipeTechEG/ecosys-pptx-toolkit/upload/main
2. Glisser le fichier `TEMPLATE_ECOSYS_ANNOTE.pptx` dans le dossier `references/`

## Scripts

### aide_pptx.py
Bibliothèque d'utilitaires pour le remplissage des slides. Fournit :
- `forme(slide, nom)` — retrouve une forme par son nom (Groupes inclus)
- `ecrire(shape, lignes)` — écrit en préservant la mise en forme
- `ph(slide, idx, lignes)` — écrit sur un placeholder natif
- `remplir_sommaire(slide, chapitres)` — remplit le sommaire
- `appliquer_degrade_titres(prs)` — dégradé vert → bleu sur les titres
- `bandeau_resultats_cles(slide, entrees)` — bandeau navy à tuiles chiffrées
- `pipeline_numerote(slide, etapes)` — cartes numérotées avec flèches
- `couches_empilees(slide, couches)` — bandes horizontales empilées
- `trois_icones(slide, volets)` — pastilles navy à icônes Ecosys
- `rotation_variantes(slides_par_theme, nombre, theme)` — rotation des variantes

### controler_deck.py
```bash
python scripts/controler_deck.py sortie.pptx
```
Signale les textes qui débordent, les illustrations non remplacées et les invites oubliées.

### verifier_template.py
```bash
python scripts/verifier_template.py
```
Vérifie la cohérence entre le catalogue JSON et le template PPTX (275 contrôles).

### add_slide.py
```bash
python scripts/add_slide.py deck.pptx --slide 2 --copies 3 --output deck_modifie.pptx
```
Duplique une slide (python-pptx n'a pas d'API native pour cela).

### clean.py
```bash
python scripts/clean.py deck.pptx
```
Purge les médias orphelins après suppression de slides.

### validate.py
```bash
python scripts/validate.py deck.pptx --original references/TEMPLATE_ECOSYS_ANNOTE.pptx
```
Valide le deck généré contre le template source.

### soffice.py
```bash
python scripts/soffice.py --headless --convert-to pdf deck.pptx --outdir ./
```
Wrapper pour conversion LibreOffice headless.

## Icônes Ecosys

12 pictos au trait, blanc pur, conçus pour un fond sombre (pastille navy du gabarit) :

analyse · brevet · budget · délai · données · équipe · feuille_de_route · livrable · périmètre · portefeuille · résultat · risque

**Note** : Les icônes sont au format SVG. La fonction `chemin_icone()` de `aide_pptx.py`
construit le chemin vers `.png` par défaut. Convertir les SVG en PNG si nécessaire.

## Utilisation avec la skill Vibe

La skill `presentation-pptx-ecosys` est installée dans `skills/presentation-pptx-ecosys/`.
Elle contient le `SKILL.md` (instructions) et les catalogues JSON pour la planification.
Ce repo fournit les scripts et le template pour l'exécution.

Workflow :
1. L'agent lit le `SKILL.md` et les catalogues JSON
2. Clone ce repo pour récupérer scripts et template
3. Génère le deck en exécutant les scripts Python
4. Contrôle et livre le fichier .pptx

## Licence

Usage interne Ecosys Group / EG-EBT.
