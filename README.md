# ecosys-pptx-toolkit

Toolkit de generation de presentations Ecosys - scripts Python, template PPTX annote, icones.
Utilise par la skill Vibe presentation_pptx_ecosys.

## Contenu

ecosys-pptx-toolkit/
  scripts/
    aide_pptx.py          - utilitaires de remplissage (forme, ecrire, ph, motifs visuels)
    controler_deck.py     - controle du deck genere (debordements, illustrations, oublis)
    verifier_template.py  - coherence catalogue <-> template (275 controles)
    add_slide.py          - duplication de slides (python-pptx n'a pas d'API native)
    clean.py              - purge des medias orphelins apres suppressions
  references/
    TEMPLATE_ECOSYS_ANNOTE.pptx   - template maitre 81 slides - JAMAIS modifie en place
    catalogue_da_ecosys.json      - 24 types, 166 champs, regles transversales
    structure_propale_ecosys.json - analyse de 5 decks clients, banque de 25 references
  assets/
    icones/
      analyse.png, brevet.png, budget.png, delai.png, donnees.png, equipe.png
      feuille_de_route.png, livrable.png, perimetre.png, portefeuille.png, resultat.png, risque.png

## Utilisation

Ce repo est consomme par la skill Vibe presentation_pptx_ecosys. Au debut de
chaque generation, la skill clone ce repo pour recuperer les scripts, le template
et les icones.

    git clone https://github.com/EquipeTechEG/ecosys-pptx-toolkit.git

### Pre-requis

- Python 3.8+
- python-pptx (pip install python-pptx)
- LibreOffice (soffice) pour le controle visuel (optionnel)
- pdftoppm pour l'export PNG (optionnel)

### Ajouter les fichiers binaires

Les fichiers suivants ne sont pas dans ce repo (trop lourds pour l'API) -
les uploader manuellement via GitHub web UI (Add file -> Upload files) :

1. references/TEMPLATE_ECOSYS_ANNOTE.pptx - le template maitre 81 slides
2. assets/icones/*.png - les 12 icones Ecosys

### Maintenance

Apres toute retouche du template ou du catalogue :

    python scripts/verifier_template.py

Le script verifie que chaque shape_name et chaque placeholder_idx declare au
catalogue existe reellement dans la slide correspondante (275 controles).
Il doit afficher TOUT EST COHERENT.
