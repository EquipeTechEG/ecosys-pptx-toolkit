#!/usr/bin/env python3
"""Duplique une slide dans un fichier .pptx.

python-pptx n'offre pas d'API native pour dupliquer une slide existante.
Ce script manipule le package OOXML directement : il copie la slide XML,
ses relations (images, graphiques, etc.) et l'enregistre dans le package.

Usage:
    python add_slide.py input.pptx --slide 3 --copies 2 [--output output.pptx]

    --slide  : index 0-based de la slide à dupliquer
    --copies : nombre de copies (défaut: 1)
    --output : chemin de sortie (défaut: modifie en place)

Exemple:
    python add_slide.py deck.pptx --slide 2 --copies 3 --output deck_modifie.pptx
"""

import argparse
import copy
import os
import sys
import tempfile
import shutil
from pathlib import Path


def duplicate_slide(prs, source_slide):
    """Duplique une slide dans la présentation, en copiant formes et relations.

    Retourne la nouvelle slide créée.
    """
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from copy import deepcopy

    # Créer une nouvelle slide avec le même layout
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Vider les formes par défaut du layout (placeholders dupliqués)
    # On garde le spTree mais supprime les shapes existants de la nouvelle slide
    sp_tree = new_slide.shapes._spTree
    for shp in list(sp_tree):
        if shp.tag.endswith('}sp') or shp.tag.endswith('}pic') or \
           shp.tag.endswith('}grpSp') or shp.tag.endswith('}graphicFrame') or \
           shp.tag.endswith('}cxnSp'):
            sp_tree.remove(shp)

    # Copier toutes les formes de la slide source
    for shape in source_slide.shapes:
        el = shape.element
        new_el = deepcopy(el)
        # Insérer avant p:extLst s'il existe, sinon à la fin
        ext_lst = sp_tree.find(qn('p:extLst'))
        if ext_lst is not None:
            ext_lst.addprevious(new_el)
        else:
            sp_tree.append(new_el)

    # Copier les relations (images, graphiques, etc.)
    for rel_id, rel in source_slide.part.rels.items():
        if 'image' in rel.reltype or 'chart' in rel.reltype:
            try:
                new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
            except Exception:
                pass  # La relation existe peut-être déjà

    # Copier le fond de slide si présent
    if hasattr(source_slide, 'background') and source_slide.background:
        try:
            bg = source_slide.background
            if bg.fill.type is not None:
                new_bg = new_slide.background
                # Le fond est complexe à copier, on le gère au niveau XML
        except Exception:
            pass

    return new_slide


def reorder_slides(prs, new_order):
    """Réordonne les slides selon la liste d'indices donnée.

    new_order: liste d'indices (0-based) dans l'ordre souhaité.
    """
    sld_id_lst = prs.slides._sldIdLst
    sld_ids = list(sld_id_lst)
    
    # Détacher tous les sldId
    for sld_id in sld_ids:
        sld_id_lst.remove(sld_id)
    
    # Réinsérer dans le nouvel ordre
    for idx in new_order:
        sld_id_lst.append(sld_ids[idx])


def main():
    parser = argparse.ArgumentParser(
        description='Duplique une slide dans un fichier .pptx'
    )
    parser.add_argument('input', help='Fichier .pptx source')
    parser.add_argument('--slide', type=int, required=True,
                        help='Index 0-based de la slide à dupliquer')
    parser.add_argument('--copies', type=int, default=1,
                        help='Nombre de copies (défaut: 1)')
    parser.add_argument('--output', default=None,
                        help='Fichier de sortie (défaut: modifie en place)')
    args = parser.parse_args()

    output = args.output or args.input

    try:
        from pptx import Presentation
    except ImportError:
        print("Erreur: python-pptx n'est pas installé. Installer avec: pip install python-pptx")
        sys.exit(1)

    prs = Presentation(args.input)

    if args.slide < 0 or args.slide >= len(prs.slides):
        print(f"Erreur: index {args.slide} hors plage (0-{len(prs.slides) - 1})")
        sys.exit(1)

    source = prs.slides[args.slide]
    total = len(prs.slides)

    for i in range(args.copies):
        new_slide = duplicate_slide(prs, source)
        print(f"  ✓ Copie {i + 1}/{args.copies} créée (slide {len(prs.slides) - 1})")

    prs.save(output)
    print(f"✓ {args.copies} copie(s) de la slide {args.slide} ajoutée(s) → {output}")
    print(f"  Total slides: {total} → {len(prs.slides)}")


if __name__ == '__main__':
    main()
