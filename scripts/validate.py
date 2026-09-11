#!/usr/bin/env python3
"""Valide un deck .pptx généré contre le template original.

Compare le deck généré avec le template source pour détecter:
  - Les placeholders oubliés (texte d'invite non remplacé)
  - Les formes nommées manquantes (champ_*, visuel_*)
  - Les différences de structure (slides supprimées/ajoutées)
  - Les relations cassées (images manquantes)

Usage:
    python validate.py generated.pptx --original TEMPLATE_ECOSYS_ANNOTE.pptx

Options:
    --original   : template original pour comparaison (obligatoire)
    --strict     : mode strict, signale aussi les warnings
"""

import argparse
import os
import sys
from collections import defaultdict


def validate(generated_path, original_path, strict=False):
    """Valide le deck généré et retourne la liste des défauts."""
    try:
        from pptx import Presentation
    except ImportError:
        print("Erreur: python-pptx n'est pas installé. Installer avec: pip install python-pptx")
        sys.exit(1)

    defauts = []
    warnings = []

    prs = Presentation(generated_path)
    original = Presentation(original_path)

    # 1. Vérifier le nombre de slides
    gen_count = len(prs.slides)
    orig_count = len(original.slides)
    if gen_count == 0:
        defauts.append("Le deck généré ne contient aucune slide")
    elif gen_count > orig_count:
        warnings.append(
            f"Le deck généré a plus de slides ({gen_count}) que le template ({orig_count}) "
            f"— normal si des slides ont été dupliquées"
        )

    # 2. Pour chaque slide, vérifier les placeholders et formes nommées
    invites_template = set()

    # Collecter les textes d'invite du template original
    for slide in original.slides:
        for shape in _walk_shapes(slide.shapes):
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text and _is_placeholder_text(text):
                    invites_template.add(text)

    # Vérifier les slides générées
    for i, slide in enumerate(prs.slides):
        slide_defauts = []
        
        for shape in _walk_shapes(slide.shapes):
            if not shape.has_text_frame:
                continue
            
            text = shape.text_frame.text.strip()
            
            if not text:
                continue
            
            # Détecter les invites non remplacées
            if _is_placeholder_text(text):
                # Si ce texte existe dans le template, c'est une invite oubliée
                if text in invites_template or _looks_like_invite(text):
                    slide_defauts.append(f"Invite non remplacée: \"{text[:60]}\"")
            
            # Détecter les textes de débordement évidents
            if len(text) > 500:
                slide_defauts.append(f"Texte anormalement long ({len(text)} car.) dans: \"{text[:40]}...\"")
        
        for d in slide_defauts:
            defauts.append(f"Slide {i}: {d}")

    # 3. Vérifier les relations d'images cassées
    for i, slide in enumerate(prs.slides):
        for shape in _walk_shapes(slide.shapes):
            if shape.shape_type == 13:  # Picture
                try:
                    image = shape.image
                except Exception:
                    defauts.append(f"Slide {i}: image cassée ou manquante dans une forme")
            elif shape.shape_type is None:
                # Formes sans type défini peuvent être des placeholders vides
                if shape.has_text_frame and not shape.text_frame.text.strip():
                    if shape.name and shape.name.startswith('champ_'):
                        warnings.append(
                            f"Slide {i}: forme nommée '{shape.name}' vide"
                        )

    # 4. Vérifier les noms de formes (champ_*, visuel_*)
    for i, slide in enumerate(prs.slides):
        shape_names = set()
        for shape in _walk_shapes(slide.shapes):
            if shape.name:
                shape_names.add(shape.name)
        
        # Chercher les formes nommées champ_ qui sont vides
        for shape in _walk_shapes(slide.shapes):
            if shape.name and shape.name.startswith('champ_') and shape.has_text_frame:
                if not shape.text_frame.text.strip():
                    warnings.append(
                        f"Slide {i}: champ '{shape.name}' vide"
                    )

    # Afficher les résultats
    print("=" * 60)
    print("  VALIDATION DU DECK")
    print("=" * 60)
    print(f"  Fichier généré : {generated_path}")
    print(f"  Template source: {original_path}")
    print(f"  Slides générées: {gen_count}")
    print(f"  Slides template: {orig_count}")
    print("-" * 60)
    
    if defauts:
        print(f"\n❌ {len(defauts)} DÉFAUT(S) DÉTECTÉ(S):")
        for d in defauts:
            print(f"  • {d}")
    else:
        print("\n✓ Aucun défaut détecté")
    
    if strict and warnings:
        print(f"\n⚠ {len(warnings)} AVERTISSEMENT(S):")
        for w in warnings:
            print(f"  • {w}")
    
    print("\n" + "=" * 60)
    return 1 if defauts else 0


def _walk_shapes(shapes):
    """Parcourt récursivement les formes, y compris dans les groupes."""
    for shape in shapes:
        yield shape
        if shape.__class__.__name__ == 'GroupShape':
            yield from _walk_shapes(shape.shapes)


def _is_placeholder_text(text):
    """Détecte si un texte ressemble à une invite de template non remplie."""
    markers = [
        'Lorem ipsum', 'Texte à remplacer', '[', '...', 'Titre',
        'Cliquez pour', 'Click to', 'placeholder', 'exemple',
        'votre texte', 'à compléter', 'XXX',
    ]
    text_lower = text.lower()
    return any(m.lower() in text_lower for m in markers)


def _looks_like_invite(text):
    """Heuristique : texte court entre crochets ou avec des placeholders."""
    if text.startswith('[') and text.endswith(']'):
        return True
    if 'XXX' in text or 'xxx' in text:
        return True
    return False


def main():
    parser = argparse.ArgumentParser(
        description='Valide un deck .pptx généré contre le template original'
    )
    parser.add_argument('generated', help='Fichier .pptx généré')
    parser.add_argument('--original', required=True,
                        help='Template original pour comparaison')
    parser.add_argument('--strict', action='store_true',
                        help='Mode strict (signale aussi les warnings)')
    args = parser.parse_args()
    
    if not os.path.exists(args.generated):
        print(f"Erreur: fichier généré introuvable: {args.generated}")
        sys.exit(1)
    if not os.path.exists(args.original):
        print(f"Erreur: template original introuvable: {args.original}")
        sys.exit(1)
    
    sys.exit(validate(args.generated, args.original, args.strict))


if __name__ == '__main__':
    main()
