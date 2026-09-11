#!/usr/bin/env python3
"""Controle un deck genere et signale les defauts que l'oeil rate.

    python scripts/controler_deck.py sortie.pptx

Trois familles de defauts, constatees en usage reel :

  1. TEXTE QUI DEBORDE — la boite du gabarit est dimensionnee pour l'invite du
     template, pas pour le contenu reel. On estime la hauteur du texte et on la
     compare a celle de la boite.
  2. ILLUSTRATION NON REMPLACEE — emplacement d'image vide, ou photo du template
     laissee telle quelle alors qu'elle n'a aucun rapport avec le contenu.
  3. IMAGE REPETEE — un meme visuel reutilise sur plusieurs slides. Une regle
     documentee mais non verifiee finit oubliee : c'est arrive, un collegue a
     genere un deck ou toutes les slides portaient la meme photo.
  4. TEXTE DU TEMPLATE OUBLIE — invites du gabarit (Lorem ipsum, 'Texte',
     'Prenom NOM', 'XXX'...) restees dans le deck livre.

Le script ne corrige rien : il liste, pour que la correction soit une decision.
"""
import re
import sys

from pptx import Presentation
from pptx.util import Emu

# invites du template : leur presence dans un deck livre est une erreur
INVITES = re.compile(
    r'^(texte|titre|sous-titre|section|kpi|lorem ipsum.*|pr[ée]nom nom|'
    r'annexe n°_?|ajouter du texte|objectif \d|livrable|illustration|'
    r'x+\s*x*\s*€?|xxx mois|date|logo client|mois \d|lot \d|'
    r'[ée]tape / lot / chantier|chantier \d|paris)$',
    re.IGNORECASE)

# largeur moyenne d'un caractere, en fraction de la taille de police, et
# hauteur de ligne : approximations calibrees sur Montserrat/Calibri
RATIO_LARGEUR = 0.52
RATIO_INTERLIGNE = 1.22

# Au-dela de ce facteur seulement, le debordement est reel. Un seuil serre
# genere surtout du bruit : beaucoup de boites du gabarit debordent de quelques
# centiemes sans consequence visible, et l'estimation reste une approximation.
SEUIL_DEBORDEMENT = 1.35

# decor du gabarit : ces formes debordent par construction, les signaler noie
# les vrais defauts
CHROME = ('pied de page', 'numéro de diapositive', 'numero de diapositive',
          'slide number', 'footer')


def walk(shapes, echelle=(1.0, 1.0)):
    """Parcourt les formes en propageant l'echelle des Groupes.

    ATTENTION : les formes imbriquees dans un Groupe ont leurs coordonnees dans
    l'espace INTERNE du groupe (chOff/chExt), pas en pouces de la slide. Comparer
    directement leur largeur a des pouces donne des mesures fausses — c'est ce
    qui produisait de faux debordements.
    """
    for sh in shapes:
        yield sh, echelle
        if sh.__class__.__name__ == 'GroupShape':
            el = sh._element
            chExt = el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}chExt')
            fx, fy = echelle
            if chExt is not None:
                cx, cy = int(chExt.get('cx')), int(chExt.get('cy'))
                if cx and cy:
                    fx, fy = fx * sh.width / cx, fy * sh.height / cy
            yield from walk(sh.shapes, (fx, fy))


def taille_police(paragraphe, defaut=18.0):
    if paragraphe.font.size is not None:
        return paragraphe.font.size.pt
    for r in paragraphe.runs:
        if r.font.size is not None:
            return r.font.size.pt
    return defaut


def est_chrome(shape):
    nom = shape.name.lower()
    return any(c in nom for c in CHROME)


def texte_ajustable(shape):
    """La forme reduit-elle sa police pour tenir, ou laisse-t-elle deborder ?"""
    txBody = shape.text_frame._txBody
    return txBody.find(
        '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr/'
        '{http://schemas.openxmlformats.org/drawingml/2006/main}normAutofit') is not None


def hauteur_estimee(shape, echelle=(1.0, 1.0)):
    """Hauteur qu'occupera le texte, en pouces de la slide."""
    largeur_po = Emu(shape.width).inches * echelle[0]
    if largeur_po <= 0:
        return 0.0
    # sans retour a la ligne, le texte s'etale horizontalement : une seule ligne
    if shape.text_frame.word_wrap is False:
        pts = [taille_police(p) for p in shape.text_frame.paragraphs] or [18.0]
        return sum(pt * RATIO_INTERLIGNE / 72 for pt in pts)
    total = 0.0
    for p in shape.text_frame.paragraphs:
        texte = p.text
        pt = taille_police(p)
        if not texte:
            total += pt * RATIO_INTERLIGNE / 72
            continue
        # nombre de caracteres tenant sur une ligne, puis nombre de lignes
        par_ligne = max(int(largeur_po * 72 / (pt * RATIO_LARGEUR)), 1)
        lignes = max(-(-len(texte) // par_ligne), 1)
        total += lignes * pt * RATIO_INTERLIGNE / 72
    return total


# Les photos de couverture, de sommaire et d'intercalaire appartiennent au
# LAYOUT, pas a la slide : elles sont donc invisibles a une detection par image.
# Ce qui se repete, c'est le layout reutilise. Ces familles existent en variantes
# sectorielles — reutiliser la meme fait revenir la meme photo.
FAMILLES_A_PHOTO = ('Page de garde', 'Page de titre', 'Sommaire', 'Page intercalaire')
REPETITIONS_TOLEREES = 2
SURFACE_MINI_PHOTO = 2.0        # pouces carres, pour ecarter logos et icones


def visuels_repetes(prs):
    """Layouts a photo reutilises, et images posees en double sur les slides."""
    import collections
    layouts = collections.defaultdict(list)
    images = collections.defaultdict(list)

    for n, slide in enumerate(prs.slides, 1):
        nom = slide.slide_layout.name
        if nom.startswith(FAMILLES_A_PHOTO):
            layouts[nom].append(n)
        vus = set()
        for sh, echelle in walk(slide.shapes):
            if sh.shape_type != 13:
                continue
            try:
                surface = (Emu(sh.width).inches * echelle[0]
                           * Emu(sh.height).inches * echelle[1])
                cle = sh.image.sha1
            except Exception:
                continue
            if surface >= SURFACE_MINI_PHOTO and cle not in vus:
                vus.add(cle)
                images[cle].append(n)

    repetitions = []
    for nom, slides in layouts.items():
        if len(slides) > REPETITIONS_TOLEREES:
            repetitions.append((f'layout {nom}', slides))
    for _, slides in images.items():
        if len(slides) > REPETITIONS_TOLEREES:
            repetitions.append(('même image', slides))
    return repetitions


def controler(chemin):
    prs = Presentation(chemin)
    debordements, illustrations, oublis = [], [], []
    repetitions = visuels_repetes(prs)

    for n, slide in enumerate(prs.slides, 1):
        for sh, echelle in walk(slide.shapes):
            # --- 2) illustrations
            if sh.shape_type == 13:            # PICTURE
                if sh.name.startswith('visuel_') or 'placeholder' in sh.name.lower():
                    pass                        # image nommee : intentionnelle
            if sh.is_placeholder:
                try:
                    type_ph = sh.placeholder_format.type
                except ValueError:
                    type_ph = None
                # un emplacement d'image vide s'affiche comme un cadre gris
                if type_ph is not None and 'PICTURE' in str(type_ph):
                    illustrations.append(
                        (n, sh.name, 'emplacement image vide — le remplir ou supprimer la forme'))

            if not sh.has_text_frame:
                continue
            texte = sh.text_frame.text.strip()

            # --- 3) invites du template oubliees
            if texte and not est_chrome(sh) and INVITES.match(texte):
                oublis.append((n, sh.name, texte[:44]))

            # --- 1) debordement
            if texte and sh.height and not est_chrome(sh) and not texte_ajustable(sh):
                estimee = hauteur_estimee(sh, echelle)
                dispo = Emu(sh.height).inches * echelle[1]
                if estimee > dispo * SEUIL_DEBORDEMENT:
                    debordements.append(
                        (n, sh.name, round(estimee, 2), round(dispo, 2), texte[:40]))

    return debordements, illustrations, oublis, repetitions


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    debordements, illustrations, oublis, repetitions = controler(sys.argv[1])

    if debordements:
        print(f'\n⚠ TEXTE QUI DEBORDE ({len(debordements)})')
        for n, nom, est, dispo, extrait in debordements:
            print(f'  slide {n:>2}  {nom:<32} {est}" pour {dispo}" dispo — {extrait!r}')
    if illustrations:
        print(f'\n⚠ ILLUSTRATIONS ({len(illustrations)})')
        for n, nom, motif in illustrations:
            print(f'  slide {n:>2}  {nom:<32} {motif}')
    if repetitions:
        print(f'\n⚠ IMAGE REPETEE ({len(repetitions)})')
        for quoi, slides in repetitions:
            apercu = ', '.join(str(x) for x in slides[:8]) + ('...' if len(slides) > 8 else '')
            print(f'  {quoi} réutilisé sur {len(slides)} slides : {apercu}')
            print('    -> faire tourner les variantes avec rotation_variantes() '
                  'de scripts/aide_pptx.py')
    if oublis:
        print(f'\n⚠ TEXTE DU TEMPLATE OUBLIE ({len(oublis)})')
        for n, nom, extrait in oublis:
            print(f'  slide {n:>2}  {nom:<32} {extrait!r}')

    total = len(debordements) + len(illustrations) + len(oublis) + len(repetitions)
    print(f'\n{total} point(s) a corriger' if total else '\nAucun defaut detecte')
    return 1 if total else 0


if __name__ == '__main__':
    sys.exit(main())
