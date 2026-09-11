"""Verifie que chaque champ declare dans catalogue_da_ecosys.json existe bien
dans TEMPLATE_ECOSYS_ANNOTE.pptx. A relancer apres toute retouche du template
ou du catalogue.  Usage : python scripts/verifier_template.py
"""
import json
import os
from pptx import Presentation

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REFS = os.path.join(RACINE, 'references')

d = json.load(open(os.path.join(REFS, 'catalogue_da_ecosys.json'), encoding='utf-8'))
prs = Presentation(os.path.join(REFS, 'TEMPLATE_ECOSYS_ANNOTE.pptx'))

def walk(shs):
    for s in shs:
        yield s
        if s.__class__.__name__ == 'GroupShape':
            yield from walk(s.shapes)

def slides_de(p):
    """toutes les slides PPTX declarees par une entree pptx du catalogue"""
    out = []
    for k, v in p.items():
        if 'slide' not in k.lower() or k == 'slides_a_exclure':
            continue
        if isinstance(v, int):
            out.append(v)
        elif isinstance(v, dict):
            for vv in v.values():
                out += vv if isinstance(vv, list) else [vv]
    return sorted({x for x in out if isinstance(x, int)})

ko = tot = 0
for t in d['types']:
    p = t.get('pptx')
    if not p:
        print(f"KO {t['tag']}: aucun mapping pptx"); ko += 1; continue
    sl = slides_de(p)
    if not sl:
        print(f"KO {t['tag']}: aucune slide declaree"); ko += 1; continue
    par_variante = p.get('champs_par_variante') or {}
    for sn in sl:
        s = prs.slides[sn - 1]
        noms = {sh.name for sh in walk(s.shapes)}
        idxs = {ph.placeholder_format.idx for ph in s.placeholders}
        # certains champs n'existent que sur une variante precise
        attendus = None
        for cle, champs in par_variante.items():
            if cle.split('_')[0] == str(sn):
                attendus = set(champs)
        for c in p.get('champs', []):
            if attendus is not None and c.get('shape_name') and c['shape_name'] not in attendus:
                continue
            tot += 1
            if 'shape_name' in c and c['shape_name'] not in noms:
                print(f"KO {t['tag']} slide {sn}: forme absente {c['shape_name']!r}"); ko += 1
            if 'placeholder_idx' in c and c['placeholder_idx'] not in idxs:
                print(f"KO {t['tag']} slide {sn}: placeholder idx {c['placeholder_idx']} absent"); ko += 1
    print(f"ok {t['tag']:<28} slides {sl if len(sl)<=6 else str(sl[:6])+'...'} "
          f"({len(p.get('champs', []))} champs)")

print(f"\n{tot} verifications, {ko} erreur(s)" + ("  — TOUT EST COHERENT" if not ko else ""))
