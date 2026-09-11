#!/usr/bin/env python3
"""Purge les médias orphelins d'un fichier .pptx.

Après suppression de slides, certaines images et médias peuvent rester
dans le package sans être référencés par aucune slide. Ce script les
identifie et les supprime pour réduire la taille du fichier et éviter
les problèmes de corruption.

Usage:
    python clean.py input.pptx [--output output.pptx]

    --output : chemin de sortie (défaut: modifie en place)
"""

import argparse
import os
import sys
import shutil
import tempfile
import zipfile
from pathlib import Path


def find_referenced_media(pptx_path):
    """Retourne l'ensemble des noms de fichiers média référencés par les slides."""
    referenced = set()
    
    with tempfile.TemporaryDirectory() as tmp:
        extract_dir = Path(tmp) / 'pptx'
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            zf.extractall(extract_dir)
        
        slides_dir = extract_dir / 'ppt' / 'slides'
        rels_dir = slides_dir / '_rels'
        
        if not slides_dir.exists():
            return referenced
        
        # Parcourir les fichiers de relations de chaque slide
        for rels_file in (slides_dir / '_rels').glob('*.rels'):
            with open(rels_file, 'r', encoding='utf-8') as f:
                content = f.read()
                # Chercher les références aux médias
                for line in content.split():
                    if '../media/' in line or 'media/' in line:
                        # Extraire le nom du fichier
                        part = line.split('media/')[-1].rstrip('"').rstrip("'")
                        if part:
                            referenced.add(part)
        
        # Aussi vérifier les slideLayouts et slideMasters
        for subdir in ['slideLayouts', 'slideMasters']:
            layout_rels = extract_dir / 'ppt' / subdir / '_rels'
            if layout_rels.exists():
                for rels_file in layout_rels.glob('*.rels'):
                    with open(rels_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        for line in content.split():
                            if '../media/' in line or 'media/' in line:
                                part = line.split('media/')[-1].rstrip('"').rstrip("'")
                                if part:
                                    referenced.add(part)
    
    return referenced


def clean_orphaned_media(pptx_path, output_path=None):
    """Supprime les médias non référencés du fichier .pptx."""
    output_path = output_path or pptx_path
    
    with tempfile.TemporaryDirectory() as tmp:
        extract_dir = Path(tmp) / 'pptx'
        extract_dir.mkdir(parents=True)
        
        # Extraire le pptx
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            zf.extractall(extract_dir)
        
        media_dir = extract_dir / 'ppt' / 'media'
        if not media_dir.exists():
            print("✓ Aucun dossier média trouvé — rien à nettoyer")
            # Recompresser tel quel
            _recompress(extract_dir, output_path)
            return 0
        
        # Trouver tous les fichiers média
        all_media = set(f.name for f in media_dir.iterdir() if f.is_file())
        
        # Trouver les médias référencés
        referenced = find_referenced_media(pptx_path)
        
        # Les orphelins = tous - référencés
        orphaned = all_media - referenced
        
        if not orphaned:
            print("✓ Aucun média orphelin trouvé")
            _recompress(extract_dir, output_path)
            return 0
        
        # Supprimer les orphelins
        removed = 0
        for name in orphaned:
            media_file = media_dir / name
            if media_file.exists():
                media_file.unlink()
                removed += 1
        
        # Nettoyer aussi les Content_Types et les relations orphelines
        _clean_content_types(extract_dir, orphaned)
        
        print(f"✓ {removed} média(s) orphelin(s) supprimé(s):")
        for name in sorted(orphaned):
            print(f"  - {name}")
        
        # Recompresser
        _recompress(extract_dir, output_path)
        return removed


def _clean_content_types(extract_dir, orphaned):
    """Nettoie les références aux médias supprimés dans [Content_Types].xml."""
    content_types = extract_dir / '[Content_Types].xml'
    if not content_types.exists():
        return
    
    with open(content_types, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Pas besoin de modifier Content_Types pour les médias individuels,
    # ils partagent tous le même type d'extension
    pass


def _recompress(extract_dir, output_path):
    """Recompresse le dossier extrait en .pptx."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(extract_dir)
                zf.write(file_path, arcname)


def main():
    parser = argparse.ArgumentParser(
        description='Purge les médias orphelins d un fichier .pptx'
    )
    parser.add_argument('input', help='Fichier .pptx à nettoyer')
    parser.add_argument('--output', default=None,
                        help='Fichier de sortie (défaut: modifie en place)')
    args = parser.parse_args()
    
    output = args.output or args.input
    
    if not os.path.exists(args.input):
        print(f"Erreur: fichier introuvable: {args.input}")
        sys.exit(1)
    
    # Si on modifie en place, passer par un temporaire
    if output == args.input:
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp_path = tmp.name
        clean_orphaned_media(args.input, tmp_path)
        shutil.move(tmp_path, args.input)
    else:
        clean_orphaned_media(args.input, output)


if __name__ == '__main__':
    main()
