#!/usr/bin/env python3
"""Wrapper pour conversion LibreOffice headless.

Convertit un fichier .pptx (ou .docx, .odp, etc.) en PDF ou autre format
via LibreOffice en mode headless. Ce script masque les complexités de
l'appel en ligne de commande et gère les chemins.

Usage:
    python soffice.py --headless --convert-to pdf input.pptx [--outdir ./]

Prérequis:
    - LibreOffice (soffice) installé et accessible dans le PATH
    - Sur macOS: /Applications/LibreOffice.app/Contents/MacOS/soffice

Options:
    --headless       : mode sans interface (requis)
    --convert-to     : format de sortie (pdf, png, jpg, html...)
    --outdir         : répertoire de sortie (défaut: répertoire courant)
    --portable       : chemin vers une installation portable de LibreOffice
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def find_soffice(portable_path=None):
    """Trouve l'exécutable LibreOffice."""
    if portable_path:
        candidate = Path(portable_path)
        if candidate.exists():
            return str(candidate)
        print(f"Erreur: chemin soffice portable introuvable: {portable_path}")
        sys.exit(1)
    
    # Chercher dans le PATH
    soffice = shutil.which('soffice')
    if soffice:
        return soffice
    
    # Chercher dans les emplacements communs
    common_paths = [
        '/usr/bin/soffice',
        '/usr/local/bin/soffice',
        '/opt/libreoffice/program/soffice',
        '/snap/bin/libreoffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return None


def convert(input_file, output_format, outdir=None, portable_path=None):
    """Convertit le fichier via LibreOffice headless."""
    input_path = Path(input_file).resolve()
    if not input_path.exists():
        print(f"Erreur: fichier introuvable: {input_path}")
        sys.exit(1)
    
    outdir = outdir or str(input_path.parent)
    outdir_path = Path(outdir).resolve()
    outdir_path.mkdir(parents=True, exist_ok=True)
    
    soffice = find_soffice(portable_path)
    if not soffice:
        print("Erreur: LibreOffice (soffice) introuvable.")
        print("Installer LibreOffice ou utiliser --portable pour spécifier le chemin.")
        sys.exit(1)
    
    # Construire la commande
    cmd = [
        soffice,
        '--headless',
        '--norestore',
        '--nologo',
        '--convert-to', output_format,
        '--outdir', str(outdir_path),
        str(input_path),
    ]
    
    print(f"Conversion: {input_path.name} → {output_format}")
    print(f"Sortie: {outdir_path}")
    print(f"Commande: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            # Vérifier que le fichier de sortie existe
            output_name = input_path.stem + '.' + output_format
            # LibreOffice peut produire des extensions légèrement différentes
            expected = outdir_path / output_name
            if expected.exists():
                print(f"✓ Conversion réussie: {expected}")
                return str(expected)
            else:
                # Chercher le fichier de sortie
                candidates = list(outdir_path.glob(f"{input_path.stem}*.{output_format}*"))
                if candidates:
                    print(f"✓ Conversion réussie: {candidates[0]}")
                    return str(candidates[0])
                else:
                    print(f"⚠ Conversion terminée mais fichier de sortie introuvable")
                    if result.stdout:
                        print(f"  stdout: {result.stdout}")
                    return None
        else:
            print(f"❌ Échec de la conversion (code {result.returncode})")
            if result.stderr:
                print(f"  stderr: {result.stderr}")
            if result.stdout:
                print(f"  stdout: {result.stdout}")
            sys.exit(1)
    except subprocess.TimeoutExpired:
        print("❌ Timeout: la conversion a dépassé 120 secondes")
        sys.exit(1)
    except FileNotFoundError:
        print(f"Erreur: exécutable introuvable: {soffice}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description='Wrapper pour conversion LibreOffice headless'
    )
    parser.add_argument('--headless', action='store_true',
                        help='Mode sans interface (requis, ignoré — toujours headless)')
    parser.add_argument('--convert-to', required=True,
                        help='Format de sortie (pdf, png, jpg, html...)')
    parser.add_argument('--outdir', default=None,
                        help='Répertoire de sortie')
    parser.add_argument('--portable', default=None,
                        help='Chemin vers une installation portable de LibreOffice')
    parser.add_argument('input', help='Fichier à convertir')
    args = parser.parse_args()
    
    convert(args.input, args.convert_to, args.outdir, args.portable)


if __name__ == '__main__':
    main()
