<div align="center">

<h1>Extracteur VIN et MRN</h1>

</div>

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/8a867e5e4ebe4c11824e35cea688f8cf)](https://app.codacy.com/gh/clementfornes13/Extracteur-VIN-MRN?utm_source=github.com&utm_medium=referral&utm_content=clementfornes13/Extracteur-VIN-MRN&utm_campaign=Badge_Grade)

<details>

<summary>Sommaire</summary>

- [À propos](#%C3%A0-propos)
- [Mise en place](#mise-en-place)
  - [Pré-requis](#pr%C3%A9-requis)
- [Historique des versions](#historique-des-versions)
- [Roadmap](#roadmap)
- [Ressources](#ressources)

</details>

## À propos

L'extracteur VIN et MRN est un exécutable qui permet aux utilisateurs d'extraire directement les VIN et MRN présents sur des PDF dans un dossier donné

<div align="center">

[![Screenshot](https://github.com/clementfornes13/Extracteur-VIN-MRN/blob/main/images/Screenshot%20Interface.png)](https://github.com/clementfornes13/Extracteur-VIN-MRN)

</div>

---

<div align="right">

[🔼 Revenir en haut](#%C3%A0-propos)

</div>

## Mise en place

### Pré-requis

- Windows
- Python 3.11

  ```py

  https://www.python.org/downloads/
  ```

- PyPDF2

  ```py

  pip install PyPDF2
  ```

- openpyxl

  ```py

  pip install openpyxl
  ```

---

<div align="right">

[🔼 Revenir en haut](#%C3%A0-propos)

</div>

## Historique des versions

- Version 1 :
  - Extraction du texte et cherche les VINs
- Version 1.1 :
  - Définition d'un pattern pour les VINs
    - Interface graphique simple
- Version 1.2 :
  - Ajout du choix de l'emplacement de destination
  - Ouverture automatique du fichier à la fin de l'extraction
- Version 2.0 :
  - Nouvelle Interface plus complète et rapide avec PyQt6
  - Définition d'un pattern pour les MRNs
  - Détection des MRNs et VINs associés
  - Traite tout les fichiers dans un même dossier
- Version 3.0 :
  - Nouvelle Interface avec Tkinter (meilleure réactivité)
  - Choix entre VIN, MRN ou VIN + MRN avec des cases à cocher
  - Barre de progression pour avoir un aperçu de l'avancement
  - Utilisation de différents threads pour améliorer la vitesse d'exécution et l'expérience utilisateur

---

<div align="right">

[🔼 Revenir en haut](#%C3%A0-propos)

</div>

## Roadmap

✅ Extraction de VIN\
✅ Extraction de MRN\
✅ Barre de progression\
✅ Interface graphique\
✅ Utilisation de threads\
✅ Traitement par dossier\
✅ Facilité d'utilisation et d'installation

---

<div align="right">

[🔼 Revenir en haut](#%C3%A0-propos)

</div>

## Ressources

- [PyPDF2 Documentation](https://pypdf2.readthedocs.io/en/3.0.0/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/en/stable/)
- [Tkinter Documentation](https://docs.python.org/fr/3/library/tkinter.html)
- [auto-py-to-exe](https://pypi.org/project/auto-py-to-exe/)

<div align="right">

[🔼 Revenir en haut](#%C3%A0-propos)

</div>
