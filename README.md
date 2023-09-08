<!-- LOGO PROJET -->
<div align="center">
  <h3 align="center">Extracteur VIN et MRN</h3>
</div>

<!-- SOMMAIRE -->
<details>
  <summary>Sommaire</summary>
  <ol>
    <li>
      <a href="#a-propos">À propos</a>
    </li>
    <li>
      <a href="#mise-en-place">Mise en place</a>
      <ul>
        <li><a href="#prerequis">Pré-requis</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#utilisation">Utilisation</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <lia><a href="#versionhistory">Historique des versions</a></li>
    <li><a href="#ressources">Ressources utilisés</a></li>
  </ol>
</details>

<!-- À propos -->
## À propos


L'extracteur VIN et MRN est un exécutable qui permet aux utilisateurs d'extraire directement les VIN et MRN présents sur des PDF dans un dossier donné

<p align="center">
  <img src="https://github.com/clementfornes13/Extracteur-VIN-MRN/blob/main/images/Screenshot%20Interface.png" alt="Screenshot" />
</p>

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>


<!-- Mise en place -->
## Mise en place

### Pré-requis

* Windows

* Python 3.11
  ```sh
  https://www.python.org/downloads/
  ```
  
* PyPDF2
  ```sh
  pip install PyPDF2
  py -m pip install PyPDF2
  ```

* openpyxl
  ```sh
  pip install openpyxl
  py -m pip install openpyxl
  ```

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>

<!-- Historique des versions -->
## Historique des versions

- Version 1 : 

	> Extraction du texte et cherche les VINs

- Version 1.1 :

	> Définition d'un pattern pour les VINs
  > Interface graphique simple

- Version 1.2 :

  > Ajout du choix de l'emplacement de destination

  > Ouverture automatique du fichier à la fin de l'extraction

- Version 2.0 :

  > Nouvelle Interface plus complète et rapide avec PyQt6
  
  > Définition d'un pattern pour les MRNs

  > Détection des MRNs et VINs associés
 
  > Traite tout les fichiers dans un même dossier

- Version 3.0 :

  > Nouvelle Interface avec Tkinter (meilleure réactivité)

  > Choix entre VIN, MRN ou VIN + MRN avec des cases à cocher

  > Barre de progression pour avoir un aperçu de l'avancement

  > Utilisation de différents threads pour améliorer la vitesse d'exécution et l'expérience utilisateur

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>

<!-- ROADMAP -->
## Roadmap

- [x] Extraction de VIN
- [x] Extraction de MRN
- [x] Barre de progression
- [x] Interface graphique
- [x] Utilisation de threads
- [x] Traitement par dossier
- [x] Facilité d'utilisation et d'installation

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>

<!-- Ressources utilisées -->
## Ressources

* [PyPDF2 Documentation](https://pypdf2.readthedocs.io/en/3.0.0/)
* [OpenPyXL Documentation](https://openpyxl.readthedocs.io/en/stable/)
* [Tkinter Documentation](https://docs.python.org/fr/3/library/tkinter.html)
* [auto-py-to-exe](https://pypi.org/project/auto-py-to-exe/)

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>
