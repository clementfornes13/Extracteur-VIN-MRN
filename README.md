<!-- LOGO PROJET -->
<div align="center">
  <h3 align="center">Extracteur VIN MRN PDF</h3>
</div>

<!-- SOMMAIRE -->
<details>
  <summary>Sommaire</summary>
  <ol>
    <li>
      <a href="#a-propos">A propos</a>
    </li>
    <li>
      <a href="#mise-en-place">Mise en place</a>
      <ul>
        <li><a href="#prerequis">Pré-requis</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#utilisation">Utilisation</a></li>
    <lia><a href="#versionhistory">Historique des versions</a></li>
    <li><a href="#ressources">Ressources utilisés</a></li>
  </ol>
</details>

<!-- A propos -->
## A propos


L'extracteur VIN PDF est un exécutable qui permet aux utilisateurs d'extraire directement les VIN et MRN présents sur des PDF dans un dossier donné

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>


<!-- Mise en place -->
## Mise en place

### Pré-requis

* Windows
* Python
  ```sh
  https://www.python.org/downloads/
  ```
    
* PySimpleGUI
  ```sh
  pip install pysimplegui
  py -m pip install pysimplegui
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

	> Analyse du texte pour trouver le texte avant le VIN

- Version 1.1 :

	> Détection directe des VINs selon un pattern

- Version 1.2 :

  > Ajout du choix de l'emplacement de destination

  > Ouverture automatique du fichier à la fin de l'extraction

- Version 2.0 :

  > Nouvelle Interface
  
  > Détection VIN et MRN
 
  > Traite tout les fichiers dans un dossier

<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>

<!-- Ressources utilisées -->
## Ressources

* [PyPDF2 Documentation](https://pypdf2.readthedocs.io/en/3.0.0/)
* [OpenPyXL Documentation](https://openpyxl.readthedocs.io/en/stable/)
* [PySimpleGUI Documentation](https://www.pysimplegui.org/en/latest/)
<p align="right">(<a href="#readme-top">Revenir en haut</a>)</p>
