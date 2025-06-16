# **ARCHIVAGE AUTOMATIQUE DES DOSSIERS**

[![Made with Python](https://img.shields.io/badge/Made%20with-Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)

> **Application pour automatiser l'archivage des dossiers mandats en masse, suivant la procédure [PRD189](https://qmt-group.atlassian.net/wiki/spaces/PRD101/pages/2893348870/PRD189_F_SL+-+Gestion+administrative+d+une+affaire#Archivage-des-affaires)**

---

## 1. Présentation du projet :

- Ce projet a pour but de permettre l'**archivage en masse** de tous les dossiers mandats concernés (du disque "*H:*") vers le "*CohesityArchives*" (le disque "*Z:*"), de manière automatique.
- Il s'appuie sur la **procédure [PRD189](https://qmt-group.atlassian.net/wiki/spaces/PRD101/pages/2893348870/PRD189_F_SL+-+Gestion+administrative+d+une+affaire#Archivage-des-affaires)** pour l'archivage des mandats.
- Il est prévu de pouvoir l'exécuter en **ligne de commande**, ou de manière **programmée** (toutes les semaines, par exemple).

## 2. Fonctionnement :

Voici l'aroresence typique du dossier "*H:*" :

```bash
H:\
├── Dosssier affaire (ex : 1234\)
│   ├── Dossier articles (pas toujours présent) (ex : _1234-AA (Articles)\)
│   ├── Dossier mandat (ex : 1234-01\)         
│   ...
...
```

- L'application va lire les dossiers mandats dans le disque "*H:*", et va retenir ceux qui ont l'attribut "Le dossier est prêt à être archivé" (*Propriétés* > *Général* > *Avancé*). En utilisation normale, cet attribut est **à cocher par le responsable du mandat** quand celui-ci est terminé,
- Puis, ce dossier est copié dans le dossier "*CohesityArchives*" (le disque "*Z:\Affaires*"), selon une **arborescence similaire** à celle du disque "*H:*" (dans notre exemple, "*Z:\Affaires\1234\1234-01*"),
- Ensuite, l'application récupère, le cas échéant, le dossier articles de l'affaire, et **copie une archive de ce dossier** dans le sous dossier "*000_Zip articles livres*" du dossier mandats archivé (dans notre exemple, "*Z:\Affaires\1234\1234-01\000_Zip articles livres\_1234-AA (Articles).zip*"). Si le dossier articles n'existe pas, cette étape est ignorée.
- Enfin, le **dossier mandat est supprimé** du disque "*H:*".

---

## 3. Arborescence du projet

```bash
Archivage affaires\
├── ajouter_dossier_archivage.py    # Ajouter l'attribut "Le dossier est prêt à être archivé" à la sélection de mandats
├── executer_archivage.html         # Archivage autmatique des mandats
├── requirements.txt                # Dépendances Python
└── README.md                       # Ce fichier
```

---

## 4. Installation et lancement rapide de l'application

Avant de lancer le projet, **installer les dépendances**
```bash
cd Archivage affaires
pip install -r requirements.txt
```

Une fois les installations réussies, **lancer l'application**
```bash
python executer_archivage.py
```

Pour mettre à jour les mandats à archiver, cocher l'attribut "Le dossier est prêt à être archivé" manuellement, ou **lancer l'application**
```bash
python ajouter_dossier_archivage.py mandat1 mandat2 mandat3...
```

## 5. Aide sur les fonctions et les paramètres

### 5.1 Script `ajouter_dossier_archivage.py`

Ce script contient **une fonction** :
- `pret_a_archiver(liste_mandats, chemin)`: Ajoute l'attribut "Le dossier est prêt à être archivé" aux mandats passés en paramètre dans `liste_mandats` et contenus dans le dossier source `chemin`.

#### Hypothèse d'utilisation et limites:

- i. **L'utilisateur a accès au dossier source `chemin` (par défaut, "*H:*"), et peut donc lire et écrire dans ce dossier**, sinon l'application lèvera une exception au moment de sauvegarder le log.
- ii. **Les mandats passés en paramètre doivent exister dans le dossier source `chemin`, en tant que sous-dossiers d'un dossier affaire**, sinon les mandats ne seront pas trouvés par l'application, qui continuera son execution sans prendre en compte les mandats en question.
- iii. **Les mandats ont un nom strictement sous la forme X-Y** (n'importe quelle suite de caractère non nulle, puis un tiret, puis n'importe que suite de caractère) et **les dossiers affaire parents ont un nom strictement sous la forme X** (ce qu'il y a avant le tiret du mandat), sinon les mandats ne seront pas trouvés par l'application, qui continuera son execution sans prendre en compte les mandats en question.


| Hypothèse     | Exception levée | Mandats marqués comme "prêt à être archivés" |
| ------------- | --------------- | -------------------------------------------- |
| i.            | ✔️              | ❌                                          |
| ii.           | ❌              | ❌                                          |
| iii.          | ❌              | ❌                                          |

### 5.2 Script `executer_archivage.py`

Ce script contient **deux fonctions** :
- `archivage(chemin, chemin_archive)`: Archive tous les mandats marqués comme "prêt à être archivé" dans le dossier source `chemin`, et les copie dans le dossier d'archive `chemin_archive`.
- `barre_de_progression(pourcentage, curseur, longueur=100, temps_depart=None, affichage_temps_restant=False)`: Affiche une barre de progression dans la console (utile uniquement dans le terminal pour voir l'avancement de l'archivage).

#### Hypothèse d'utilisation et limites:

- i. **L'utilisateur a accès au dossier source `chemin` (par défaut, "*H:*") et au dossier d'archivage `chemin_archivage` (par défaut, "*Z\Affaires:*"), et peut donc lire `chemin` et écrire dans `chemin_archivage`**, sinon l'application lèvera une exception au moment de copier les mandats et de sauvegarder le log.
- ii. **Les mandats à archiver doivent exister dans le dossier source `chemin`, en tant que sous-dossiers d'un dossier affaire**, sinon les mandats ne seront pas trouvés par l'application, qui continuera son execution sans prendre en compte les mandats en question.
- iii. **Les dossiers articles sont de la forme _X** (un underscore, puis n'importe quelle suite de caractère), sinon les dossiers articles ne seront pas archivés et copiés dans le dossier d'archivage.


| Hypothèse     | Exception levée | Mandats lus depuis `chemin` | Mandats copiés dans `chemin_archivage` | Dossier articles copié dans `chemin_archivage` | 
| ------------- | --------------- | --------------------------- | -------------------------------------- | ---------------------------------------------- |
| i.            | ✔️              | ❌                         | ❌                                     | ❌                                            |
| ii.           | ❌              | ❌                         | ✔️                                     | ❌                                            |
| iii.          | ❌              | ✔️                         | ✔️                                     | ❌                                            |
