#Résumé
Ce petit projet sert à remplir des fichiers PPT standards avec une source de données excel. Sous forme de user story, cela s'écrit de la manière suivante :

```UserStory

En tant que Manager / chef de projet / contrôleur de gestion / responsable d'activité
Je dois utiliser des modèles de fichier powerpoint Et les remplir manuellement avec des fichiers excel standard
Pour produire des présentations types pour mon chef / patron / client / etc
Et c'est franchement inutile !
```

#Comment faire

Pour faire marcher la machine il vous faut :

1. Télécharger les fichiers .bas
2. Ouvrir un fichier powerpoint (de préférence celui que vous voulez alimenter automatiquement), ouvrir la partie Macro et les activer
3. Importer les fichiers .bas dans votre powerpoint (clique droit sur le panneau des composants, puis importer)
4. Ajouter les références manquantes (Dans le menu de l'éditeur de macro, Outils --> References):
  * Microsoft Excel component
  * Microsoft Script Control
  * Microsoft Scripting runtime
  * Microsoft Scriptlet librairy
  * Microsoft Shells control and Automation
  * Microsoft VBScript Regular Expression
  * Microsoft WMI Scripting Librairy
  * Microsoft JScript
  * Windows Script Host Object Model
  * WSHControllerLibrairy
5. Sauvegarder votre fichier sous la forme d'un pptm ou potm pour conserver les macros
6. Ajouter les commandes dans une syntaxe moustache pour expliquer d'où viennent les données
7. Lancer la macro `populateMyPPT`

#Syntaxe

Le choix de mustache pour la syntaxe provient de 1) ma connaissance du format, 2) du fait qu'il permet de repérer le format sans complexité pour le code 3) qu'il y a peut de chance que vous utilisiez des moustaches dans vos textes.

Par exemple, pour aller chercher la cellule B3 d'une feuille de calcul "S01", il faut déclarer à l'endroit où vous voulez utiliser cette donnée `{{S01!B3}}`.
Par défaut, le formatage est conservé.

#Version

##0.1

@date : 20160114

@content :

* Mustache-like formating for one cell data
* Choose your Excel file
* Save the file automatically to `filename_populated.pptm`

