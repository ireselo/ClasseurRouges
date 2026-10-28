# Classeurs Rouges — Suivi pluriannuel des élèves (Cycle d'Orientation)

Application locale (**100 % hors-ligne et confidentielle**) permettant de stocker et d'analyser le suivi scolaire longitudinal des élèves durant leurs 3 années au Cycle d'Orientation (**9CO • 10CO • 11CO**) et de générer des vues synthétiques ainsi que des rapports Excel (`.xlsx`) complets par classe/cohorte ou par élève.

## Fonctionnalités principales

- **Importation multiformat avec enregistrement automatique :**
  - **Fichiers XML (`Notes*.xml` et `Absences*.xml`) :** importe automatiquement les notes trimestrielles (`T1`, `T2`, `T3`), calcule les moyennes annuelles et les moyennes des disciplines principales (en respectant les notes indicatives `Compte="0"` affichées entre parenthèses sans les inclure dans le calcul des moyennes), et enregistre les absences par trimestre et au total. Les données XML sont fusionnées et enregistrées automatiquement dans la base de données lors de l'importation.
  - **Liste des professeurs (`*MaitresParGroupe.xls`, `.xlsx`, `.csv`) :** associe automatiquement les professeurs de chaque discipline et le/la *Maître-sse de classe* à chaque cours, classe et année scolaire. Peut être importée simultanément avec les fichiers XML ou ultérieurement pour mettre à jour les dossiers déjà enregistrés.
  - **Bulletins scolaires en PDF (`*.pdf`) :** extraction locale des bulletins PDF avec écran de vérification visuelle côte à côte avant sauvegarde.
- **Affichage des professeurs dans les fiches individuelles :**
  - Le nom du/des professeur(s) s'affiche sous chaque discipline dans les **Fiches Individuelles des Élèves** (aussi bien dans l'interface que dans les onglets individuels des rapports Excel exportés), tout en gardant la **Vue Classe / Cohorte** claire et synthétique.
- **Filtrage par Année scolaire et par Classe :**
  - Permet de filtrer par **Année scolaire** (`Année scolaire`) et par **Classe** dans la vue de cohorte ainsi que dans les fiches individuelles des élèves.
  - La sélection d'une année scolaire affiche les classes correspondantes et les élèves ayant fréquenté cette classe durant l'année choisie, tout en conservant l'historique complet des 3 années (`9CO`, `10CO`, `11CO`) et la détection des redoublements.
- **Navigation fluide avec barre de défilement horizontal toujours visible :**
  - Les tableaux de suivi disposent d'une barre de défilement horizontal persistante (*sticky*) fixée en bas de la zone visible de l'écran, permettant de faire défiler les colonnes horizontalement à tout moment sans devoir descendre jusqu'en bas de la page.
- **Exportation Excel (`.xlsx`) unitaire ou par lot (1 fichier Excel par classe) :**
  - **Rapport Excel de la sélection :** génère un classeur Excel multi-feuilles comprenant la synthèse pluriannuelle de la cohorte, le détail des notes pour `9CO`, `10CO` et `11CO`, ainsi qu'un onglet individuel formaté pour chaque élève (disciplines, professeurs, moyennes, absences, situation de promotion et orientation).
  - **Génération de tous les rapports Excel par classe d'une année scolaire :** un bouton dédié (**`📚 Générer 1 Excel par Classe`**) permet d'exporter en une seule opération tous les rapports Excel d'une année scolaire, à raison d'un fichier `.xlsx` complet par classe dans le dossier choisi.
- **Clôture de fin d'année scolaire (tri des élèves de 11CO) :**
  - Le bouton **`🎓 Fin d'année (Élèves 11CO)`** ouvre une fenêtre permettant de sélectionner les élèves qui redoublent en `11CO` (avec sélection automatique possible des élèves « non promus ») afin de conserver leur dossier et de supprimer les élèves de `11CO` sortants qui quittent le Cycle d'Orientation, tout en gardant intacts les élèves de `9CO` et `10CO` (avec sauvegarde automatique préalable dans `data/backups/`).

## Structure du projet

- `Lancer_Classeurs_Rouges.bat` : lanceur de l'application (démarre le serveur local PowerShell et ouvre l'interface dans Microsoft Edge en mode application).
- `src/` : code source de l'application :
  - `Server.ps1` : serveur HTTP local (`127.0.0.1`) et routes de l'API.
  - `Engine.ps1` : moteur d'extraction et de traitement (XML, Excel, CSV, PDF/OCR), résolution des codes de cours du CO, calcul des moyennes, gestion de la base de données JSON et génération des rapports Excel.
  - `ui.html` : interface graphique de l'application.
  - `NativePdfExtractor.cs` : extracteur natif en C# pour les bulletins PDF numériques.
  - `XlsxWriter.cs` : générateur natif en C# de classeurs Excel `.xlsx` (OpenXML) sans dépendance externe.
- `data/` : base de données locale (`classeurs_rouges_db.json`) et sauvegardes automatiques (`backups/`) — *ignoré par Git*.
- `Rapports_Excel/` : dossier par défaut pour l'enregistrement des classeurs Excel générés — *ignoré par Git*.

## Prérequis et utilisation

1. Exécuter `Lancer_Classeurs_Rouges.bat` sous Windows (nécessite PowerShell 5.1+ et Microsoft Edge ; l'importation directe de fichiers `.xls`/`.xlsx` utilise Microsoft Excel s'il est installé, ou accepte directement les fichiers `.csv` sans Excel).
2. Dans l'onglet **Importer (PDF / XML / CSV)**, glisser-déposer ou sélectionner les fichiers XML de notes et d'absences, le fichier Excel/CSV des professeurs et/ou les bulletins PDF.
3. Consulter le suivi dans **Vue Classe / Cohorte (3 ans)** ou **Fiches Individuelles Élèves** en filtrant par **Année scolaire** et par **Classe**.
4. Cliquer sur **Générer le Classeur Excel (.xlsx)** pour exporter la sélection actuelle, ou sur **Générer 1 Excel par Classe** pour générer automatiquement un fichier Excel pour chaque classe de l'année scolaire choisie.
