# Mémo Power BI — Apprendre avec `ventes_retail_france.xlsx`

> Objectif : maîtriser les bases de Power BI (import Excel, Power Query, modèle, DAX, calendrier, visuels) en construisant un vrai rapport de ventes.

---

## 0. Ressources externes (à parcourir en premier)

- OpenClassrooms — **Utilisez DAX dans PowerBI** (libre accès, ~6 h, difficulté "difficile") : https://openclassrooms.com/fr/courses/8449411-utilisez-dax-dans-powerbi
  Prérequis annoncés : savoir créer un dashboard Power BI (nettoyage, modélisation, visualisation) et manipuler Excel. Le cours couvre colonnes calculées, mesures et filtres dans DAX.
- Coursera — **Modélisation des données dans Power BI** (Microsoft) : https://www.coursera.org/fr-FR/learn/modeling-data-in-power-bi
  Schémas de modélisation, syntaxe DAX, optimisation des performances.
- Coursera — **Power BI : Modélisation des données et Analyse des données** (préparation PL-300) : https://www.coursera.org/fr-FR/learn/power-bi-data-modeling-and-data-analysis
  Star schema, DAX, modélisation avancée.

Ordre conseillé : lis ce mémo d'abord (vue d'ensemble), fais le projet en parallèle, puis approfondis DAX avec le cours OpenClassrooms.

---

## 1. Diagnostic du fichier

| Élément | Constat |
|---|---|
| Feuille | 1 seule : `Ventes` (A1:O3001), **pas de Tableau Excel** (ni table nommée) |
| Volume | 3 000 lignes × 15 colonnes, 1 ligne = 1 commande (`ID Commande` unique) |
| Période | 01/01/2023 → 27/12/2025 (3 années quasi complètes) |
| Qualité | Aucune valeur manquante, aucun doublon, `Ventes = Qté × Prix × (1 − Remise)` vérifié sur 100 % des lignes, aucune expédition antérieure à la commande |
| Dates | Deux colonnes : `Date Commande` et `Date Expédition` (jusqu'au 01/01/2026) |
| Géographie | 8 régions, 24 villes (chaque ville appartient à une seule région → hiérarchie propre) |
| Produits | 3 catégories → 9 sous-catégories → 26 produits |
| Clients | 3 segments (Particulier, Professionnel, Grand Compte) ; 2 983 noms uniques pour 3 000 lignes, **pas d'ID client** |

Colonnes : `ID Commande`, `Date Commande`, `Date Expédition`, `Client`, `Segment Client`, `Région`, `Ville`, `Catégorie`, `Sous-Catégorie`, `Produit`, `Quantité`, `Prix Unitaire (€)`, `Remise (%)`, `Ventes (€)`, `Profit (€)`.

Points d'attention pour l'import :
- `Remise (%)` est stockée en **fraction** (0,05 = 5 %) → à formater en pourcentage dans Power BI, pas à multiplier par 100.
- Le fichier est très propre et très régulier (CA annuel plat, délais uniformes) : il ressemble à un jeu de données généré. Parfait pour apprendre, mais ne cherche pas de "vraies" tendances de marché dedans.

---

## 2. Ce que le fichier permet d'analyser (résultats réels, calculés sur les données)

**Chiffres clés** : CA total ≈ 2,71 M€ · Profit ≈ 384 k€ · marge ≈ 14,2 % · 3 000 commandes · panier moyen ≈ 902 €.

| Question business | Ce que montrent les données |
|---|---|
| Évolution annuelle | CA stable : 872 k€ (2023), 917 k€ (2024), 917 k€ (2025). Marge 14,6 % → 13,7 % → 14,3 % |
| Saisonnalité | Pic net en novembre–décembre (≈ 27,5 % du CA cumulé sur ces 2 mois) ; creux en février et mai |
| Mix catégories | **Technologie = 56 % du CA mais seulement 9,6 % de marge** ; Mobilier = 41 % du CA à 20 % de marge ; Fournitures = 2,6 % du CA à 21 % de marge |
| Produits | "Station de travail" = ~582 k€ de CA (≈ 21 % du total) avec ~9 % de marge. Les produits les plus rentables en % sont mobilier/fournitures |
| Remises | Marge de 14,9 % sans remise, mais 11–12 % dès 10 % de remise |
| Régions | PACA 1ʳᵉ en CA (373 k€), Île-de-France meilleure marge (15,3 %). Écarts régionaux modestes (13,6 % à 15,3 %) |
| Segments | Particuliers = 45 % du CA, Professionnels 35 %, Grands Comptes 20 % |
| Anomalies | 9 lignes à profit négatif (clavier mécanique, casque audio, portable 15"…) |
| Logistique | Délai d'expédition de 1 à 5 jours, moyenne ≈ 3 jours |

**Ce que le fichier NE permet PAS** (à savoir pour ne pas faire de contresens) :
- Fidélité / rétention / cohortes clients : pas d'ID client fiable.
- Objectifs vs réalisé : pas de budget (on peut en créer un avec une table saisie).
- Coûts détaillés, stocks, retours : absents (seul `Profit` existe).
- Carte précise : pas de latitude/longitude (mais `Ville` et `Région` suffisent à Power BI pour géocoder).

---

## 3. L'architecture de Power BI Desktop (à comprendre avant de cliquer)

Power BI Desktop = 3 couches, dans cet ordre :

1. **Power Query (Transformer les données)** : on se connecte à la source, on nettoie, on typage. Langage : **M**. Se rejoue à chaque actualisation.
2. **Modèle (vue Modèle)** : tables, relations, hiérarchies. On y pense en **schéma en étoile**.
3. **Rapport (DAX + visuels)** : mesures calculées dynamiquement selon les filtres, puis graphiques.

Règle d'or : **nettoyer en amont (Power Query), calculer en aval (DAX)**. Ne pas faire dans DAX ce qui se fait proprement dans Power Query, et inversement.

Les 3 vues à gauche : Rapport · Données (tableau) · Modèle.

---

## 4. Importer et manipuler le fichier Excel (Power Query)

### 4.1 Connexion
Accueil → **Obtenir des données → Excel** → sélectionner `ventes_retail_france.xlsx` → cocher la feuille `Ventes` → **Transformer les données** (pas "Charger" directement : on veut contrôler les types).

> Astuce : chemin local = fragile. Garde le fichier dans un dossier stable (ex. `D:\FORMATION\DATA et IA\...\data\`) et, plus tard, utilise un **paramètre** pour le chemin.

### 4.2 Étapes de nettoyage recommandées
1. Vérifier que la 1ʳᵉ ligne est bien promue en en-têtes.
2. **Typer chaque colonne** (c'est l'étape la plus importante) :
   - `Date Commande`, `Date Expédition` → **Date** (pas Date/Heure)
   - `Quantité` → Nombre entier
   - `Prix Unitaire (€)`, `Ventes (€)`, `Profit (€)` → Nombre décimal fixe (devise)
   - `Remise (%)` → Nombre décimal (puis format % dans le modèle)
   - Le reste → Texte
3. Renommer les colonnes avec des noms courts, sans caractères spéciaux (ex. `Ventes`, `Profit`, `Remise`, `PrixUnitaire`). Plus lisible dans les formules DAX.
4. Ajouter une colonne `Délai Expédition` (jours) — voir code M ci-dessous.
5. Ne **jamais** modifier le fichier source : tout est enregistré comme "étapes appliquées" et rejoué à chaque actualisation.

### 4.3 Code M (éditeur avancé) — requête `Ventes`

```m
let
    Source = Excel.Workbook(File.Contents("D:\FORMATION\DATA et IA\ventes_retail_france.xlsx"), null, true),
    Feuille = Source{[Item="Ventes", Kind="Sheet"]}[Data],
    EnTetes = Table.PromoteHeaders(Feuille, [PromoteAllScalars=true]),
    Types = Table.TransformColumnTypes(EnTetes, {
        {"ID Commande", type text},
        {"Date Commande", type date},
        {"Date Expédition", type date},
        {"Client", type text},
        {"Segment Client", type text},
        {"Région", type text},
        {"Ville", type text},
        {"Catégorie", type text},
        {"Sous-Catégorie", type text},
        {"Produit", type text},
        {"Quantité", Int64.Type},
        {"Prix Unitaire (€)", Currency.Type},
        {"Remise (%)", type number},
        {"Ventes (€)", Currency.Type},
        {"Profit (€)", Currency.Type}
    }),
    Delai = Table.AddColumn(Types, "Délai Expédition (j)",
        each Duration.Days([Date Expédition] - [Date Commande]), Int64.Type)
in
    Delai
```

(Si tu changes la langue/locale du fichier, le typage peut échouer : précise alors la culture, ex. `Table.TransformColumnTypes(..., "fr-FR")`.)

### 4.4 Créer des tables de dimensions par référence
Clic droit sur la requête `Ventes` → **Référence** (et non Dupliquer : on ne recharge pas la source plusieurs fois). Puis :

- `Dim_Produit` : garder `Produit`, `Sous-Catégorie`, `Catégorie` → **Supprimer les doublons** sur `Produit`.
- `Dim_Geographie` : garder `Ville`, `Région` → supprimer les doublons sur `Ville`.
- (Optionnel) `Dim_Segment` : une colonne `Segment Client` dédoublonnée.

**Ne supprime pas** `Sous-Catégorie`, `Catégorie` et `Région` dans la requête `Ventes` : les dimensions créées par Référence partent de sa dernière étape et seraient cassées. Une fois le modèle chargé, on **masque** ces colonnes dans la table de faits (vue Rapport/Modèle → clic droit → Masquer) pour n'utiliser que celles des dimensions.

> **Référence vs Dupliquer** : Référence = nouvelle requête qui part de l'étape finale de l'autre (dépendance). Dupliquer = copie indépendante de toutes les étapes. Pour un schéma en étoile, on utilise Référence.

Désactive "Activer le chargement" sur les requêtes intermédiaires que tu ne veux pas dans le modèle.

---

## 5. Modèle : schéma en étoile

```
            Dim_Calendrier
                  |  (Date → Date Commande)  [relation ACTIVE]
                  |  (Date → Date Expédition) [relation INACTIVE]
Dim_Produit —— Ventes (faits) —— Dim_Geographie
```

- **Table de faits** (`Ventes`) : événements mesurables (commandes, montants).
- **Dimensions** : axes d'analyse (qui, quoi, où, quand).
- Relations : **un-à-plusieurs**, filtre dans le sens **dimension → faits** (direction unique). Évite le bidirectionnel tant que tu ne sais pas pourquoi tu en as besoin.
- Pourquoi une étoile plutôt qu'une table plate ? Performances, filtres cohérents, DAX plus simple, calendrier unique.

**Cas particulier de ce fichier** : deux colonnes de date → une seule relation peut être active vers le calendrier. On met `Date Commande` en active (c'est l'axe de référence du CA) et `Date Expédition` en inactive, activable dans une mesure avec `USERELATIONSHIP` (voir §7.5).

---

## 6. Le calendrier (table de dates)

### 6.1 Pourquoi
Les fonctions d'intelligence temporelle DAX (`SAMEPERIODLASTYEAR`, `TOTALYTD`, `DATEADD`…) exigent une **table de dates continue** (sans trou), marquée comme telle. Ne jamais compter sur la hiérarchie de dates automatique de Power BI pour ça.

### 6.2 Création en DAX
Modélisation → **Nouvelle table** :

```dax
Dim_Calendrier =
VAR DateMin = DATE ( YEAR ( MIN ( Ventes[Date Commande] ) ), 1, 1 )
VAR DateMax = DATE ( YEAR ( MAX ( Ventes[Date Expédition] ) ), 12, 31 )
RETURN
    ADDCOLUMNS (
        CALENDAR ( DateMin, DateMax ),
        "Année", YEAR ( [Date] ),
        "Trimestre", "T" & QUARTER ( [Date] ),
        "Mois N°", MONTH ( [Date] ),
        "Mois", FORMAT ( [Date], "MMMM", "fr-FR" ),
        "Année-Mois", FORMAT ( [Date], "YYYY-MM" ),
        "Jour Sem. N°", WEEKDAY ( [Date], 2 ),
        "Jour Sem.", FORMAT ( [Date], "dddd", "fr-FR" )
    )
```

Pourquoi `DateMax` s'appuie sur `Date Expédition` : la dernière expédition est le **01/01/2026**. Si le calendrier s'arrêtait fin 2025, cette ligne n'aurait pas de date correspondante pour la relation inactive. Conséquence : 2026 apparaîtra dans les segments ; filtre-le (volet Filtres : Année ≤ 2025) ou masque-le.

### 6.3 Étapes obligatoires après création
1. **Trier `Mois` par `Mois N°`** : colonne `Mois` → Outils de colonne → *Trier par colonne* → `Mois N°` (sinon avril avant janvier, ordre alphabétique).
2. Idem pour `Jour Sem.` trié par `Jour Sem. N°`.
3. Clic droit sur la table → **Marquer comme table de dates** → colonne `Date`.
4. Vue Modèle : relier `Dim_Calendrier[Date]` → `Ventes[Date Commande]` (active) et → `Ventes[Date Expédition]` (inactive, trait pointillé).
5. Masquer `Ventes[Date Commande]` et `[Date Expédition]` dans la vue Rapport pour forcer l'usage du calendrier.

### 6.4 Alternative Power Query
Tu peux aussi générer le calendrier en M (`List.Dates`) — choix équivalent. DAX est plus court ; M a l'avantage d'être calculé au chargement. Pour apprendre, commence par DAX.

---

## 7. DAX — les concepts qui comptent

### 7.1 Colonne calculée vs Mesure (la distinction n°1)

| | Colonne calculée | Mesure |
|---|---|---|
| Calcul | Ligne par ligne, **à l'actualisation** | **Dynamique**, selon le contexte de filtre du visuel |
| Stockage | Stockée dans le modèle (mémoire) | Non stockée |
| Usage typique | Catégoriser, trier, trancher | Totaux, ratios, KPI |

Règle : **pour un chiffre qui s'agrège → mesure**. Une colonne calculée sert à créer une catégorie ou un axe.

Crée une table vide `_Mesures` (Saisir des données) pour ranger toutes tes mesures.

### 7.2 Contexte de filtre et contexte de ligne
- **Contexte de filtre** : ce que le visuel, les segments et les filtres imposent (ex. la ligne "2024 / Mobilier"). Une mesure s'évalue **dans ce contexte**.
- **Contexte de ligne** : "on est positionné sur une ligne" (colonnes calculées, itérateurs `SUMX`…).
- `CALCULATE` = **modifier le contexte de filtre**. C'est la fonction centrale de DAX.

### 7.3 Mesures de base (copier-coller)

```dax
Total Ventes = SUM ( Ventes[Ventes (€)] )

Total Profit = SUM ( Ventes[Profit (€)] )

Marge % = DIVIDE ( [Total Profit], [Total Ventes] )

Nb Commandes = DISTINCTCOUNT ( Ventes[ID Commande] )

Panier Moyen = DIVIDE ( [Total Ventes], [Nb Commandes] )

Quantité Vendue = SUM ( Ventes[Quantité] )

Délai Moyen (j) = AVERAGE ( Ventes[Délai Expédition (j)] )
```

Nomme les mesures `Total Ventes` et `Total Profit` (et non `Ventes` / `Profit`) pour ne pas les confondre avec la table `Ventes`.

Utilise toujours `DIVIDE` (gère la division par zéro) plutôt que `/`.
Séparateur : selon la langue de ton Power BI, DAX utilise `,` (anglais) ou `;` (français) entre arguments. Si une formule est refusée alors qu'elle est correcte, vérifie ça en premier.

### 7.4 CALCULATE et filtres

```dax
Ventes Technologie =
CALCULATE ( [Total Ventes], Dim_Produit[Catégorie] = "Technologie" )

Ventes Remisées =
CALCULATE ( [Total Ventes], Ventes[Remise (%)] > 0 )

Part du CA % =
DIVIDE ( [Total Ventes], CALCULATE ( [Total Ventes], ALLSELECTED ( Dim_Produit ) ) )

Profit Négatif (lignes) =
CALCULATE ( COUNTROWS ( Ventes ), Ventes[Profit (€)] < 0 )
```

- `ALL` : enlève tous les filtres (grand total absolu). `ALLSELECTED` : respecte les segments mais ignore le contexte du visuel → idéal pour des **parts en %**.
- `FILTER` est un itérateur coûteux : dans `CALCULATE`, préfère une condition simple quand c'est possible.

### 7.5 Intelligence temporelle (nécessite le calendrier)

```dax
Ventes N-1 =
CALCULATE ( [Total Ventes], SAMEPERIODLASTYEAR ( Dim_Calendrier[Date] ) )

Évolution N-1 % =
DIVIDE ( [Total Ventes] - [Ventes N-1], [Ventes N-1] )

Ventes YTD =
TOTALYTD ( [Total Ventes], Dim_Calendrier[Date] )

Ventes 3 Mois Glissants =
CALCULATE ( [Total Ventes], DATESINPERIOD ( Dim_Calendrier[Date], MAX ( Dim_Calendrier[Date] ), -3, MONTH ) )

Ventes par Date d'Expédition =
CALCULATE ( [Total Ventes], USERELATIONSHIP ( Dim_Calendrier[Date], Ventes[Date Expédition] ) )
```

Attention : `Ventes N-1` est **vide pour 2023** (aucune donnée 2022). C'est normal, pas un bug.

### 7.6 Itérateurs et classements

```dax
Marge Ligne Moyenne % =
AVERAGEX ( Ventes, DIVIDE ( Ventes[Profit (€)], Ventes[Ventes (€)] ) )

Rang Produit (Profit) =
RANKX ( ALL ( Dim_Produit[Produit] ), [Total Profit] )
```

`SUMX`, `AVERAGEX`, `RANKX` = itérateurs : ils évaluent une expression **ligne par ligne**, puis agrègent.

### 7.7 Mesures utiles pour ce projet

```dax
Marge Cible = 0.15   -- exemple de paramètre

Écart vs Marge Cible = [Marge %] - [Marge Cible]

Couleur Marge =
IF ( [Marge %] >= [Marge Cible], "#2E7D32", "#C62828" )
```

(La dernière sert à la mise en forme conditionnelle : *Format → Couleur → Valeur du champ*.)

### 7.7 bis Objectifs vs réalisé (exercice 10)

Table `Objectifs` saisie à la main (`Catégorie`, `Marge Cible`), **sans relation** avec le modèle : `Dim_Produit` a plusieurs lignes par catégorie, la relation serait mal orientée. On récupère la cible avec `LOOKUPVALUE`.

```dax
Marge Cible % =
VAR Cat = SELECTEDVALUE ( Dim_Produit[Catégorie] )
RETURN
    LOOKUPVALUE ( Objectifs[Marge Cible], Objectifs[Catégorie], Cat )

Écart Marge = [Marge %] - [Marge Cible %]

Statut Marge =
IF (
    ISBLANK ( [Marge Cible %] ),
    BLANK (),
    IF ( [Marge %] >= [Marge Cible %], "Atteint", "Sous l'objectif" )
)

Profit Cible =
SUMX ( VALUES ( Dim_Produit[Catégorie] ), [Total Ventes] * [Marge Cible %] )

Écart Profit = [Total Profit] - [Profit Cible]
```

- `SELECTEDVALUE` renvoie la catégorie si une seule est active, sinon vide : `Marge Cible %` est donc vide sur la ligne du total.
- `SUMX` itère sur les catégories : `Profit Cible` a, lui, un total correct.
- Objectifs d'exemple utilisés : Technologie 12 %, Mobilier 20 %, Fournitures de bureau 20 %.
- Valeurs obtenues : écart de profit ≈ −36 788 € en Technologie, ≈ −34 490 € au total.
- Alternative plus propre à terme : une vraie `Dim_Categorie` reliée à `Dim_Produit` et à `Objectifs`.

### 7.8 Bonnes pratiques DAX
- Une mesure = un rôle. Compose les mesures entre elles plutôt que de répéter les formules.
- `VAR … RETURN` pour la lisibilité et la performance.
- Commente (`--` ou `//`) et formate (le bouton "Mettre en forme" ou daxformatter.com).
- Format des mesures : € (devise), % (une décimale), nombre entier.

---

## 8. Construire le rapport (visuels)

**Page 1 — Vue d'ensemble**
- Cartes : `Total Ventes`, `Total Profit`, `Marge %`, `Nb Commandes`, `Panier Moyen`.
- Courbe : `Ventes` et `Ventes N-1` par `Année-Mois` (axe = Dim_Calendrier).
- Histogramme empilé : `Ventes` par `Catégorie`.
- Segments : `Année`, `Segment Client`, `Région`.

**Page 2 — Rentabilité**
- Nuage de points : axe X = `Ventes`, axe Y = `Marge %`, détails = `Sous-Catégorie` (révèle que Technologie est large mais peu rentable).
- Matrice : `Catégorie > Sous-Catégorie > Produit` × `Ventes`, `Profit`, `Marge %` avec mise en forme conditionnelle.
- Histogramme `Marge %` par `Remise (%)`.

**Page 3 — Géographie**
- Carte (ou carte remplie) par `Région`, taille = `Ventes`, couleur = `Marge %`.
- Hiérarchie Région → Ville (drill-down).

**Page 4 — Opérations**
- Histogramme `Nb Commandes` par `Délai Expédition (j)`.
- Tableau des lignes à profit négatif (filtre visuel `Profit < 0`).

Fonctions à connaître : **Info-bulles personnalisées**, **Drill-through** (page de détail d'un produit), **Signets**, **Interactions entre visuels** (Format → Modifier les interactions), **Analyseur de performances**.

---

## 9. Exercices progressifs (à faire dans l'ordre)

| # | Exercice | Compétence |
|---|---|---|
| 1 | Importer le fichier, typer toutes les colonnes, ajouter `Délai Expédition (j)` | Power Query |
| 2 | Créer `Dim_Produit` et `Dim_Geographie` par Référence | Modélisation Power Query |
| 3 | Créer `Dim_Calendrier`, la marquer, trier les mois, créer les relations | Calendrier |
| 4 | Écrire les 7 mesures de base et les afficher en cartes | DAX de base |
| 5 | `Part du CA %` par catégorie avec `ALLSELECTED` | `CALCULATE` |
| 6 | `Ventes N-1`, `Évolution N-1 %`, `Ventes YTD` | Intelligence temporelle |
| 7 | Comparer CA par date de commande vs date d'expédition | `USERELATIONSHIP` |
| 8 | Répondre : "Quelle sous-catégorie détruit de la marge ?" (nuage de points) | Analyse |
| 9 | Ajouter une page Drill-through "Détail Produit" | Navigation |
| 10 | Créer une table `Objectifs` (marge cible par catégorie) et comparer au réalisé | Modèle + DAX |

---

## 10. Pièges fréquents

- **Types non définis** : une date importée en texte casse tout le calendrier.
- **Relation bidirectionnelle par réflexe** : provoque des ambiguïtés et des lenteurs.
- **Tout en colonnes calculées** : alourdit le modèle. Préfère les mesures.
- **Oublier de marquer le calendrier** → les fonctions temporelles renvoient des résultats faux ou vides.
- **Mois triés alphabétiquement** → trier par `Mois N°`.
- **Somme implicite** : Power BI additionne automatiquement les colonnes numériques glissées dans un visuel. Crée des **mesures explicites** et masque les colonnes brutes.
- **Moyenne d'une moyenne** : une `Marge %` ne se moyenne pas, elle se recalcule (`Profit / Ventes`) — d'où la mesure DAX.
- **Séparateur DAX** `,` vs `;` selon la langue de l'interface.
- **Chemin du fichier en dur** : l'actualisation échoue si le fichier bouge → utiliser un paramètre.

---

## 11. Pour aller plus loin
- Paramètres Power Query (chemin, année de début).
- Modèle d'objectifs / what-if (paramètre de scénario sur la remise).
- Sécurité au niveau des lignes (RLS) par région.
- Publication sur le service Power BI et actualisation planifiée (passerelle pour un fichier local).
- Lien avec ta feuille de route : brancher la même table sur PostgreSQL au lieu d'Excel (même modèle, source différente) et comparer Import vs DirectQuery.
