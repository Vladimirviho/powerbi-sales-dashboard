# Modèle Power BI : Power Query, relations et mesures DAX

## Tables

| Table | Rôle | Origine |
|---|---|---|
| `Ventes` | Table de faits (une ligne par commande) | Excel, feuille `Ventes` |
| `Dim_Produit` | Produit, Sous-Catégorie, Catégorie | Référence de `Ventes` |
| `Dim_Geographie` | Ville, Région | Référence de `Ventes` |
| `Dim_Calendrier` | Table de dates, marquée comme table de dates | DAX |
| `Objectifs` | Marge cible par catégorie (valeurs d'exemple) | Saisie manuelle |
| `_Mesures` | Table vide qui regroupe les mesures | Saisie manuelle |

## Relations

| De | Vers | Cardinalité | État |
|---|---|---|---|
| `Dim_Calendrier[Date]` | `Ventes[Date Commande]` | un à plusieurs | Active |
| `Dim_Calendrier[Date]` | `Ventes[Date Expédition]` | un à plusieurs | Inactive |
| `Dim_Produit[Produit]` | `Ventes[Produit]` | un à plusieurs | Active |
| `Dim_Geographie[Ville]` | `Ventes[Ville]` | un à plusieurs | Active |

`Objectifs` n'a aucune relation : la cible est lue avec `LOOKUPVALUE`.

## Power Query (langage M)

### Requête `Ventes`

Remplacer le chemin par celui du fichier Excel.

```m
let
    Source = Excel.Workbook(File.Contents("C:\chemin\vers\data\ventes_retail_france.xlsx"), null, true),
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

### Requêtes `Dim_Produit` et `Dim_Geographie` (par référence à `Ventes`)

```m
let
    Source = Ventes,
    Colonnes = Table.SelectColumns(Source, {"Produit", "Sous-Catégorie", "Catégorie"}),
    SansDoublons = Table.Distinct(Colonnes, {"Produit"})
in
    SansDoublons
```

```m
let
    Source = Ventes,
    Colonnes = Table.SelectColumns(Source, {"Ville", "Région"}),
    SansDoublons = Table.Distinct(Colonnes, {"Ville"})
in
    SansDoublons
```

Contrôles : 26 produits, 24 villes.

## Table de dates (DAX)

La date de fin s'appuie sur la date d'expédition : la dernière expédition tombe le 01/01/2026.

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

Étapes manuelles : trier `Mois` par `Mois N°` et `Jour Sem.` par `Jour Sem. N°`, puis marquer la table comme table de dates (colonne `Date`). Contrôle : 1 461 lignes.

## Mesures DAX

### Base

```dax
Total Ventes = SUM ( Ventes[Ventes (€)] )

Total Profit = SUM ( Ventes[Profit (€)] )

Marge % = DIVIDE ( [Total Profit], [Total Ventes] )

Nb Commandes = DISTINCTCOUNT ( Ventes[ID Commande] )

Panier Moyen = DIVIDE ( [Total Ventes], [Nb Commandes] )

Quantité Vendue = SUM ( Ventes[Quantité] )

Délai Moyen (j) = AVERAGE ( Ventes[Délai Expédition (j)] )
```

### Part du chiffre d'affaires

```dax
Part du CA % (ALL) =
DIVIDE ( [Total Ventes], CALCULATE ( [Total Ventes], ALL ( Dim_Produit ) ) )
```

Une variante avec `ALLSELECTED` a été essayée sans succès et n'a pas été retenue.

### Intelligence temporelle

```dax
Total Ventes N-1 =
CALCULATE ( [Total Ventes], SAMEPERIODLASTYEAR ( Dim_Calendrier[Date] ) )

Évolution N-1 % =
DIVIDE ( [Total Ventes] - [Total Ventes N-1], [Total Ventes N-1] )

Total Ventes YTD =
TOTALYTD ( [Total Ventes], Dim_Calendrier[Date] )

Total Ventes (expédition) =
CALCULATE (
    [Total Ventes],
    USERELATIONSHIP ( Dim_Calendrier[Date], Ventes[Date Expédition] )
)
```

`Total Ventes N-1` est vide en 2023 : il n'y a pas de données en 2022.

### Objectifs de marge

Table `Objectifs` : Technologie 0,12 ; Mobilier 0,20 ; Fournitures de bureau 0,20 (valeurs d'exemple).

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

## Valeurs de contrôle

| Mesure | Valeur attendue |
|---|---|
| Total Ventes | 2 706 412 € |
| Total Profit | 384 496 € |
| Marge % | 14,2 % |
| Nb Commandes | 3 000 |
| Panier Moyen | 902 € |
| Quantité Vendue | 12 188 |
| Délai Moyen (j) | 3,0 |
| Écart Profit, Technologie | environ −36 788 € |
