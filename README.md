# Module VBA – Gestion du presse-papiers Windows

Ce dépôt fournit un module VBA permettant de **lire** et **écrire** du texte dans le presse-papiers Windows, fonctionnalité non disponible nativement en VBA.

---

## Objectif

Le module expose deux fonctions simples :

- `LirePressePapiers() As String`  
  Retourne le texte actuellement présent dans le presse-papiers.

- `EcrirePressePapiers(texte As String)`  
  Remplace le contenu du presse-papiers par la chaîne fournie.

Ces fonctions encapsulent les API Windows nécessaires (Win32) pour une utilisation directe en VBA.

---

## Contenu du dépôt

- `PressePapiers.bas` : module VBA contenant les déclarations API et les fonctions publiques.
- `README.md` : ce fichier de documentation.
- `LICENSE` : licence MIT.

Aucune dépendance externe n’est requise.

---

## Installation

1. Ouvrir l’éditeur VBA (`Alt` + `F11`).
2. Menu **Fichier** → **Importer un fichier…**.
3. Sélectionner le fichier `PressePapiers.bas`.
4. Le module est alors disponible dans tout le projet VBA.

---

## Utilisation

### Lire le presse-papiers

```vba
Dim txt As String
txt = LirePressePapiers()
MsgBox txt
