# 📚 Bibliothèque VBA

![Langage](https://img.shields.io/badge/langage-VBA-blue)
![Licence](https://img.shields.io/badge/Licence-MIT-green)

Module VBA regroupant des fonctions pour lire et écrire dans le presse-papiers de Windows via les API système.

---

## 📄 Licence

Ce projet est distribué sous licence **MIT**.  
Consultez le fichier [`LICENSE`](LICENSE) pour plus de détails.

---

## 🧰 Prérequis

- Environnement : **Microsoft Visual Basic for Applications (VBA)**
- Compatible Excel (Windows)

---

# 🧩 Fonctions et procédures disponibles

Le module expose deux fonctions simples :

|Module|Descriptif|
|------|----------|
|LirePressePapiers|Retourne le texte actuellement présent dans le presse-papiers.|
|EcrirePressePapiers|Remplace le contenu du presse-papiers par la chaîne fournie.|

Ces fonctions encapsulent les API Windows nécessaires (Win32) pour une utilisation directe en VBA.
