# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil en Python pour générer automatiquement les fichiers d'évaluation Excel individuels par élève à partir de vos listes de classes (`.xlsx` ou `.txt`) et d'un template modèle (`.xlsx`).

Compatible avec **macOS**, **Linux** et **Windows**.

---

## 📁 Fichiers du projet

* **`dupliquer.py`** : Script Python autonome pour exécuter la génération automatique (avec support multi-classes).
* **`app_tui.py`** : Interface interactive (TUI) guidée avec mode multi-classes et contrôle des saisies.
* **`template.xlsx`** : Exemple de fichier modèle Excel.
* **`liste_eleves.xlsx`** : Exemple de fichier de liste d'élèves (*Nom*, *Prénom*, *Groupe/Classe*).

---

## ✨ Fonctionnalités clés

* 🚀 **Traitement par lots Multi-classes (Batch Mode)** : Traitez 5 (ou plus) listes de classes d'un coup en **1 seule exécution** sans relancer le script ni copier le fichier `.py` !
* 📂 **Organisation automatique en sous-dossiers** : Si plusieurs classes sont sélectionnées, le script crée automatiquement un sous-dossier propre par classe (ex: `./fichiers_eleves/1M1/`, `./fichiers_eleves/1M2/`...).
* 🏷️ **Nommage personnalisé** : Format `[Classe]_[Nom]_[Prénom].xlsx` (ex: `1M1_Dupont_Alice.xlsx`).
* ✏️ **Cellule C3 automatique** : Écrit le nom complet (`Prénom Nom`) dans la cellule **C3** de chaque fichier.
* 🛡️ **Validation des erreurs TUI** : Si l'utilisateur saisit une option invalide (ex: `lol`), le programme redemande la saisie au lieu de continuer avec une valeur par défaut.

---

## 🚀 Utilisation

### Mode 1 : Interface interactive avec mode Multi-classes (Recommandé)

Lancez le menu guidé et choisissez **Traitement par lots (Multi-classes)** pour cocher les listes que vous voulez traiter :

```bash
python3 app_tui.py
```

---

### Mode 2 : Lancement automatique en ligne de commande

Traite automatiquement toutes les listes de classes présentes dans le dossier :

```bash
python3 dupliquer.py
```

Vous pouvez aussi spécifier le template et les fichiers de listes directement :

```bash
python3 dupliquer.py template.xlsx liste_1M1.xlsx liste_1M2.xlsx liste_1M3.xlsx
```
