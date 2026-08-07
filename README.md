# 🎓 Générateur de Fichiers Élèves (Template Excel)

Un outil simple et efficace pour dupliquer automatiquement un fichier modèle (Template Excel `.xlsx`) pour une liste d'élèves (fournie via un fichier texte `.txt`).

 Compatible avec **macOS**, **Linux** et **Windows**.

---

## 📁 Structure du projet

* **`app_tui.py`** : Interface interactive (TUI) complète et guidée en Python (recommandé).
* **`dupliquer.py`** : Script Python autonome léger et rapide.
* **`dupliquer.sh`** : Script Shell / Zsh alternatif pour environnement Unix.
* **`eleves.txt`** : Fichier exemple contenant la liste des élèves (un nom par ligne).

---

## 🚀 Utilisation

### 1. Interface Interactive (Recommandé)

Lancez l'interface guidée dans le terminal :

```bash
python3 app_tui.py
```

**Fonctionnalités de l'interface (`app_tui.py`) :**
* 🔍 Détection automatique des fichiers `.xlsx` et `.txt` dans le dossier courant.
* 🖥️ Intégration du sélecteur de fichiers/dossiers natif macOS (`osascript`).
* 📁 Choix d'un dossier de destination personnalisé.
* 🏷️ Ajout d'un préfixe ou d'un suffixe aux noms de fichiers (ex: `Marie_Curie_CM2.xlsx` ou `2026_Marie_Curie.xlsx`).
* 📋 Aperçu récapitulatif avant génération.

---

### 2. Script Python simple

Pour une exécution rapide basée sur un fichier `template.xlsx` et `eleves.txt` situés dans le même dossier :

```bash
python3 dupliquer.py
```

---

### 3. Script Shell (Bash / Zsh)

Alternative en script shell :

```bash
chmod +x dupliquer.sh
./dupliquer.sh
```

---

## 📝 Format de la liste d'élèves (`eleves.txt`)

Créez un fichier texte nommé `eleves.txt` contenant un nom ou prénom d'élève par ligne :

```text
Marie Curie
Albert Einstein
Isaac Newton
```

Les espaces dans les noms seront automatiquement remplacés par des tirets bas (`_`) lors de la création des fichiers.
