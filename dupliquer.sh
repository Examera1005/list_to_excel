#!/bin/zsh

# Script de duplication de template Excel par élève
# Usage : ./dupliquer.sh

# 1. Chemin du fichier template
TEMPLATE="template.xlsx"

# 2. Vérification de l'existence du template
if [ ! -f "$TEMPLATE" ]; then
    echo "❌ Erreur : Le fichier '$TEMPLATE' est introuvable."
    echo "Assurez-vous qu'il est présent dans le même dossier que le script."
    exit 1
fi

# 3. Liste des élèves (Fichier eleves.txt s'il existe, sinon liste par défaut)
LISTE_ELEVES="eleves.txt"

if [ -f "$LISTE_ELEVES" ]; then
    echo "📄 Lecture de la liste depuis '$LISTE_ELEVES'..."
    while IFS= read -r line || [ -n "$line" ]; do
        # Ignorer les lignes vides
        [ -z "$line" ] && continue
        
        # Nettoyage du nom (remplace les espaces par des underscores)
        NOM_CLEAN=$(echo "$line" | tr ' ' '_')
        
        # Nom du nouveau fichier
        NOUVEAU_FICHIER="${NOM_CLEAN}.xlsx"
        
        # Duplication
        cp "$TEMPLATE" "$NOUVEAU_FICHIER"
        echo "✅ Créé : $NOUVEAU_FICHIER"
    done < "$LISTE_ELEVES"
else
    echo "⚠️ Aucun fichier '$LISTE_ELEVES' trouvé."
    echo "💡 Vous pouvez créer un fichier 'eleves.txt' avec un prénom/nom par ligne."
fi

echo "🎉 Opération terminée !"
