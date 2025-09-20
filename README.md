# Fusionneur de Fichiers Excel

Une application Python avec interface graphique pour fusionner facilement des dizaines de fichiers Excel en un seul fichier consolidé.

## Fonctionnalités

- ✅ Interface graphique simple et intuitive
- ✅ Fusion automatique de tous les fichiers .xlsx et .xls d'un dossier
- ✅ Option pour ajouter une colonne avec le nom du fichier source
- ✅ Gestion des en-têtes (option pour ignorer les en-têtes des fichiers sources)
- ✅ Barre de progression en temps réel
- ✅ Journal détaillé des opérations
- ✅ Gestion d'erreurs robuste
- ✅ Support des formats Excel (.xlsx et .xls)

## Installation

1. Assurez-vous d'avoir Python 3.7 ou plus récent installé
2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation

1. Lancez l'application :
```bash
python main.py
```

2. Dans l'interface :
   - **Sélectionnez le dossier** contenant vos fichiers Excel
   - **Choisissez l'emplacement** où sauvegarder le fichier fusionné
   - **Configurez les options** selon vos besoins :
     - Ajouter une colonne avec le nom du fichier source
     - Ignorer les en-têtes dans les fichiers sources
   - **Cliquez sur "Fusionner les fichiers"**

3. Suivez la progression dans la barre de progression et le journal

## Options disponibles

- **Ajouter une colonne avec le nom du fichier source** : Ajoute une colonne "Fichier_Source" pour identifier l'origine de chaque ligne
- **Ignorer les en-têtes dans les fichiers sources** : Garde seulement les en-têtes du premier fichier (utile si tous les fichiers ont la même structure)

## Structure des fichiers supportés

L'application peut fusionner des fichiers Excel avec des structures différentes. Les colonnes seront alignées automatiquement par nom.

## Dépendances

- `pandas` : Manipulation des données
- `openpyxl` : Lecture/écriture des fichiers .xlsx
- `xlrd` : Lecture des fichiers .xls (ancien format)

## Exemple d'utilisation

1. Placez tous vos fichiers Excel dans un dossier
2. Lancez l'application
3. Sélectionnez le dossier
4. Choisissez où sauvegarder le résultat
5. Cliquez sur "Fusionner"
6. Récupérez votre fichier consolidé !

## Support

L'application gère automatiquement les erreurs et affiche des messages informatifs dans le journal des opérations.
