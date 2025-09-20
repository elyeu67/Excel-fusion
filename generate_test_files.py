import pandas as pd
import os
import random
from datetime import datetime, timedelta
import numpy as np

def generate_test_excel_files():
    """Génère des fichiers Excel de test avec des données variées"""
    
    # Créer le dossier de test
    test_folder = "fichiers_excel_test"
    if not os.path.exists(test_folder):
        os.makedirs(test_folder)
        print(f"📁 Dossier créé: {test_folder}")
    
    # Données de base pour générer des fichiers variés
    noms = ["Jean", "Marie", "Pierre", "Sophie", "Paul", "Julie", "Marc", "Claire", "Thomas", "Emma"]
    villes = ["Paris", "Lyon", "Marseille", "Toulouse", "Nice", "Nantes", "Strasbourg", "Montpellier", "Bordeaux", "Lille"]
    produits = ["Ordinateur", "Téléphone", "Tablette", "Casque", "Souris", "Clavier", "Écran", "Imprimante", "Routeur", "Webcam"]
    
    # Générer 15 fichiers Excel différents
    for i in range(1, 16):
        # Créer des données aléatoires pour chaque fichier
        nb_lignes = random.randint(10, 50)
        
        # Données communes à tous les fichiers
        data = {
            'ID': range(1, nb_lignes + 1),
            'Nom': [random.choice(noms) for _ in range(nb_lignes)],
            'Ville': [random.choice(villes) for _ in range(nb_lignes)],
            'Prix': [round(random.uniform(10, 1000), 2) for _ in range(nb_lignes)],
            'Date': [(datetime.now() - timedelta(days=random.randint(1, 365))).strftime('%Y-%m-%d') 
                    for _ in range(nb_lignes)]
        }
        
        # Ajouter des colonnes spécifiques selon le type de fichier
        if i <= 5:  # Fichiers de ventes
            data['Produit'] = [random.choice(produits) for _ in range(nb_lignes)]
            data['Quantité'] = [random.randint(1, 20) for _ in range(nb_lignes)]
            data['Total'] = [data['Prix'][j] * data['Quantité'][j] for j in range(nb_lignes)]
            filename = f"ventes_{i:02d}.xlsx"
            
        elif i <= 10:  # Fichiers de clients
            data['Email'] = [f"{data['Nom'][j].lower()}@email.com" for j in range(nb_lignes)]
            data['Téléphone'] = [f"0{random.randint(100000000, 999999999)}" for _ in range(nb_lignes)]
            data['Statut'] = [random.choice(['Actif', 'Inactif', 'Prospect']) for _ in range(nb_lignes)]
            filename = f"clients_{i-5:02d}.xlsx"
            
        else:  # Fichiers de commandes
            data['Commande_ID'] = [f"CMD{random.randint(1000, 9999)}" for _ in range(nb_lignes)]
            data['Statut_Commande'] = [random.choice(['En cours', 'Livré', 'Annulé']) for _ in range(nb_lignes)]
            data['Méthode_Paiement'] = [random.choice(['Carte', 'Virement', 'Chèque', 'Espèces']) for _ in range(nb_lignes)]
            filename = f"commandes_{i-10:02d}.xlsx"
        
        # Créer le DataFrame
        df = pd.DataFrame(data)
        
        # Sauvegarder le fichier
        filepath = os.path.join(test_folder, filename)
        df.to_excel(filepath, index=False)
        
        print(f"✅ Fichier créé: {filename} ({nb_lignes} lignes)")
    
    # Créer quelques fichiers avec des structures légèrement différentes
    print("\n🔄 Création de fichiers avec structures variées...")
    
    # Fichier avec des colonnes en plus
    extra_data = {
        'ID': range(1, 21),
        'Nom': [random.choice(noms) for _ in range(20)],
        'Ville': [random.choice(villes) for _ in range(20)],
        'Prix': [round(random.uniform(10, 1000), 2) for _ in range(20)],
        'Date': [(datetime.now() - timedelta(days=random.randint(1, 365))).strftime('%Y-%m-%d') 
                for _ in range(20)],
        'Produit': [random.choice(produits) for _ in range(20)],
        'Quantité': [random.randint(1, 20) for _ in range(20)],
        'Total': [0] * 20,  # Sera calculé
        'Remise': [random.uniform(0, 0.2) for _ in range(20)],
        'Notes': [f"Note {i}" for i in range(1, 21)]
    }
    extra_data['Total'] = [extra_data['Prix'][i] * extra_data['Quantité'][i] * (1 - extra_data['Remise'][i]) 
                          for i in range(20)]
    
    df_extra = pd.DataFrame(extra_data)
    df_extra.to_excel(os.path.join(test_folder, "ventes_avec_remises.xlsx"), index=False)
    print("✅ Fichier créé: ventes_avec_remises.xlsx (20 lignes)")
    
    # Fichier avec moins de colonnes
    simple_data = {
        'ID': range(1, 15),
        'Nom': [random.choice(noms) for _ in range(14)],
        'Prix': [round(random.uniform(10, 1000), 2) for _ in range(14)]
    }
    
    df_simple = pd.DataFrame(simple_data)
    df_simple.to_excel(os.path.join(test_folder, "donnees_simples.xlsx"), index=False)
    print("✅ Fichier créé: donnees_simples.xlsx (14 lignes)")
    
    # Fichier avec des données manquantes
    missing_data = {
        'ID': range(1, 25),
        'Nom': [random.choice(noms) if random.random() > 0.1 else None for _ in range(24)],
        'Ville': [random.choice(villes) if random.random() > 0.15 else None for _ in range(24)],
        'Prix': [round(random.uniform(10, 1000), 2) if random.random() > 0.05 else None for _ in range(24)],
        'Date': [(datetime.now() - timedelta(days=random.randint(1, 365))).strftime('%Y-%m-%d') 
                if random.random() > 0.08 else None for _ in range(24)]
    }
    
    df_missing = pd.DataFrame(missing_data)
    df_missing.to_excel(os.path.join(test_folder, "donnees_avec_vides.xlsx"), index=False)
    print("✅ Fichier créé: donnees_avec_vides.xlsx (24 lignes)")
    
    print(f"\n🎉 Génération terminée !")
    print(f"📊 Total: {len(os.listdir(test_folder))} fichiers Excel créés dans le dossier '{test_folder}'")
    print(f"📁 Chemin complet: {os.path.abspath(test_folder)}")
    print(f"\n💡 Vous pouvez maintenant tester l'application de fusion avec ces fichiers !")

if __name__ == "__main__":
    generate_test_excel_files()
