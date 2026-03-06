import os
import zipfile
import win32com.client as win32

# Supression fichiers unzipés pour gagner de l'espace disque

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"

folder = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés\Fichiers dezip"

for file in os.listdir(folder):
    file_path = os.path.join(folder, file)
    try:
        if os.path.isfile(file_path):
            os.remove(file_path) # supprime le fichier
            #print(f"🗑️ supprimé : {file_path}")
    except Exception as e:
        print(f"❌ erreur avec {file_path} : {e}")

#print("✅ Tous les fichiers ont été supprimés du dossier.")

# Supression fichiers unzipés pour gagner de l'espace disque
folder = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés"

for file in os.listdir(folder):
    file_path = os.path.join(folder, file)
    try:
        if os.path.isfile(file_path):
            os.remove(file_path) # supprime le fichier
            #print(f"🗑️ supprimé : {file_path}")
    except Exception as e:
        print(f"❌ erreur avec {file_path} : {e}")

#print("✅ Tous les fichiers ont été supprimés du dossier.")