import os
import zipfile
import win32com.client as win32
import time

start_time=time.time()

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"


print ("Dézip et réparation des exports NCE")
#  Dossier contenant tes .zip
zip_dir = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\INPUTS"
#  Dossier où on va tout extraire
extract_dir = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés\Fichiers dezip"
#  Dossier final pour les fichiers réparés
repaired_dir = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés"

os.makedirs(extract_dir, exist_ok=True)
os.makedirs(repaired_dir, exist_ok=True)

# Étape 1️⃣ : Dézipper tous les fichiers
for file in os.listdir(zip_dir):
    if file.lower().endswith(".zip"):
        zip_path = os.path.join(zip_dir, file)
        print(f"⏳ Dézippage : {file}")
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
            print(f"✅ Dézippé dans : {extract_dir}")
        except Exception as e:
            print(f"❌ Erreur dézippage {file} :", e)

# Étape 2️⃣ : Réparer tous les .xlsx extraits
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

for root, _, files in os.walk(extract_dir):
    for file in files:
        if file.lower().endswith(".xlsx",):
            in_path = os.path.join(root, file)
            out_path = os.path.join(repaired_dir, file)

            print(f"⏳ Réparation de : {file}")
            try:
                wb = excel.Workbooks.Open(os.path.abspath(in_path))
                # 51 = xlOpenXMLWorkbook (.xlsx)
                wb.SaveAs(os.path.abspath(out_path), FileFormat=51)
                wb.Close(False)
                print(f"✅ Fichier réparé ")
            except Exception as e:
                print(f"❌ Erreur sur {file} :", e)

excel.Quit()

print(" Tous les fichiers ont été dézippés et réparés.")

for _ in range(1000000):
    pass
end_time=time.time()

elapsed_time=end_time-start_time

min = elapsed_time // 60
sec= elapsed_time % 60

print(f"Temps d'execution : {int(min)} minutes et {sec:.2f} secondes")