# Changments par 
import pandas as pd
import re
import warnings
import os

warnings.simplefilter("ignore")


print("⏳Traitement du fichier Export Service Port")
print("Veuillez patienter svp, les exports NCE sont lourds ")

print ("Fusion des MA5800-X17_ServicePort")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"

# CHEMINS
DOSSIER = chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés"
PREFIX = "MA5800-X17_ServicePort"
OUTPUT = os.path.join(DOSSIER, "Fusion_MA5800-X17_ServicePort.csv")
#SKIP_FILE = None # on va définir quel fichier doit sauter 7 lignes
# =====================

# lister fichiers
files = sorted([f for f in os.listdir(DOSSIER) if f.startswith(PREFIX) and f.endswith(".xlsx")])
if len(files) < 2:
    raise SystemExit(f"Moins de 2 fichiers trouvés avec {PREFIX} dans {DOSSIER}")

print("Fichiers trouvés :", files)

first_file=True
with open(OUTPUT,"w",encoding="utf-8", newline="")as outfile:
    for f in files:
        path = os.path.join(DOSSIER, f)

        # si le nom ne contient pas '@' → skip 7 lignes
        skiprows = 7 if "@" not in f else 0
        print(f"Lecture de {f} avec skiprows={skiprows}")

        df = pd.read_excel(path, skiprows=skiprows, engine="openpyxl",dtype=str)
        df.to_csv(outfile, index=False,header=first_file)
        first_file=False

print("Fusion terminée")

#print("Les deux exports ont été fusionnés avec succés")

#charger fichier Compacté
df=pd.read_csv(OUTPUT, sep=",")
#print ("Source chargée  avec succés")

#Supprimer colonnes sans utilité
df=df.drop(columns=["DNET ID","Inner VLAN ID","Max. Learnable MAC Addresses","Service Type"])

#print("Colonnes non utilisés supprimés avc succés")


#Creer Colonne 3 a partie de Colonne Name


VLAN =df["Name"].str.replace(
   
    r"/Multi(.*)","",regex=True).str.replace(
       
        r"/Single(.*)","",regex=True)


index_insertion1= df.columns.get_loc("Name")+1

df.insert(index_insertion1,"VLAN",VLAN)

#print ("Changements dans Frame faits avec succés")


#Creer colonne MSAN
MSAN= "MA5800-X17"
index_insertion3= df.columns.get_loc("VLAN") +1

df.insert(index_insertion3,"MSAN",MSAN)

#print ("Colonne MSAN ajoutée avec succés")

#Creation de Frame

    #Creation colonne intermediaire

Frame =df["Interface Information"].str.replace(
   
    r"ONTID","Ont",regex=True).str.replace(
       
        r"/GEM Port(.*)","",regex=True)
          

index_insertion2= df.columns.get_loc("Interface Information")+1

df.insert(index_insertion2,"Frame",Frame)

df=df.drop(columns=["Interface Information"])


#print ("Changements dans Frame faits avec succés")



#Creation de ID
df["ID"]= df["VLAN"]+r"v"+df["MSAN"]+r"v"+df["Device Name"]
#print("Creation de ID")

#Creation de SP
df["SP"]= df["Frame"]+r"v"+df["MSAN"]+r"v"+df["Device Name"]
#print("Creation de SP")


#sauvegarder dans un fichier excel
df.to_csv(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ExportSP.csv",index=False)
#print("fichier Export Service Port sauvegardé avec succés")

print("✅ExportServicePort crée avec succés")
