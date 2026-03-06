# Changments par 
import pandas as pd
import re
import warnings
import glob

warnings.simplefilter("ignore")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"

dossier=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés"
#glob.glob(dossier + r"\operation-compte*.csv")[0]


print("⏳Traitement du fichier Export XDSL")
print("Veuillez patienter svp, les exports NCE sont lourds ")

#definir les sources et la destination

# fichier ADSL MA5600T
ADSL1 =glob.glob(dossier + r"\MA5600T_ADSL*.xlsx")[0]

# fichier ADSL MA5603T
ADSL2 =glob.glob(dossier + r"\MA5603T_ADSL*.xlsx")[0]

# fichier VDSL MA5600T
VDSL1 =glob.glob(dossier + r"\MA5600T_VDSL*.xlsx")[0]

# fichier VDSL MA5603T
VDSL2 =glob.glob(dossier + r"\MA5603T_VDSL*.xlsx")[0]

#fichier output
resultat=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Export XDSL.xlsx"

#charger les sources
df1=pd.read_excel(ADSL1,skiprows=7,header=0,engine="openpyxl")
df2=pd.read_excel(ADSL2,skiprows=7,header=0,engine="openpyxl")
df3=pd.read_excel(VDSL1,skiprows=7,header=0,engine="openpyxl")
df4=pd.read_excel(VDSL2,skiprows=7,header=0,engine="openpyxl")

#print("Tous les 4 fichiers ont été chargés avec succés")

#Creer colonne MSAN pour tous les 4 fichiers
MSAN1= "MA5600T"
MSAN2= "MA5603T"

#print("Creation des colonnes MSAN...")
#ADSLMA5600T
index_insertion1= df1.columns.get_loc("Name") +1
df1.insert(index_insertion1,"MSAN",MSAN1)

#print ("Colonne MSAN ajoutée avec succés")

#VDSLMA5600T
index_insertion3= df3.columns.get_loc("Name") +1
df3.insert(index_insertion3,"MSAN",MSAN1)

#print ("Colonne MSAN ajoutée avec succés")

#ADSLMA5603T
index_insertion2= df2.columns.get_loc("Name") +1
df2.insert(index_insertion2,"MSAN",MSAN2)

#print ("Colonne MSAN ajoutée avec succés")
#VDSLMA5603T
index_insertion4= df4.columns.get_loc("Name") +1
df4.insert(index_insertion4,"MSAN",MSAN2)

#print ("Colonne MSAN ajoutée avec succés")


#print("Paramétrage des fichiers VDSL...")
#Ajouter colonne vide aux fichiers VDSL
CV=""

index_insertion5= df3.columns.get_loc("Alarm Template") +1
df3.insert(index_insertion5,"Extended Profile",CV)

index_insertion6= df4.columns.get_loc("Alarm Template") +1
df4.insert(index_insertion6,"Extended Profile",CV)

#Fusion des fichiers VDSL

#print("Fusion des fichiers VDSL")

fusion1=pd.concat([df3,df4],ignore_index=True)

#Renommage des colonnes VDSL

fusion1.rename(columns={"Line Template":"Line Profile","Alarm Template":"Alarm Profile","Type":"Port Type"},inplace=True)
print("Fusion des fichiers ADSL VDSL")


#fusion1.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\VDSL.xlsx",index=False)
#fusionner les  fichiers
fusion2=pd.concat([df1,df2,fusion1],ignore_index=True)

#print("Les quatres exports ont été fusionnés avec succés")

#charger fichier Compacté
#df=pd.read_excel(fusion2)
#print ("Source chargée  avec succés")


#Creation de ID
fusion2["ID"]= fusion2["Name"]+r"v"+fusion2["MSAN"]+r"v"+fusion2["Device Name"]
#print("Creation de ID")


fusion2.to_excel(resultat,index=False)

#print("fichier Export XDSL sauvegardé avec succés")

print("✅ExportXDSL crée avec succés")
