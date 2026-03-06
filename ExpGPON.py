# Changments par 
import pandas as pd
import re
import warnings
import glob
from openpyxl import load_workbook
import zipfile
import io
import os
import win32com.client as win32
import glob

print("⏳Traitement du fichier Export GPON")
print("Veuillez patienter svp, les exports NCE sont lourds ")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"

dossier=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés"
dest =chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ExportGPON.xlsx"

#glob.glob(dossier + r"\operation-compte*.csv")[0]

#definir la source et la destination
#source =glob.glob(dossier + r"\All_GPON_ONU*.xlsx")[0]


input_file = glob.glob(dossier + r"\All_GPON_ONU*.xlsx")[0]

df_source=pd.read_excel(input_file,skiprows=7,header=0,engine="openpyxl")

#print(df_source.head())

#df_source=pd.read_excel(source,skiprows=7,header=0,engine="openpyxl")

#Selectionner colons à copier
colonnes=["Device Name","Name","Alias","SN"]
df_selection= df_source[colonnes]

#enregistrer colonnes slectionnés dans destination
df_selection.to_excel(dest,index=False)
#print(f"Les colonnes {colonnes} ont étés copiés dans le fichier excel {dest}")

#charger destnation
df=pd.read_excel(dest)
#print ("Destination crée  avec succés")

#Creer Colonne 3 a partie de Colone Name
nouvelle_colonne= df['Name']
index_insertion1= df.columns.get_loc("Name") +1

df.insert(index_insertion1,"Frame1",nouvelle_colonne)

#print ("Colonne Frame ajoutée avec succés")


#Changements pour avoir Frame

Frame =df["Frame1"].str.replace(
   
    r".*Frame0","Frame:0",regex=True).str.replace(
       
        r"OnuID","Ont:",regex=True).str.replace(
       
            r"Slot","Slot:",regex=True).str.replace(
       
                r"Port","Port:",regex=True).str.replace(
       
                    r".*VLAN ID:","",regex=True)

index_insertion2= df.columns.get_loc("Frame1")+1

df.insert(index_insertion2,"Frame",Frame)

#print ("Changements dans Frame faits avec succés")

#Supprimer Colonne Frame1 intermediaire pour avoir Frame

df=df.drop(columns=["Frame1","Name"])

#print ("Colonnes intermediares supprimés avec succés")

#Creer colonne MSAN
MSAN= "MA5800-X17"
index_insertion3= df.columns.get_loc("Frame") +1

df.insert(index_insertion3,"MSAN",MSAN)

#print ("Colonne MSAN ajoutée avec succés")


#Creation de ID
df["ID"]= df["Frame"]+r"v"+df["MSAN"]+r"v"+df["Device Name"]
#print("Creation de ID")


#sauvegarder dans un fichier excel
df.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ExportGPON.xlsx",index=False)
#print("fichier ExportGPON sauvegardé avec succés")

print("✅ExportGPON crée avec succés")
