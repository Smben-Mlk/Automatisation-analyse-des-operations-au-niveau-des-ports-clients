import pandas as pd 
import chardet
import warnings
import glob

warnings.simplefilter("ignore")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"


dossier=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\PARC ET CASE"

Parc=glob.glob(dossier + r"\Z_PARC_TECHFAM*.txt")[0]
               
#glob.glob(dossier + r"\operation-compte*.csv")[0]



with open(Parc, "rb") as f:
    raw_data=f.read(100000)
result= chardet.detect(raw_data)
#print(result["encoding"])

print("Chargement de parc")
df=pd.read_csv(Parc, sep= "|",encoding="ISO-8859-1",skiprows=1,on_bad_lines="skip")

df=df.drop(index=[0])

#print("Nommage des colonnes")


df.columns=df.columns.str.replace(" ","",regex=True)

#print("Selection des colonnes")


df=df[['NCLI','ND','ACCESRESEAU','ACCESFIBRE','ACCESADSL','BOUQUETTV']]


#print("Suppression des doublons")


df=df.drop_duplicates(subset=["ND"])

#print("Conversion de ND en float")

df['ND']=pd.to_numeric(df['ND'],errors="coerce")

df=df.dropna(subset=['ND'])


#print(df['ND'].dtype)


#print("Parc traité avec succés")


df.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\PARC.xlsx" ,index=False, engine="openpyxl")

print("✅Parc créé avec succés")