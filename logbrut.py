import pandas as pd 
import warnings
import glob
import chardet
import os

warnings.simplefilter("ignore")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"

dossier=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Fichiers réparés\Fichiers dezip"
#glob.glob(dossier + r"\operation-compte*.csv")[0]


log=glob.glob(dossier + r"\OperationLog*.csv")[0]

with open(log, "rb") as f:
    raw_data=f.read(100000)
result= chardet.detect(raw_data)
print(result["encoding"])

df=pd.read_csv(log, sep=",")

df[['Msan Model','Msan Name','Msan Ip Address']]=df['Operation Object'].str.split(':', expand=True)

df= df.drop(columns=['Operation Object'])

df1=df[df['Result']=='Successful']

df1=df1.rename(columns={'Operation':'Operation Name','Level':'Risk Level','Time':'Operation Time','Result':'Operation Result','Source':'Operation Category','Terminal IP Address':'Operation Terminal'})


df1=df1[['Operation Name','Risk Level','Operator','Operation Time','Operation Category','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details']]

df1.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\LogNCE.xlsx",index=False)


df2=df1[df1['Operation Category'].str.contains(r"MA58&EA58", na=False,case=False) & df1['Details'].str.contains(r"/0_", na=False,case=False)] 

print("Filtrage de log Service Port")

#df2.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\LogServicePortbrut.xlsx",index=False)

#SERVICE PORT
print("⏳Traitement de Log Service Port")

#creer nvlle colonne Frame

#rempacements
df2["Frame"]= df2["Details"].str.replace(
    r".*Service Port:","",regex=True).str.replace(
        r".*Name:","",regex=True).str.replace(
            r"/Multi-Service(.*)","",regex=True).str.replace(
                r"/Single(.*)","",regex=True).str.replace(
                    r".*VLAN ID:","",regex=True)


#Creation de ID
df2["ID"]= df2["Frame"]+r"v"+df2["Msan Model"]+r"v"+df2["Msan Name"]
print("Creation de ID")

df2['ID']=df2['ID'].str.replace(
        
            r" ","")

#sauvegarder dans un fichier excel
df2.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Log_SP.xlsx",index=False)

print("✅fichier Log Service Port traité avec succes")

df3=df1[df1['Operation Category'].str.contains(r"MA58&EA58", na=False,case=False) & ~df1['Details'].str.contains(r"/0", na=False,case=False)] 

print("Filtrage de log GPON")

#df3.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\LogGPONbrut.xlsx",index=False)

#GPON

print("⏳Traitement de LogGPON")

#rempacements
df3["Frame"]= df3["Details"].str.replace(

    r".*Frame","Frame",regex=True).str.replace(
        
        r"FrameFrameFrame","Frame",regex=True).str.replace(
            
            r"Frame0","Frame:0",regex=True).str.replace(
                
                r"/Slot","/Slot:",regex=True).str.replace(
                    
                    r"/Port","/Port:",regex=True).str.replace(
                        
                        r"/OnuID","/Ont:",regex=True).str.replace(
                            
                            r",OntID","/Ont",regex=True).str.replace(
                                
                                r",Slot","/Slot",regex=True).str.replace(
                                    
                                    r",Port","/Port",regex=True).str.replace(
                                        
                                        r",(.*)","",regex=True).str.replace(
                                            
                                            r" xPON(.*)","",regex=True).str.replace(
                                                
                                                r"\.","",regex=True).str.replace(
                                                    
                                                    r" ","",regex=True).str.replace(
                                                        
                                                        r"::",":",regex=True)


#print("changement fait avec succés")

#Creation de ID
df3["ID"]= df3["Frame"]+r"v"+df3["Msan Model"]+r"v"+df3["Msan Name"]
#print("Creation de ID")

df3['ID']=df3['ID'].str.replace(
        
            r" ","")


#sauvegarder dans un fichier excel
df3.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Log_GPON.xlsx",index=False)

print("✅fichier LogGPON créé avec succes")

df4=df1[df1['Operation Category'].str.contains(r"MA5600T", na=False,case=False) | df1['Operation Category'].str.contains(r"MA5603T", na=False,case=False)] 

print("Filtrage de log XDSL")

#df4.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\LogXDSLbrut.xlsx",index=False)
#XDSL

print("⏳Traitement de logXDSL")

#rempacements
df4["Frame"]= df4["Details"].str.replace(

    r".*Frame","Frame",regex=True).str.replace(
        
        r".*/0_","Frame:0,Slot:",regex=True).str.replace(
            
            r"_",",Port:",regex=True).str.replace(
                
                r"/Slot",",Slot",regex=True).str.replace(
                    
                    r"/Port",",Port",regex=True).str.replace(
                        
                        r"/(.*)","",regex=True).str.replace(
                            
                            r" Line(.*)","",regex=True).str.replace(

                                r" Bind(.*)","",regex=True).str.replace(

                                    r" Extend(.*)","",regex=True).str.replace(

                                        r" Port Admin(.*)","",regex=True).str.replace(

                                            r" AdminStatus(.*)","",regex=True).str.replace(
                                
                                                r",Slot","/Slot",regex=True).str.replace(
                                                    
                                                    r",Port","/Port",regex=True).str.replace(
                                                        
                                                        r",(.*)","",regex=True).str.replace(
                                                            
                                                            r"Frame:0/Slot:Frame","Frame",regex=True).str.replace(
                                                                
                                                                r"\.","",regex=True).str.replace(
                                                                    
                                                                    r" ","",regex=True)

#print("changements effectués avec succés")

#Creation de ID
df4["ID"]= df4["Frame"]+r"v"+df4["Msan Model"]+r"v"+df4["Msan Name"]
#print("Creation de ID")

df4['ID']=df4['ID'].str.replace(
        
            r" ","")


#sauvegarder dans un fichier excel
df4.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Log_XDSL.xlsx",index=False)

print("✅fichier Log XDSL créé avec succes")

