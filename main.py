import pandas as pd
import warnings
import glob

warnings.simplefilter("ignore")

chemin=r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO"


dossier=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\PARC ET CASE"

out=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS"
int=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\\Intermediaires"
db=chemin+r"\TRAITEMENT INPUTS ADSL\DEBITS"

#glob.glob(db + r"

LogGPON= glob.glob(int + r"\Log_GPON.xlsx")[0]
ExpGPON= glob.glob(int + r"\ExportGPON.xlsx")[0]
ExpSP= glob.glob(int + r"\ExportSP.csv")[0]
login=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\LOGIN\login.xlsx"
Case=glob.glob(dossier + r"\Vue Recherche avancée Cases*.xlsx")[0]
PastExp=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\EXPORTS S-1\ExportGPON.xlsx"
LogSP= glob.glob(int + r"\Log_SP.xlsx")[0]
Parc=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\PARC.xlsx"
UP=glob.glob(db + r"\DB UP.xlsx")[0]
DOWN=glob.glob(db + r"\DB DOWN.xlsx")[0]
REF=glob.glob(db + r"\REF DB FIBRE.xlsx")[0]
LogX= glob.glob(int + r"\Log_XDSL.xlsx")[0]
ExpX= glob.glob(int + r"\Export XDSL.xlsx")[0]
DB=glob.glob(db + r"\DB ADSL.xlsx")[0]
REFADSL=glob.glob(db + r"\REF DB ADSL.xlsx")[0]

print ("⏳Chargement et aggrégation des Logs, Exports, Case,login et Parc...")

#Chargement ds fichiers
dfGPON= pd.read_excel(LogGPON)
df2= pd.read_csv(ExpSP,sep=",")
df3=pd.read_excel(ExpGPON)
df4=pd.read_excel(login)
df6=pd.read_excel(Case)
dfSP= pd.read_excel(LogSP)
df5=pd.read_excel(Parc)
dfUP= pd.read_excel(UP)
dfDOWN= pd.read_excel(DOWN)
dfREF= pd.read_excel(REF)
dfXDSL= pd.read_excel(LogX)
df1= pd.read_excel(ExpX)
dfDB= pd.read_excel(DB)
dfREFXDSL= pd.read_excel(REFADSL)

#XDSL

df6=df6.drop_duplicates(subset=['Lignes du clients'])

df6['Lignes du clients']=pd.to_numeric(df6['Lignes du clients'],errors="coerce")

df6=df6[['Numéro du case','Lignes du clients']]

df6=df6.rename(columns={'Lignes du clients':'Alias','Numéro du case':'Case'})

df6=df6.dropna(subset=['Alias'])

#print("Croisement Log et Export ...")

#Croisement par ID

merged= dfXDSL.merge(df1, on='ID',how='left')

resultmerged=merged[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','Alias','Line Profile']]

#print("Croisement avec login ...")

#Croisement avec le login

df4['Operator']=df4['Operator'].astype(str).str.strip().str.lower()

resultmerged['Operator']=resultmerged['Operator'].astype(str).str.strip().str.lower()

Loginmerged=resultmerged.merge(df4, on='Operator', how='left')

Rloginmerged=Loginmerged[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','Alias','Line Profile']]

#print("Creation de actions XDSL")

Rloginmerged.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ActionsXDSL.xlsx",index=False)

print("Creation de CONTROLE XDSL DESC")


DESC=Rloginmerged[Rloginmerged['Structure']!='HORS DESC']

#DESC.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Actions XDSL DESC.xlsx",index=False)

#RECUPERER CEUX SANS ND

DESCNOND= DESC[~DESC['Alias'].astype(str).str.startswith(('338','339'))]


#CROISEMENT AVEC EXPORT S-1

ExpXS1=chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\EXPORTS S-1\Export XDSL.xlsx"

df7=pd.read_excel(ExpXS1)

df7=df7[['ID','Alias']]

DESCNOND.drop(columns=['Alias'], inplace=True)

NOND= DESCNOND.merge(df7, on='ID',how='left')

resultNOND=NOND[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','Alias','Line Profile']]

resultNOND['Alias']=pd.to_numeric(resultNOND['Alias'],errors="coerce")

#resultNOND.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\resultNOND.xlsx",index=False)


DESCNOND2= resultNOND[~resultNOND['Alias'].astype(str).str.startswith(('338','339'))]

DESCNOND2['Case']=""

#DESCNOND2.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\NOND.xlsx",index=False)

#DESCNOND.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\PYTHONTEST\NCE\output2\CTRLE XDSL DESC PARC sans ND.xlsx",index=False)


#Supprimer les incoherences

#df4['ND']=pd.to_numeric(df4['ND'],errors="coerce")


DESC['Alias']=pd.to_numeric(DESC['Alias'],errors="coerce")

df6['Alias']=pd.to_numeric(df6['Alias'],errors="coerce")


#Renommer colonne ND de PARC
df5=df5.rename(columns={'ND':'Alias'})

df9=df5[['Alias','ACCESADSL']]

#FILTER  CEUX AVEC ND CONCATENER AVEC PARC 
DESCND3=DESC.dropna(subset=['Alias'])

DESCND1=pd.concat([DESCND3,resultNOND], ignore_index=True)

#CONCATENER LES ACTIONS AVEC ND AVEC LE PARC

mergedParc1=DESCND1.merge(df9, on='Alias', how='left')

#Concatener avec Cases

#print("Croisement avec Cases")

df6=df6.drop_duplicates(subset=['Alias'])

mergedCases=mergedParc1.merge(df6, on='Alias', how='left')

#mergedParc1.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\PYTHONTEST\NCE\output2\CTRLE XDSL DESC PARC avec ND.xlsx",index=False)

#CONCATENER LES DEUX FICHIERS

ALL=pd.concat([mergedCases,DESCNOND2], ignore_index=True)

ALL=ALL.drop_duplicates(subset=['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','Alias','Line Profile','Case'])

#ALL.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE XDSL DESC.xlsx",index=False)

#ALL['Constat']=""

nonconf=ALL[ALL['Case'].isna()]

nonconf1=pd.concat([nonconf,resultNOND],ignore_index=True)

modif=ALL[ALL["Operation Name"].str.startswith(('Modify','Configure'), na =False)]

#DEBITS

modif1=modif.merge(dfDB, on='Line Profile', how='left')

#REF
modif3=modif1.merge(dfREFXDSL, on='ACCESADSL', how='left')

modif3["CONSTAT"]= modif3.apply(lambda row: "OK" if row["REF DEBIT"] == row["DEBIT LIGNE"] else "NOK", axis=1)

NOK=modif3[modif3["CONSTAT"]== "NOK"]

#modif3=modif3[['STRUCTURE','User','Event Time','Session','Application','Operation','NE/Agent','Object','Arguments','Result','ID','ID DEBIT','SERIAL_NUMBER','ND','DB UP VOIP','DB UP TV','DB DOWN TV','ACCESFIBRE','CASE','DEBIT UP','DEBIT DOWN','DB UP','REF UP','CONSTAT UP','DB DOWN','REF DOWN','CONSTAT DOWN','CONSTAT']]


modif3=modif3[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','Alias','Case','ACCESADSL','Line Profile','REF DEBIT','DEBIT LIGNE','CONSTAT']]


with pd.ExcelWriter(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE XDSL DESC.xlsx",engine="xlsxwriter") as writer:

    ALL.to_excel(writer, sheet_name="DESC",index=False)

    nonconf1.to_excel(writer, sheet_name="A JUSTIFS",index=False)

    modif3.to_excel(writer, sheet_name="CHANGEMENTS DE DB",index=False)

print ("✅Controle XDSL effectué avec succés")


print("Nombre d'opérations XDSL NCE   = ",len(dfXDSL))
print("Nombre d'opérations XDSL NCE DESC = ",len(ALL))
print("Nombre de changements de débit XDSL NCE  = ",len(modif3))
print("Nombre de changements de débits suspects XDSL NCE = ",len(NOK))

#SP

#Croisement par AVEC EXPORT SP POUR AVOIR INFOSPORT ET DB

merged= dfSP.merge(df2, on='ID',how='left')

merged=merged.rename(columns={'Frame_x':'Frame'})


merged1=merged[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','SP','Upstream Traffic Profile','Downstream Traffic Profile']]
#CROISEMENT AVEC EXPORTGPON POUR AVOIR ND

df11=df3

df11=df11.rename(columns={'ID':'SP'})

df11=df11[['SP','Alias']]

mergeGPON= merged1.merge(df11, on='SP',how='left')

mergeGPON=mergeGPON[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','SP','Alias','Upstream Traffic Profile','Downstream Traffic Profile']]

mergeGPON['Alias']=pd.to_numeric(mergeGPON['Alias'],errors="coerce")

#mergeGPON.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\testerror.xlsx",index=False)

#mergeGPON=mergeGPON.rename(columns={'Frame_x':'Frame'})

#Croisement par SP
#df2=df2[['SP','Alias']]

#merged1=resultmerged.merge(df2, on='SP',how='left' )

#merged1=merged1[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID_x','SP','Alias','Upstream Traffic Profile','Downstream Traffic Profile']]

#Croisement avec Login
#df3=df3.rename(columns={'ND':'Alias'})

df4['Operator']=df4['Operator'].astype(str).str.strip().str.lower()

mergeGPON['Operator']=mergeGPON['Operator'].astype(str).str.strip().str.lower()


Loginmerged= mergeGPON.merge(df4,on='Operator',how='left')

resultLoginmerged=Loginmerged[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','SP','Alias','Upstream Traffic Profile','Downstream Traffic Profile']]


resultLoginmerged.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ActionsSP.xlsx",index=False)

#print ("Actions Service Port créé avec succés")

#Filter sur DESC

print("Creation de CONTROLE SP DESC")


DESC=resultLoginmerged[resultLoginmerged['Structure']!='HORS DESC']

#DESC.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Actions SP DESC.xlsx",index=False)

#RECUPERER CEUX SANS ND

DESCNOND= DESC[~DESC['Alias'].astype(str).str.startswith(('338','339'))]

DESCNOND['ACCESFIBRE']=""
DESCNOND['Case']=""

#DESCNOND.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\CTRLE SP DESC  sans ND.xlsx",index=False)

#Supprimer les incoherences

#df5['ND']=pd.to_numeric(df5['ND'],errors="coerce")

DESC['Alias']=pd.to_numeric(DESC['Alias'],errors="coerce")

#print("Verification des types des 2 Colonnes à croiser")

#print(df5['ND'].dtype)
#print(DESC['Alias'].dtype)
#print(df6['Alias'].dtype)

#print("Croisement avec Parc")

#Renommer colonne ND de PARC

df5=df5.rename(columns={'ND':'Alias'})

df5=df5[['Alias','ACCESFIBRE']]

#FILTER  CEUX AVEC ND CONCATENER AVEC PARC
DESCND=DESC.dropna(subset=['Alias'])

#CONCATENER LES ACTIONS AVEC ND AVEC LE PARC

mergedParc1=DESCND.merge(df5, on='Alias', how='left')

#mergedParc1.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\CTRLE SP PARC avec ND.xlsx",index=False)

#Concatener avec Cases

#print("Croisement avec Cases")

mergedCases=mergedParc1.merge(df6, on='Alias', how='left')


#CONCATENER LES DEUX FICHIERS

ALL=pd.concat([mergedCases,DESCNOND], ignore_index=True)

#ALL.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE SP DESC.xlsx",index=False)

ALL['Constat']=""

nonconf=ALL[ALL['Case'].isna()]

nonconf1=pd.concat([nonconf,DESCNOND],ignore_index=True)

#CHANGEMENTS DE DB
modif5=ALL[ALL["Operation Name"].str.startswith(('Modify','Create'), na =False)]

modif=modif5[modif5["ID"].str.startswith(('50/','45/'), na =False)]

#modif7=ALL[ALL["Operation Name"].str.startswith(('Create'), na =False)]

#modif=pd.concat([modif6,modif7],ignore_index=True)

#DEBITS

dfUP2=dfUP.rename(columns={'DEBIT':'Upstream Traffic Profile'})

modif1=modif.merge(dfUP2, on='Upstream Traffic Profile', how='left')

dfDOWN2=dfDOWN.rename(columns={'DEBIT':'Downstream Traffic Profile'})

modif2=modif1.merge(dfDOWN2, on='Downstream Traffic Profile', how='left')

#REF

modif3=modif2.merge(dfREF, on='ACCESFIBRE', how='left')


modif3["CONSTAT UP"]= modif3.apply(lambda row: "OK" if row["REF UP"] == row["DB UP"] else "NOK", axis=1)

modif3["CONSTAT DOWN"]= modif3.apply(lambda row: "OK" if row["REF DOWN"] == row["DB DOWN"] else "NOK", axis=1)

modif3["CONSTAT"]= modif3.apply(lambda row: "DEBIT OK" if row["CONSTAT DOWN"] == "OK" and row["CONSTAT UP"] == "OK" else "DEBIT NOK", axis=1)

#modif3=modif3[['STRUCTURE','User','Event Time','Session','Application','Operation','NE/Agent','Object','Arguments','Result','ID','ID DEBIT','SERIAL_NUMBER','ND','DB UP VOIP','DB UP TV','DB DOWN TV','ACCESFIBRE','CASE','DEBIT UP','DEBIT DOWN','DB UP','REF UP','CONSTAT UP','DB DOWN','REF DOWN','CONSTAT DOWN','CONSTAT']]


modif3=modif3[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame','ID','SP','Alias','ACCESFIBRE','Case','Upstream Traffic Profile','Downstream Traffic Profile','DB UP','REF UP','CONSTAT UP','DB DOWN','REF DOWN','CONSTAT DOWN','CONSTAT']]

NOK2=modif3[modif3["CONSTAT"]== "DEBIT NOK"]


with pd.ExcelWriter(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE CHANGEMENT DE DEBIT DESC.xlsx",engine="xlsxwriter") as writer:

    ALL.to_excel(writer, sheet_name="DESC",index=False)

    nonconf1.to_excel(writer, sheet_name="A JUSTIFS",index=False)

    modif5.to_excel(writer, sheet_name="CONFIGURATIONS",index=False)

    modif3.to_excel(writer, sheet_name="CHANGEMENTS DE DB",index=False)


print ("✅Controle Service Port effectué avec succés")

print("Nombre d'opérations Service Port   = ",len(dfSP))
print("Nombre d'opérations Service Port DESC = ",len(ALL))
print("Nombre de changements de débit FIBRE NCE  = ",len(modif3))
print("Nombre de changements de débits FIBRE suspects = ",len(NOK2))

#GPON

df12=df3

#Creation de la colonne SN pour les cas de suppression de port

#Pour LGPON
dfGPON['SN']=dfGPON['Details'].str.extract(r'SN:\s*(.{1,16})|Sn:->\s*(.{1,16})').bfill(axis=1).iloc[:,0]

#print("Colonne SN créé dans LogGPON avec succés")

#Pour Export GPON
df12['SN']=df12['SN'].str[:16]

#df3

#print("Colonne SN créé dans ExportGPON avec succés")

#Pour Export GPON S-1
df7=pd.read_excel(PastExp)

df7['SN']=df7['SN'].str[:16]

#Croisement des fichiers

#Croisement par ID

merged= dfGPON.merge(df12, on='ID',how='left')

resultmerged=merged[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID','SN_x','Alias']]

#Filtrage des colonnes sans ND pui croisement par SN

#Colonnes sans ND
dfSN=resultmerged[resultmerged['Alias'].isna()]

dfSN=dfSN.rename(columns={'SN_x':'SN'})

#Croisement par SN
mergedSN= dfSN.merge(df7,on='ID',how='left')

resultmergedSN=mergedSN[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID','SN_x','Alias_y']]

#resultmergedSN.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\SNtest.xlsx",index=False)

#Recherche des ND restants dans details apres Regitration ID:

resultmergedSN1=resultmergedSN[resultmergedSN['Alias_y'].isna()]

resultmergedSN1['Alias_y']=resultmergedSN1['Alias_y'].fillna(resultmergedSN1['Details'])

resultmergedSN1['Alias_y']=resultmergedSN1['Alias_y'].str.extract(r'Registration ID:\s*(.{1,9})|Alias:->\s*((?:\d+\s*){8})').bfill(axis=1).iloc[:,0]

#Renommer les colonnes pour pouvoir concatener

resultmergedSN=resultmergedSN.dropna(subset=['Alias_y'])

resultmergedSN2=pd.concat([resultmergedSN,resultmergedSN1],ignore_index=True)

resultmergedSN2=resultmergedSN2.rename(columns={'Frame_x':'Frame','ID_x':'ID','Alias_y':'Alias'})

#Supprimer les collones sans ND dans resultmerged

resultmerged=resultmerged.dropna(subset=['Alias'])

#Renommer les colonnes pour pouvoir concatener

resultmerged=resultmerged.rename(columns={'Frame_x':'Frame','SN_x':'SN'})

#Concatener les 2 fichiers 
result=pd.concat([resultmerged,resultmergedSN2],ignore_index=True)

#result.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\PYTHONTEST\NCE\output2\ActionsGPON1.xlsx",index=False)

#Croiser avec export service port pour avoir débits

mergedSP=result.merge(df2, on='Alias', how='left')

resultmergedSP=mergedSP[['Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID_x','SN','Alias']]

#resultmergedSP.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\PYTHONTEST\NCE\output2\ActionsGPON2.xlsx",index=False)

#Croisement avec Login

df4['Operator']=df4['Operator'].astype(str).str.strip().str.lower()

resultmergedSP['Operator']=resultmergedSP['Operator'].astype(str).str.strip().str.lower()

mergedLogin= resultmergedSP.merge(df4,on='Operator',how='left')

resultmergedLogin=mergedLogin[['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID_x','SN','Alias']]

resultmergedLogin.to_excel(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\ActionsGPON.xlsx",index=False)

#Filter sur DESC

print("Creation de CONTROLE GPON DESC")

DESC=resultmergedLogin[resultmergedLogin['Structure']!='HORS DESC']

#DESC.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\Actions GPON DESC.xlsx",index=False)

#RECUPERER CEUX SANS ND

DESCNOND= DESC[~DESC['Alias'].astype(str).str.startswith(('338','339'))]

#DESCNOND['ACCES FIBRE']=""
DESCNOND['Case']=""

#DESCNOND.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\CTRLE GPON DESC PARC sans ND.xlsx",index=False)

#Supprimer les incoherences

#df5['ND']=pd.to_numeric(df5['ND'],errors="coerce")


DESC['Alias']=pd.to_numeric(DESC['Alias'],errors="coerce")


#print("Verification des types des 2 Colonnes à croiser")

#print(df5['ND'].dtype)
#print(DESC['Alias'].dtype)
#print(df6['Alias'].dtype)


#print("Croisement avec Parc")

#Renommer colonne ND de PARC
#df5=df5.rename(columns={'ND':'Alias'})

#df5=df5[['Alias','ACCESFIBRE']]

#FILTER  CEUX AVEC ND CONCATENER AVEC PARC 
DESCND=DESC.dropna(subset=['Alias'])

#CONCATENER LES ACTIONS AVEC ND AVEC LE PARC

#mergedParc1=DESCND.merge(df5, on='Alias', how='left')

#mergedParc1.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\Intermediaires\CTRLE GPON PARC avec ND.xlsx",index=False)

#Concatener avec Cases

#print("Croisement avec Cases")

mergedCases=DESCND.merge(df6, on='Alias', how='left')

mergedCases=mergedCases.drop_duplicates(subset=['Structure','Operation Name','Risk Level','Operator','Operation Time','Operation Terminal','Msan Model','Msan Name','Msan Ip Address','Operation Result','Details','Frame_x','ID_x','SN','Alias','Case'])

#CONCATENER LES DEUX FICHIERS

ALL=pd.concat([mergedCases,DESCNOND], ignore_index=True)

#ALL.to_excel(r"C:\Users\tmp_sembene76821\OneDrive - Orange Sonatel\Bureau\AUTO\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE GPON DESC.xlsx",index=False)

ALL['Constat']=""

nonconf=ALL[ALL['Case'].isna()]

nonconf1=pd.concat([nonconf,DESCNOND],ignore_index=True)

with pd.ExcelWriter(chemin+r"\TRAITEMENT INPUTS ADSL\CURRENT WEEK\NCE\OUTPUTS\CTRLE NCE\CTRLE ACTIVATION SUPPRESSION DESC.xlsx",engine="xlsxwriter") as writer:

    ALL.to_excel(writer, sheet_name="DESC",index=False)

    nonconf1.to_excel(writer, sheet_name="A JUSTIFS",index=False)


print ("✅Controle GPON effectué avec succés")

print("Nombre d'opérations GPON   = ",len(dfGPON))
print("Nombre d'opérations GPON DESC = ",len(ALL))