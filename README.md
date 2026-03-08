# Automatisation-des-analyses-des-operations-au-niveau-des-ports-clients

## Objectif

Ce projet a pour objectifs de:

- Identifier toutes les opérations faites au niveau des ports clients
- Identfier les auteurs et leurs droits (fiabilisation)
- Corréler les ports aux infos (numéros, configurations de debit, TV , VOIP ) 
- s’assurer que toutes les opérations sur les ports se font sur demande formelle;
- s’assurer que toutes les configurations de débits sont conformes aux abonnements des clients (anti fraude).

[Voir rapport complet] (https://github.com/Smben-Mlk/Automatisation-des-analyses-des-op-rations-au-niveau-des-ports-clients/blob/main/RAPPORT%20AUTOMATISTION%20DES%20OP%C3%89RATIONS%20NCE.pdf)

## Outils

- Python (pandas numpy matplolib)

- Power BI

- Excel

- SQL

## Compétences

- analyse de logs
- Deep learning
- Machine learning
- automatisation
- Analyse de données
- Netoyage de données
- aggrégation de données
- corrélation de données

## Dataset

## Inputs

- Exports NCE ( fichiers compressés qui contiennent les ports NCE et toutes les informations relatives à ces derniers.)

- Log NCE (contient toutes les opérations effectuées sur NCE)
			format: fichier compréssé(.zip)
- Parc (contient tous les clients fixes et leur abonnement.)
			format: fichier texte (.txt)
- Références Débits (contient les débits up et down correspondant à chaque abonnement)
			format: fichier excel(.xlsx)
- Export Kibaru (contient toutes les signalisations relatives aux ND)
			format: fichier excel (.xlsx)
			période de couverture: du lundi au dimanche (7 jours )
- Login NCE(contient tous les utilisateurs NCE et leurs structures respectives)
		format: fichier excel (.xlsx)
		à fiabiliser au besoin

## Outputs

- CONTRÔLE XDSL NCE
  
Croisement entre  login NCE, l’export Kibaru, le PARC ,les références de débit, le log XDSL et l’export XDSL qui nous permet de vérifier les actions au niveaux des équipements ADSL/VDSL ainsi que les configurations de débit.

feuille 1

Contient toutes les opérations XDSL DESC avec leurs cases respectives.

feuille 2

Contient toutes les opérations XDSL DESC avec sans cases.

feuille 3

Contient toutes les opérations de changement de débit déjà contrôlées.

- CONTRÔLE CHANGEMENT DE DÉBIT

feuille 1

Contient toutes les opérations FIBRE DESC avec leurs cases respectives.

feuille 2

Contient toutes les opérations FIBRE DESC avec sans cases.

feuille 3

Contient toutes les opérations de configurations TV VOIP INTERNET.

feuille 4

Contient toutes les opérations de changement de débit déjà contrôlées.

- CONTRÔLE ACTIVATION SUPPRESSION DE PORT

feuille 1

Contient toutes les opérations d’activation/désactivation de port  DESC avec leurs cases respectives.

feuille 2

Contient les opérations d’activation/desactivation de port  DESC sans cases.

## Approche d'analyse

Le script principal de traitement des données NCE nce.py est composé de plusieurs sous-scripts appelés à l’aide de subprocess.run. Chacun de ces scripts a un objectif bien défini:
- reparation.py (dézzipe et répare les exports)
- parc.py (convertit le parc en fichier excel exploitable)
- logbrut.py (départage et filtre le log brut en 3 parties)
  
		Log XDSL
Contient toutes les opérations ADSL et VDSL effectuées.
Le script logbrut.py filtre sur les équipements MA5600T et MA5603T puis fait les changements nécessaires pour créer la clef de croisement ID.

		Log GPON
Contient toutes les opérations d’activation de suppression effectuées sur les équipements fibre.
Le script logbrut.py filtre sur l’équipement MA5800-X17 et détails ne contient pas /0 puis fait les changements nécessaires pour créer la clef de croisement ID.

 		Log SP
Contient toutes les opérations de modification effectuées sur les équipements fibre.
Le script logbrut.py filtre sur l’équipement MA5800-X17 et details contient /0_ puis fait les changements nécessaires pour créer la clef de croisement ID.


- ExpXDSL.py (créait l’export Export XDSL)
Le script Expxdsl.py fusionne les 4 exports MA5600TADSL, MA5603TADSL,  MA5600TVDSL et  MA5603TVDSL puis fait les changements nécessaires pour créer la clef de croisement ID.

- ExpGPON.py (créait l’export Export GPON)
Le script ExpGPON.py traite l’export ALL_GPON_ONU puis fait les changements nécessaires pour créer la clef de croisement ID.

- ExpSP.py (créait l’export Export Service Port)
Le script ExpSP.py fusionne les exports MA5800-X17  puis fait les changements nécessaires pour créer les clefs de croisement ID et SP.

- main.py (créait les fichiers de controle)
Le script main.py génére automatiquement 3 fichiers de contrôle en format xlsx:

- Vérification sur les debits 
Les grandes lignes de la méthodologie adoptée dans le contrôle des débits sont:
Création de fichiers de référence qui attribuent à chaque abonnement du parc les débits de référence respectifs en MEGA.

Création d’un fichier qui contient toutes les possibilités de débit dans les exports NCE et leurs équivalences en MEGA.
   
Croisement de chaque action de changement de débit avec chacun des fichiers pour pour pouvoir comparer les débits en up et down.
 
Possibilités :

- DEBIT OK si REF UP =DB UP et REF DOWN = DB DOWN

- DEBIT NOK si REF UP ≠ DB UP et REF DOWN = DB DOWN

- DEBIT NOK si REF UP =DB UP et REF DOWN ≠DB DOWN

- DEBIT NOK si REF UP ≠DB UP et REF DOWN ≠DB DOWN


