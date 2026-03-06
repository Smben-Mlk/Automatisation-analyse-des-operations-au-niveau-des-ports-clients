# Automatisation-des-analyses-des-op-rations-au-niveau-des-ports-clients

## Objectif

Ce projet a pour objectif de s’assurer que:
toutes les opérations sur les ports se font sur demande formelle;
toutes les configurations de débits sont conformes aux abonnements des clients (anti fraude).

## Outils

-Python (pandas numpy matplolib)

-Power BI

-Excel

## Dataset

##Inputs

-Exports NCE ( fichiers compressés qui contiennent les ports NCE et toutes les informations relatives à ces derniers.)

-Log NCE (contient toutes les opérations effectuées sur NCE)
			format: fichier compréssé(.zip)
-Parc (contient tous les clients fixes et leur abonnement.)
			format: fichier texte (.txt)
-Références Débits (contient les débits up et down correspondant à chaque abonnement)
			format: fichier excel(.xlsx)
-Export Kibaru (contient toutes les signalisations relatives aux ND)
			format: fichier excel (.xlsx)
			période de couverture: du lundi au dimanche (7 jours )
-Login NCE(contient tous les utilisateurs NCE et leurs structures respectives)
		format: fichier excel (.xlsx)
		à fiabiliser au besoin

##Outputs

- CONTRÔLE XDSL NCE
Croisement entre  login NCE, l’export Kibaru, le PARC ,les références de débit, le log XDSL et l’export XDSL qui nous permet de vérifier les actions au niveaux des équipements ADSL/VDSL ainsi que les configurations de débit.
feuille 1(DESC)
Contient toutes les opérations XDSL DESC avec leurs cases respectives.

feuille 2(A JUSTIF)
Contient toutes les opérations XDSL DESC avec sans cases.
.
feuille 3(CHANGEMENTS DE DÉBITS)
Contient toutes les opérations de changement de débit déjà contrôlées.

-CONTRÔLE CHANGEMENT DE DÉBIT
feuille 1(DESC)
Contient toutes les opérations FIBRE DESC avec leurs cases respectives.

feuille 2(A JUSTIF)
Contient toutes les opérations FIBRE DESC avec sans cases.

feuille 3(CONFIGURATIONS)
Contient toutes les opérations de configurations TV VOIP INTERNET.

feuille 4(CHANGEMENTS DE DÉBITS)
Contient toutes les opérations de changement de débit déjà contrôlées.

-CONTRÔLE ACTIVATION SUPPRESSION DE PORT
feuille 1(DESC)
Contient toutes les opérations d’activation/désactivation de port  DESC avec leurs cases respectives.

feuille 2(A JUSTIF)
Contient les opérations d’activation/desactivation de port  DESC sans cases.

## Aproche d'analyse

Les grandes lignes de la méthodologie adoptée dans le contrôle des débits sont:
Création de fichiers de référence qui attribuent à chaque abonnement du parc les débits de référence respectifs en MEGA.

Création d’un fichier qui contient toutes les possibilités de débit dans les exports NCE et leurs équivalences en MEGA.
   
Croisement de chaque action de changement de débit avec chacun des fichiers pour pour pouvoir comparer les débits en up et down.
 
Possibilités de CONSTATS
DEBIT OK si REF UP =DB UP et REF DOWN = DB DOWN
DEBIT NOK si REF UP ≠ DB UP et REF DOWN = DB DOWN
DEBIT NOK si REF UP =DB UP et REF DOWN ≠DB DOWN
DEBIT NOK si REF UP ≠DB UP et REF DOWN ≠DB DOWN

## Dashboard




