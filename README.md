 Log_purge
=========

Un script powershell  qui permet de gérer les vieux fichiers de log (ou autre)

Ce script permet de gerer la suppression et la compression des fichiers de log.
Il se base sur 7zip.exe pour compresser les fichiers et les répertoires mais il existe une fonction pour faire sans (j'ai cependant noté des problemes de fiabilité dans ce cas).

Principe de fonctionnement:

il faut d'abord créer un fichier de consigne qui indiquera ce qu'il faut faire.
ensuite il suffit de lancer le script! un fichier de log est automatiquement généré.

 les fonctions utlisable dans le fichier de conf sont:
-  compress    = > pour compresser les fichiers de log en fichier zip
-  compressDir = > pour compresser des repertoires (un fichier zip par repertoire)
-  purge      => pour supprimer les vieux fichiers
-  compacter  => pour compacter les fichier NTFS
 
 Les parametes de chaque fonction sont les mêmes:
 -  -rep  <repertoire> => le nom du repertoire qui contient les fichiers
 - -fichier <extention> | -expression <regexp> => le type de fichier à traiter ou une expression reguliere  
 - -retention <retention>  => la durée de retention souhaitée
 - -r    => pour activer la réccursivité
 
 le script accepte les parametres suivants:
 - version =>   affiche le numéro de version et se termine sans rien faire
 - createconf =>  créer un fichier de conf dans le repertoire work
 - createtask =>   créer une tache planifiée dans le dossier Claranet.   par defaut, elle tourne à minuit tous les jours
 


