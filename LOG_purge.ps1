# -----------------------------------------------------------------
#
# 	Objet		LOG
# 	Fonction	purge les anciens logs
#
# 	Auteur 		Erwan Le Tiec
# 	Creation	09/04/2013
#   Modif       ajout des fonctions de compression
#               elt le 16/05: check sur la date de derniere ecriture au lieu de la date de creation
#               elt le 05/07/2013: corr bug du repertoire de log, modif affichage lors de la compression
#               elt le 04/10/2013: modif fonction pour ne pas compacter les fichiers deja compacté
#               elt le 12/11/2013: v1.4 - les purges spécifique au serveur sont maintenant dans work ajout d'un numero de version
#               elt le 14/11/2013: v1.5 - option de creation de la tache planifiée et du fichier local 
#               elt le 16/04/2014: v1.6 - exclude des fichiers zip lors d'une compression, test si lancement via jobscheduler
#               elt le 22/05/2014: v1.7 - correction d'un bug sur le exclude
#               elt le 30/10/2014: v1.8 - ajout de l'option expression pour gérer des regexp
#               elt le 13/11/2014: v1.9 - ajout d'une fonction de compression d'un repertoire entier
#               elt le 26/02/2015: v2.0 - modif pour integration dans la nouvelle arbo
#               elt le 10/11/2015: v2.1 - utilise 7z pour la compression des dir (pb de fiabilité sinon)
#               elt le 07/01/2016: v2.2 - elargir la fenetre pour l'affichage
#               elt le 22/01/2016: v2.3 - utlise 7z pour toutes les compressions
#               elt le 19/02/2016: v2.3.1 - ajout exit 99 si 7z non trouvé
#               elt le 29/06/2016: v2.3.2 - modif de code, suppression de la fonction fzip
#
# -----------------------------------------------------------------

# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------
[CmdletBinding()] 
param (
    [switch] $version= $false,
    [switch] $help=$false,
    [Switch] $createtask=$false,
    [switch] $createconf=$false
)

$scriptversion = "2.3.2"
$consignePurge= "LOG_purge.conf"
$7zPath="D:\sources\claranet\7z.exe"

# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$scriptName = $myInvocation.MyCommand.Name
$scriptshortname= $scriptName.substring(0,$scriptName.length-4)
$scriptLongName = $myInvocation.MyCommand.path
$Stamp=(get-date).ToString("yyMMdd_HHmm")
$ScriptPath=split-path $MyInvocation.MyCommand.Path -parent
$LogPath="D:\logs\claranet"
$workpath="D:\sources\client"



if ((Test-Path $LogPath) -eq $True)
    {$Logfile= "$LogPath\$scriptshortname-$stamp.log"
    }
else
    {write-debug "le repertoire $LogPath n'existe pas, utilisation du repertoire courant"
    $Logfile= "$scriptshortname-$stamp.log"
    }

# -----------------------------------------------------------------
# Fonction utilisables
# -----------------------------------------------------------------
Function Adapt-Screen ()
{
      $pshost = get-host
      $pswindow = $pshost.ui.rawui

      $newsize = $pswindow.buffersize
      $newsize.width = 200
      $pswindow.buffersize = $newsize
}

Function Create-ScheduledTask  ()
{
    $ComputerName = "localhost"
    $RunAsUser = "System"           
    $TaskRun = "'PowerShell.exe -NoLogo -File "  + $scriptLongName
    $Schedule = "DAILY /ST 00:00"
	                            
  
   if ($autorestart)  { $TaskRun = $TaskRun + " -autorestart" }
   if ($AutoRestartIfPending) {$TaskRun = $TaskRun + " -AutoRestartIfPending" }
   if ($nomail) {$TaskRun = $TaskRun + " -nomail" }
   if ($mailto) {$TaskRun = $TaskRun + " -mailto " + $mailto }
   $taskRun = $taskRun + "'"


   if ( [System.Environment]::OSVersion.Version.Major -gt 5) {
        write-verbose " systeme windows 2008 ou plus"
	$TaskName = "'claranet\Gestion des logs'"
        }
   else
       {
	write-verbose "system windows 2003"
	$TaskName = "'Gestion des logs'"
	}
                                                            
    $Command = "schtasks.exe /create /s $ComputerName /ru $RunAsUser /tn $TaskName /tr $TaskRun /sc $Schedule /F"            
    Invoke-Expression $Command            
}


function create_workfile ()
{
    if (Test-Path "$workPath\$consignePurge" ) 
        {
        write-host "le fichier existe déja. pas de creation"
        }
    else
       {
        '# Ce fichier contient les commandes appelées par le script de purge

# Par exemple:
#compress -rep "C:\claranet\log" -fichier "*.log" -retention 1 
#purge -rep "C:\claranet\log" -fichier "*.log" -retention 30 -r
#compacter -rep "C:\claranet\log" -fichier "*.log" -retention 2 -R
#compress -rep "C:\claranet\log"  -expression "LOG_purge-[0-9]{6}_15.+" -retention 10 -r
#compressDir -rep "C:\claranet\" -fichier "log*"  -retention 30
#ajouter vos propres purge de fichiers ici.
' | out-file "$workPath\$consignePurge"
 
    }
 }



Function Purge
{
	param(
		$rep,
		[Parameter(ParameterSetName="p1")] $fichier = "*",
		[Parameter(ParameterSetName="p2")]$expression= ".*",
		$retention = "30",
		[switch] $R = $false
		)

	

   if ((Test-Path $Rep) -ne $True) {write-host "le repertoire $rep n'existe pas!" -foregroundcolor red;return}
   if ($R) { 
      $liste=get-childitem $rep -Filter $Fichier -recurse
      ""
      "Purge du repertoire $rep : suppression des fichiers de plus de $retention jour(s) en mode reccursif"
      }
   else
      {
      $liste=get-childitem $rep -Filter $Fichier
      ""
      "Purge du repertoire $rep : suppression des fichiers de plus de $retention jour(s)"
      }

   switch ($PsCmdlet.ParameterSetName)
      {
      "p1" { "filtre utilise: $fichier `r`n"
           if ($liste) {$liste| ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |remove-item -verbose ; break}
           }
      "p2" {	"expression reguliere utilisee : $expression `r`n "
           if ($liste) { $liste| where-object {$_.name -match $expression} | ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |remove-item -verbose ; break}
           }
      }
}



function compacter
{
	param(
       $rep,
		[Parameter(ParameterSetName="p1")] $fichier = "*",
		[Parameter(ParameterSetName="p2")]$expression= ".*",
		$retention = "30",
		[switch] $R = $false
		)

	if ((Test-Path $Rep) -ne $True) {write-host "le repertoire $rep n'existe pas!" -foregroundcolor red;return}
        if ($R) { 
                $liste=get-childitem $rep -Filter $Fichier -recurse
                ""
                "Compression du repertoire $rep : compression NTFS des fichiers de plus de $retention jour(s) en mode reccursif"
        } else {
                $liste=get-childitem $rep -Filter $Fichier
                ""
                "Compression du repertoire $rep : compression NTFS des fichiers de plus de $retention jour(s)"
                }
                
	switch ($PsCmdlet.ParameterSetName)
                {
                "p1"  { "filtre utilise: $fichier `r`n"
                        if ($liste) {$liste| ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention)) -and !($_.attributes -match  "Compressed")} |ForEach-Object {compact /C $_.FullName /Q /F /I |out-default} ; break}
                      }
                "p2"  {	"expression reguliere utilisee : $expression `r`n "
                        if ($liste) {$liste| where-object {$_.name -match $expression} | ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention)) -and !($_.attributes -match  "Compressed") } |ForEach-Object {compact /C $_.FullName /Q /F /I |out-default}; break}
                      }
                }
}

function f7zip
{
    param(
        $file
        )

    $zipfilename = $file+".zip"

    write-host "Compressing [$file] to [$zipfilename]..." -nonewline

    sz u "$zipfilename" "$file"  -r |out-Null
    if ($lastexitcode -eq 0) {
        write-host "`t OK" -ForegroundColor green
        remove-item $file -Recurse -Force
    }else{
        write-host "`t KO" -ForegroundColor red
    }
}



function compress
{
	param(
		$rep,
		[Parameter(ParameterSetName="p1")] $fichier = "*",
		[Parameter(ParameterSetName="p2")]$expression= ".*",
		$retention = "30",
		[switch] $R = $false
		)

	if ((Test-Path $Rep) -ne $True) {write-host "le repertoire $rep n'existe pas!" -foregroundcolor red;return}
        
        if ($R) { 
                $liste=get-childitem $rep -Filter $Fichier -recurse
                ""
                "Compression du repertoire $rep : compression au format zip des fichiers de plus de $retention jour(s) en mode reccursif"
        } else {
                $liste=get-childitem $rep -Filter $Fichier
                ""
                "Compression du repertoire $rep : compression au format zip des fichiers de plus de $retention jour(s)"
                }
                
	switch ($PsCmdlet.ParameterSetName)
                {
                "p1"  { "filtre utilise: $fichier `r`n"
                        if ($liste) {$liste| ?{$_.name -notlike "*.zip"} | ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |ForEach-Object {$_ ; f7zip $_.FullName } ; break}
                      }
                "p2"  {	"expression reguliere utilisee : $expression `r`n "
                        if ($liste) {$liste| where-object {$_.name -match $expression} | ?{$_.name -notlike "*.zip"} | ?{!$_.PSIsContainer -and ($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |ForEach-Object {$_ ; f7zip $_.FullName }; break}
                      }
                }
}

function compressDir
{
	param(
		$rep,
		[Parameter(ParameterSetName="p1")] $fichier = "*",
		[Parameter(ParameterSetName="p2")]$expression= ".*",
		$retention = "30"
		)

	if ((Test-Path $Rep) -ne $True) {write-host "le repertoire $rep n'existe pas!" -foregroundcolor red;return}
        
        $liste=get-childitem $rep -Filter $Fichier | where-object {$_.PSIsContainer}
        ""
                       
	switch ($PsCmdlet.ParameterSetName)
                {
                "p1"  { "filtre utilise: $fichier `r`n"
                        if ($liste) {$liste| ?{$_.name -notlike "*.zip"} | ?{($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |ForEach-Object {f7zip $_.FullName } ; break}
                      }
                "p2"  {	"expression reguliere utilisee : $expression `r`n "
                        if ($liste) {$liste| where-object {$_.name -match $expression} | ?{$_.name -notlike "*.zip"} | ?{($_.LastWriteTime -lt (get-Date).adddays(-$retention))} |ForEach-Object {f7zip $_.FullName }; break}
                      }
                }
}



# -----------------------------------------------------------------
#	pre-traitements
# -----------------------------------------------------------------
if ($version) {
    Write-Host "`n$scriptName version $scriptversion"
    write-host "chemin d'installation: $scriptPath`n"
    exit 0
    }

if ($help) {
  Get-Help $scriptLongName -full
  exit 0
}
if ($createtask) {
    write-host "creation d'une tache planifiée..."
    Create-ScheduledTask
    }

if ($createconf) {
    write-host "creation du fichier local..."
    create_workfile
    exit
    }

if ($createtask -OR $createconf) {
    exit
    }

 if (-not (test-path "$7zPath")) {
    write-host "ERR - $7zPath est introuvable"
    exit 99
    }
    
 set-alias sz "$7ZPath"  
 
Adapt-Screen 
if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Start-transcript -path "$LogFile" -append |Out-Null}
# -----------------------------------------------------------------
#	debut des traitements
# -----------------------------------------------------------------
# purge des fichiers propres au serveur (dans d:\claranet\work\Log_purge.conf)
if ( test-path "$workPath\$consignePurge" ){
    foreach ($line in Get-Content($workPath+"\"+$consignePurge)) {
        if ($line.trim() -ne "") {
            Invoke-Expression $line
        }
    }
} else {
    Write-Verbose "le fichier n'existe pas. purge du repertoire de base uniquement"
}

purge -rep $LogPath -fichier "*.log" -retention 30 -r
compacter -rep $LogPath -fichier "*.log" -retention 2 -R



# -----------------------------------------------------------------
#	fin du script
# -----------------------------------------------------------------
write-host "--- fin du script ---"
if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Stop-Transcript | Out-Null}

 <#
 .SYNOPSIS
 Gere les fichiers de log

 .DESCRIPTION
 Ce script permet de purger les vieux fichiers de log et/ou de compresser les fichiers ou repertoires
 pour ajouter des fichiers, il suffit de créer un fichier de conf et d'ajouter des ligne de commandes dedans.
 l'option createconf créé un fichier de conf au bon endroit avec un exemple de commande en commentaire
 les fonctions utlisable dans le fichier de conf sont:
 compress    = > pour compresser les fichiers de log en fichier zip
 compressDir = > pour compresser des repertoires (un fichier zip par repertoire)
 purge      => pour supprimer les vieux fichiers
 compacter  => pour compacter les fichier NTFS
 
 Les parametres de chaque fonction sont les mêmes:
 -rep  <repertoire> => le nom du repertoire qui contient les fichiers
 -fichier <extention> | -expression <regexp> => le type de fichier à traiter ou une expression reguliere  
 -retention <retention>  => la durée de retention souhaitée
  
 -r  => pour parcourir les sous-répertoire de <repertoire> de façon reccursive (ne fonctionne pas avec compressDir)

 .PARAMETER version
  affiche le numéro de version et se termine sans rien faire

 .PARAMETER help
 Affiche cette aide

 .PARAMETER createconf
  créer un fichier de conf dans le repertoire work

 .PARAMETER createtask
  créer une tache planifiée dans le dossier Claranet.
  par defaut, elle tourne à minuit tous les jours
   
#>
