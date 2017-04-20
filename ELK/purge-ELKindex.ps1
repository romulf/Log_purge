# -----------------------------------------------------------------
#
# 	Objet		ELK
# 	Fonction	purge les index
#
# 	Auteur 		Erwan Le Tiec
# 	Creation	05/04/2017
#   modif 
#
# -----------------------------------------------------------------
# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------
param (
  [int] $retention=17,
  [switch] $version= $false,
  [switch] $help=$false
)

$scriptversion = 1.0 
$minDate=(get-date).adddays(-$retention)

# -----------------------------------------------------------------
# Fonction utilisables
# -----------------------------------------------------------------

function delete-index ($id ) {
  write-host "suppression de l'index $id ...`t" -nonewline
  $Uri="http://localhost:9200/" + $id
   (Invoke-WebRequest  -Uri $Uri -method DELETE).Statusdescription
 }

# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$scriptName = $myInvocation.MyCommand.Name
$scriptLongName = $myInvocation.MyCommand.path
$scriptshortname= $scriptName.substring(0,$scriptName.length-4)

$Stamp=(get-date).ToString("yyMMdd_HHmm")
$ScriptPath=split-path $MyInvocation.MyCommand.Path -parent

$Logfile= "D:\logs\claranet\$scriptshortname-$stamp.log"

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

if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Start-transcript -path "$LogFile" -append |Out-Null}
# -----------------------------------------------------------------
#	debut des traitements
# -----------------------------------------------------------------
$listeIndex=@()
$request=(Invoke-WebRequest -method GET http://localhost:9200/_cat/indices?v).content

 $request -split ("`n") | foreach-object{
  $ligne=(($_ -replace '[ ]+'," ") -split (" "))
  if ($ligne[2] -like "winlogbeat-*") {
    $Elkindex=@{}
    $Elkindex.add("health",$ligne[0])
    $Elkindex.add("status",$ligne[1])
    $Elkindex.add("index",$ligne[2])
    $Elkindex.add("uuid",$ligne[3])
    $Elkindex.add("pri",$ligne[4])
    $Elkindex.add("rep",$ligne[5])
    $Elkindex.add("docs_count",$ligne[6])
    $Elkindex.add("docs_deleted",$ligne[7])
    $Elkindex.add("store_size",$ligne[8])
    $Elkindex.add("pri_store_size",$ligne[9])
    $Elkindex.add("date",(get-date ($ligne[2] -replace "winlogbeat-","")) )
    $listeIndex+= new-Object PSObject -property $Elkindex
	}
}

foreach ( $i in $listeIndex) {
  if ($i.date -lt $minDate) {delete-index $i.index}
}
# -----------------------------------------------------------------
#	fin du script
# -----------------------------------------------------------------
if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Stop-Transcript | Out-Null}

 <#
 .SYNOPSIS
 Purge des anciens index Elk

 .DESCRIPTION
  Ce script purge les index winlogbeat de la stack elk en local

 .PARAMETER retention
 durée de rétention souhaitée (en jour)
 
 .PARAMETER version
  affiche le numéro de version et se termine sans rien faire

 .PARAMETER help
 Affiche cette aide
   
#>
