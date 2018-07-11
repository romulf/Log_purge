<#
.NOTES
Objet:           WSUS
Auteur:          Erwan Le Tiec
Creation:        10/07/2018
Modif:          

.SYNOPSIS
   Affiche des informations de configuration WSUS

.DESCRIPTION
   lit la base de registre et affiche les infos de configuration WSUS suivante:

- WUserver             : Url du serveur WSUS

- WUStatusServer       : Url du serveur qui recoit les rapports. Elle doit etre la même que WUserver

- Group                : Groupe d'appartenance de la machine dans la console WSUS

- WindowsUpdateAccess  : Indique si windows update est activé ou pas. si l'acces est désactivé. il n'y a pas de mise a jour automatique non plus

- AUOption             : Actions effectuées lors des update automatiques. il y a plusieurs types d'options possibles:
                1 - La fonction Maintenir mon ordinateur à jour a été désactivée dans le service Mises à jour automatiques.
                2 - Notification des téléchargements et des installations
                3 - téléchargement automatique et notification des installations (parametre par défaut)
                4 - téléchargement automatique et planification des installations. Si il n'y a pas de planification, les mises à jour sont installées par défaut tous les jours à 3h00.

- ScheduledInstallTime : Si les installations sont planifiées, heure a laquelle les mises a jour sont installées

- ScheduledInstallDay  : Si les installations sont planifiées, jour d'installation: 0 pour Chaque jour. 1 à 7 pour les jours de la semaine du dimanche (1) au samedi (7).

- UpdateSources        : Indique si le service Mises à jour automatiques utilise un serveur WSUS ou les serveurs Miwcorsoft Windows Update. 

- AutoUpdate           : Indique si le service Mises à jour automatiques est activé.

- WUServerCheck        : Test de l'acces au serveur WSUS
- DetectLastSuccessTime : Date de la derniere synchro reussie
- DownloadLastSuccessTime: Date du dernier téléchargement de Mise à jour réussi

.PARAMETER version
  affiche le numéro de version et se termine sans rien faire

.EXAMPLE

PS C:\> get-WUConfig.ps1
WUserver              : http://wsus.corporate.local:8530
WUStatusServer        : http://wsus.corporate.local:8530
Group                 : claranet
WindowsUpdateAccess   : Autorization not Not Defined
AUOption              : 2 = Notify before download.
UpdateSources         : WSUS server
AutoUpdate            : Activated
WUServerCheck         : acces OK
DetectLastSuccessTime : 11/07/2018 09:28:57
DownloadSuccessTime   : 21/06/2018 07:06:46
#>

# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------
[cmdletBinding()]
param (
  [switch] $version= $false
)

$scriptversion = "1.0"


# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$scriptName = $myInvocation.MyCommand.Name       # get-WUHistory.ps1
$scriptshortname= $scriptName.substring(0,$scriptName.length-4) # get-WUHistory
$scriptLongName = $myInvocation.MyCommand.path   # D:\sources\claranet\get-WUHistory.ps1

$Stamp=(get-date).ToString("yyMMdd_HHmm")

$LogPath="D:\logs\claranet"

if ((Test-Path $LogPath) -eq $True) {
    $Logfile= "$LogPath\$scriptshortname-$stamp.log"
}else{
    write-verbose "le repertoire $LogPath n'existe pas, utilisation du repertoire courant"
    $Logfile= "$scriptshortname-$stamp.log"
}


# -----------------------------------------------------------------
# Fonction utilisables
# -----------------------------------------------------------------
# -----------------------------------------------------------------
#	pre-traitements
# ----------------------------------------------------------------- 
if ($version) {
    Write-Host "`nnom                   : $scriptName"
    write-host "version               : $scriptversion"
    write-host "chemin d'installation : $scriptLongName `n"
    exit 0
    }


start-Transcript -Path $Logfile | Out-Null
# -----------------------------------------------------------------
#       traitements
# -----------------------------------------------------------------

$WUReg=get-itemProperty 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' | Select-Object WUServer, WUStatusServer, TargetGroup, TargetGroupEnabled, DisableWindowsUpdateAccess
$WUAUReg=Get-ItemProperty 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' |Select-Object AUOptions,UseWUServer,NoAutoUpdate,ScheduledInstallTime,ScheduledInstallDay

$WUConf=New-Object psobject
$WUConf | Add-Member NoteProperty -name "WUserver" -Value $WUReg.WUServer
$WUConf | Add-Member NoteProperty -name "WUStatusServer" -Value $WUReg.WUStatusServer


if (!($WUReg.TargetGroup)) {$WUReg.TargetGroup = "Not Defined"}
if ($WUReg.TargetGroupEnabled) {
    $WUConf | Add-Member NoteProperty -name "Group" -Value $WUReg.TargetGroup
}else{
    $WUConf | Add-Member NoteProperty -name "Group" -Value "$($WUReg.TargetGroup) (disabled!)"
}

#DisableWindowsUpdateAccess
if (!($WUReg.DisableWindowsUpdateAccess -eq "")) {
    $WUConf | Add-Member NoteProperty -name 'WindowsUpdateAccess' -Value 'Autorization not Not Defined'
}elseif ($WUReg.DisableWindowsUpdateAccess -eq '0'){
    $WUConf | Add-Member NoteProperty -name 'WindowsUpdateAccess' -Value 'Authorized'
}else{
    $WUConf | Add-Member NoteProperty -name 'WindowsUpdateAccess' -Value 'forbidden'
}


#AUOptions
switch ($WUAUReg.AUOptions) {
    1 { $WUConf | Add-Member NoteProperty -name "AUOption" -Value "1 = Keep my computer up to date has been disabled in Automatic Updates." }
    2 { $WUConf | Add-Member NoteProperty -name "AUOption" -Value "2 = Notify before download." }
    3 { $WUConf | Add-Member NoteProperty -name "AUOption" -Value "3 = Automatically download and notify of installation."}
    4 { $WUConf | Add-Member NoteProperty -name "AUOption" -Value "4 = Automatically download and schedule installation."
        $WUConf | Add-Member NoteProperty -name "ScheduledInstallTime" -Value $WUAUReg.ScheduledInstallTime
        $WUConf | Add-Member NoteProperty -name "ScheduledInstallDay" -Value $WUAUReg.ScheduledInstallDay
      }
    5 { $WUConf | Add-Member NoteProperty -name "AUOption" -Value "5 = Automatic Updates is required and users can configure it."}
   default {$WUConf | Add-Member NoteProperty -name "AUOption" -Value $WUReg.AUOptions}
}

#UseWUServer
switch ($WUAUReg.UseWUServer) {
    0 {$WUConf | Add-Member NoteProperty -name "UpdateSources" -Value 'Microsoft Update'}
    1 {$WUConf | Add-Member NoteProperty -name "UpdateSources" -Value 'WSUS server'}
    default {$WUConf | Add-Member NoteProperty -name "UpdateSources" -Value "Not Defined"}
}


#NoAutoUpdate
switch ($WUAUReg.NoAutoUpdate) {
    0 {$WUConf | Add-Member NoteProperty -name 'AutoUpdate' -Value 'Activated'}
    1 {$WUConf | Add-Member NoteProperty -name 'AutoUpdate' -Value 'Disabled'}
    default {$WUConf | Add-Member NoteProperty -name 'AutoUpdate' -Value 'Not Defined'}
}

if ($(Invoke-WebRequest $WUConf.WUserver).StatusCode -eq 200) {
    $WUConf | Add-Member NoteProperty -name 'WUServerCheck' -Value 'acces OK'
}else{
    $WUConf | Add-Member NoteProperty -name 'WUServerCheck' -Value 'ERR - no access!'
}

<# a checker: Cette partie ne fonctionne plus depuis debut 2018
#- DetectLastSuccessTime : Date de la derniere synchro reussie
#- DownloadLastSuccessTime: Date du dernier téléchargement de Mise à jour réussi
#- InstallLastSuccessTime : Date de la dernière installation réussie

$WUresult="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results"
if (test-path $WUresult) {
    $r=Get-ItemProperty "${WUResult}\Detect\"
    $WUConf | Add-Member NoteProperty -name 'DetectLastSuccessTime' -Value $(get-date $r.lastSuccessTime)

    $r=Get-ItemProperty "${WUResult}\Download\"
    $WUConf | Add-Member NoteProperty -name 'DownloadLastSuccessTime' -Value $(get-date $r.lastSuccessTime)

    $r=Get-ItemProperty "${WUResult}\Install\"
    $WUConf | Add-Member NoteProperty -name 'InstallLastSuccessTime' -Value $(get-date $r.lastSuccessTime)
}
#>

$r=get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WindowsUpdateClient/Operational';id=26,40} -MaxEvents 1 -ErrorAction SilentlyContinue
if ($r) {
    $WUConf | Add-Member NoteProperty -name 'DetectLastSuccessTime' -Value $r.timeCreated
}
$r=get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WindowsUpdateClient/Operational';id=41}  -MaxEvents 1 -ErrorAction SilentlyContinue
if ($r) {
    $WUConf | Add-Member NoteProperty -name 'DownloadSuccessTime' -Value $r.timeCreated
}


$WUConf

# -----------------------------------------------------------------
#       fin du script
# -----------------------------------------------------------------
Stop-Transcript | Out-Null
