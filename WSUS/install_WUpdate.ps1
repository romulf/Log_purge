# -----------------------------------------------------------------
#
#       Objet           WSUS
#       Fonction        mise à jour du client
#
#       Auteur          Erwan Le Tiec
#       Creation        05/07/2013
#
#
# -----------------------------------------------------------------

# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------
param (
        [switch] $autorestart = $false,
        [switch] $AutoRestartIfPending = $False,
        [switch] $nomail = $false,
        [string[]] $mailto = @(),
    [switch] $createtask= $false,
    [switch] $version= $false
)

$scriptversion = 1.9

$PSEmailServer="nom_du_serveur_mail"
$Criteria="IsInstalled=0 and Type='Software'"
$resultcode= @{0="Not Started"; 1="In Progress"; 2="Succeeded"; 3="Succeeded With Errors"; 4="Failed" ; 5="Aborted" }
[string[]]$destinataire ="addresse_mail_par_defaut"
$destinataire = $destinataire + $mailto

$expediteur="addresse_mail_expediteur"
$sujet="$env:computername - Rapport de mise a jour"
$message="Bonjour,<br/>Vous trouverez ci-joint le rapport de mise a jour du serveur $env:computername<br/>Cordialement,<br/>Claranet<br/>"


# -----------------------------------------------------------------
# Fonction utilisables
# -----------------------------------------------------------------
Function Adapt-Screen () {
      $pshost = get-host
      $pswindow = $pshost.ui.rawui

      $newsize = $pswindow.buffersize
      $newsize.width = 200
      $pswindow.buffersize = $newsize
}

Function Create-ScheduledTask   {
    $ComputerName = "localhost"
    $RunAsUser = "System"
    $TaskRun = "'PowerShell.exe -NoLogo -File "  + $scriptLongName
    $Schedule = "MONTHLY /M * /MO FOURTH /D TUE /ST 07:00"


   if ($autorestart)  { $TaskRun = $TaskRun + " -autorestart" }
   if ($AutoRestartIfPending) {$TaskRun = $TaskRun + " -AutoRestartIfPending" }
   if ($nomail) {$TaskRun = $TaskRun + " -nomail" }
   if ($mailto) {$TaskRun = $TaskRun + " -mailto " + $mailto }
   $taskRun = $taskRun + "'"


   if ( [System.Environment]::OSVersion.Version.Major -gt 5) {
            write-verbose " systeme windows 2008 ou plus"
            $TaskName = "'claranet\Update windows via WSUS'"
    }else {
            write-verbose "system windows 2003"
            $TaskName = "'Update windows via WSUS'"
        }

  $Command = "schtasks.exe /create /s $ComputerName /ru $RunAsUser /tn $TaskName /tr $TaskRun /sc $Schedule /F"
  Invoke-Expression $Command
 }

Function Get-WindowsUpdateConfig

{$AUSettings = (New-Object -com "Microsoft.Update.AutoUpdate").Settings

 $AUObj = New-Object -TypeName System.Object

 Add-Member -inputObject $AuObj -MemberType NoteProperty -Name "NotificationLevel" -Value $AutoUpdateNotificationLevels[$AUSettings.NotificationLevel]
 Add-Member -inputObject $AuObj -MemberType NoteProperty -Name "UpdateDays" -Value $AutoUpdateDays[$AUSettings.ScheduledInstallationDay]
 Add-Member -inputObject $AuObj -MemberType NoteProperty -Name "UpdateHour" -Value $AUSettings.ScheduledInstallationTime 
 Add-Member -inputObject $AuObj -MemberType NoteProperty -Name "Recommended updates" -Value $(IF ($AUSettings.IncludeRecommendedUpdates) {"Included."}  else {"Excluded."})
 $AuObj

} 

# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------

# mettre Continue pour afficher le test, SilentlyContinue sinon
$VerbosePreference = "Continue"


$scriptName = $myInvocation.MyCommand.Name
$scriptLongName = $myInvocation.MyCommand.path
$scriptshortname= $scriptName.substring(0,$scriptName.length-4)

$Stamp=(get-date).ToString("yyMMdd_HHmm")
$ScriptPath=split-path $MyInvocation.MyCommand.Path -parent

$Logfile= "$ScriptPath\$scriptshortname-$stamp.log"



# -----------------------------------------------------------------
#       pre-traitements
# -----------------------------------------------------------------
if ($version) {
    Write-Host "$scriptName version $scriptversion"
    write-host "chemin d'installation: $scriptPath"
    exit 0
    }
Adapt-Screen    
if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Start-transcript -path "$LogFile" -append}


write-verbose "autorestart vaut: $AutoRestart"
Write-Verbose "AutoRestartIfPending vaut: $AutoRestartIfPending"
Write-Verbose "liste des destinataires du mail: $destinataire"


if ($createtask ) {
    Write-Host "Creation d'une tache planifiée..."
    Create-ScheduledTask
    exit
    }

#Testing if there are any pending reboots from earlier Windows Update sessions
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){
    Write-host  "Il y a un reboot en attente suite à une mise à jour precedente."

    #Reboot if autorestart for pending updates is enabled
    if ($AutoRestartIfPending) {
        write-host "Reboot automatique avant les mises à jour..."
     if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Stop-Transcript}
        Invoke-Command { shutdown /r /f /d p:1:1 /t 30 /c "Mise à jour Windows" }
        }

if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Stop-Transcript}
    exit
}


# -----------------------------------------------------------------
#       traitements
# -----------------------------------------------------------------

#Checking for available updates
$updateSession = new-object -com "Microsoft.Update.Session"
write-host "Checking available updates..."
$updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates
$downloader = $updateSession.CreateUpdateDownloader()
$downloader.Updates = $Updates

#If no updates available, do nothing
if ($downloader.Updates.Count -eq "0") {
        Write-Host "Aucune mise à jour trouvée"
    if (! $nomail) {
            Write-Verbose "envoie du mail..."
            $message+=",<br />Aucune mise a jour trouvee,<br />"
            Send-MailMessage -to $destinataire -from "$expediteur" -subject "$sujet" -body "$message" -BodyAsHtml
            }
   }
else
    {
    #If updates are available, download and install
    write-host "Téléchargement de $($downloader.Updates.count) updates..."


    $Result= $downloader.Download()

    if (($Result.Hresult -eq 0) -and (($result.resultCode -eq 2) -or ($result.resultCode -eq 3)) ) {
       $updatesToInstall = New-object -com "Microsoft.Update.UpdateColl"
       $Updates | where {$_.isdownloaded} | foreach-Object {$updatesToInstall.Add($_) | out-null }

                $installer = $updateSession.CreateUpdateInstaller()
                $installer.Updates = $updatesToInstall

                write-host "Installation de $($Installer.Updates.count) updates..."

                $installationResult = $installer.Install()
                $Global:counter=0

                $Report = $installer.updates |
                                Select-Object -property Title,EulaAccepted,
                @{Name='Result';expression= {$ResultCode[$installationResult.GetUpdateResult($Global:Counter).resultCode ] }   },
                @{Name='Reboot required';expression={$installationResult.GetUpdateResult($Global:Counter).RebootRequired;$Global:Counter++ }}
        $report | ft -AutoSize| Out-String -stream -width 250

        if (! $nomail) {
            Write-Verbose "envoie du mail..."
            $message=$report|ConvertTo-Html
            Send-MailMessage -to $destinataire -from "$expediteur" -subject "$sujet" -body "$message" -BodyAsHtml
            }

                #Reboot if autorestart is enabled and one or more updates are requiring a reboot
                if ($autoRestart -and $installationResult.rebootRequired) {
            write-host "Reboot automatique demandé..."
            Invoke-Command { shutdown /r /f /d p:1:1 /t 30 /c "Mise à jour Windows" }
         }
        }
}



# -----------------------------------------------------------------
#       fin du script
# -----------------------------------------------------------------
if (!(test-path ENV:\SCHEDULER_JOB_NAME)) {Stop-Transcript}

 <#
 .SYNOPSIS
 Installe les updates mises a dispo par le serveur WSUS

 .DESCRIPTION
 Installe les mise à jour disponibles
 il est possible de rebooter ensuite si une mise à jour le demande grace à la variable AutoRestart
 il est possible de rebooter avant si necessaire grace à la variable AutoRestartIfPending

 .PARAMETER version
 affiche le numéro de version et se termine sans rien faire

.PARAMETER autorestart
 indique qu'il faut rebooter le serveur  à la fin des update si c'est necessaire

.PARAMETER AutoRestartIfPending
 Indique qu'il faut rebooter le serveur si il y a un reboot en attente.
La machine rebootera mais ne fera pas de mise a jour ensuite. il faudra relancer le script

.PARAMETER nomail
 Ne pas envoyer de mail à la fin de la procedure

.PARAMETER destinataire
 liste des destinataires supplementaire pour le mail de fin de traitement

.PARAMETER createtask
 creer une tache planifée dans le dossier Claranet pour executer le script tous les 4eme mardi du mois à 07:00
 avec le compte system. il est possible ensuite de modifier la tache planifée (a la main)

#>
