
<#
.SYNOPSIS
    Download and install updates.

.NOTES
Auteur:          Erwan Le Tiec (from ? )
Creation:        18/05/2017
Modif:

.DESCRIPTION
        Use Install-WUUpdate.ps1 to get list of available updates, next download and install it. 
        There are two types of filtering update: Pre search criteria, Post search criteria.
        - Pre search works on server side, like example: ( IsInstalled = 0 and IsHidden = 0 and CategoryIds contains '0fa1201d-4330-4fa8-8ae9-b877473b6441' )
        - Post search work on client side after downloading the pre-filtered list of updates, like example $KBArticleID -match $Update.KBArticleIDs
        
        Update occurs in four stages: 
            1. Search for updates,
            2. Choose updates,
            3. Download updates,
            4. Install updates.
        
    .PARAMETER UpdateType
        Pre search criteria. Finds updates of a specific type, such as 'Driver' and 'Software'. Default value contains all updates.

    .PARAMETER UpdateID
        Pre search criteria. Finds updates of a specific UUID (or sets of UUIDs), such as '12345678-9abc-def0-1234-56789abcdef0'.

    .PARAMETER RevisionNumber
        Pre search criteria. Finds updates of a specific RevisionNumber, such as '100'. This criterion must be combined with the UpdateID param.

    .PARAMETER CategoryIDs
        Pre search criteria. Finds updates that belong to a specified category (or sets of UUIDs), such as '0fa1201d-4330-4fa8-8ae9-b877473b6441'.

    .PARAMETER IsInstalled
        Pre search criteria. Finds updates that are installed on the destination computer.

    .PARAMETER IsHidden
        Pre search criteria. Finds updates that are marked as hidden on the destination computer. Default search criteria is only not hidden updates.

    .PARAMETER WithHidden
        Pre search criteria. Finds updates that are both hidden and not on the destination computer. Overwrite IsHidden param. Default search criteria is only not hidden upadates.
        
    .PARAMETER Criteria
        Pre search criteria. Set own string that specifies the search criteria.

    .PARAMETER ShowSearchCriteria
        Show choosen search criteria. Only works for pre search criteria.
        
    .PARAMETER Category
        Post search criteria. Finds updates that contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...
        
    .PARAMETER KBArticleID
        Post search criteria. Finds updates that contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.
    
    .PARAMETER Title
        Post search criteria. Finds updates that match part of title, such as ''

    .PARAMETER NotCategory
        Post search criteria. Finds updates that not contain a specified category name (or sets of categories name), such as 'Updates', 'Security Updates', 'Critical Updates', etc...
        
    .PARAMETER NotKBArticleID
        Post search criteria. Finds updates that not contain a KBArticleID (or sets of KBArticleIDs), such as 'KB982861'.
    
    .PARAMETER NotTitle
        Post search criteria. Finds updates that not match part of title.
        
    .PARAMETER IgnoreUserInput
        Post search criteria. Finds updates that the installation or uninstallation of an update can't prompt for user input.
    
    .PARAMETER IgnoreRebootRequired
        Post search criteria. Finds updates that specifies the restart behavior that not occurs when you install or uninstall the update.
    
    .PARAMETER ListOnly
        Show list of updates only without downloading and installing. Works similar like Get-WUList.
    
    .PARAMETER DownloadOnly
        Show list and download approved updates but do not install it. 
    
    .PARAMETER AcceptAll
        Do not ask for confirmation updates. Install all available updates.
    
    .PARAMETER AutoReboot
        Do not ask for rebbot if it needed.
    
    .PARAMETER IgnoreReboot
        Do not ask for reboot if it needed, but do not reboot automaticaly. 
    
    .PARAMETER AutoSelectOnly  
        Install only the updates that have status AutoSelectOnWebsites on true.

    .PARAMETER Debuger	
        Debug mode.

    .EXAMPLE
        Get updates from specific source with title contains ".NET Framework 4". Everything automatic accept and install.
    
        PS C:\> Install-WUUpdate.ps1 -Title ".NET Framework 4" -AcceptAll

        X Status     KB          Size Title
        - ------     --          ---- -----
        2 Accepted   KB982670   48 MB Program Microsoft .NET Framework 4 Client Profile w systemie Windows 7 dla systemów op...
        3 Downloaded KB982670   48 MB Program Microsoft .NET Framework 4 Client Profile w systemie Windows 7 dla systemów op...
        4 Installed  KB982670   48 MB Program Microsoft .NET Framework 4 Client Profile w systemie Windows 7 dla systemów op...

    .EXAMPLE
        Get updates with specyfic KBArticleID. Check if type are "Software" and automatic install all.
        
        PS C:\> $KBList = "KB890830","KB2533552","KB2539636"
        PS C:\> Install-WUUpdate.ps1 -UpdateType "Software" -KBArticleID $KBList -AcceptAll

        X Status     KB          Size Title
        - ------     --          ---- -----
        2 Accepted   KB2533552   9 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2533552)
        2 Accepted   KB2539636   4 MB Aktualizacja zabezpieczeń dla programu Microsoft .NET Framework 4 w systemach Windows ...
        2 Accepted   KB890830    1 MB Narzędzie Windows do usuwania złośliwego oprogramowania dla komputerów z procesorem x6...
        3 Downloaded KB2533552   9 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2533552)
        3 Downloaded KB2539636   4 MB Aktualizacja zabezpieczeń dla programu Microsoft .NET Framework 4 w systemach Windows ...
        3 Downloaded KB890830    1 MB Narzędzie Windows do usuwania złośliwego oprogramowania dla komputerów z procesorem x6...	
        4 Installed  KB2533552   9 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2533552)
        4 Installed  KB2539636   4 MB Aktualizacja zabezpieczeń dla programu Microsoft .NET Framework 4 w systemach Windows ...
        4 Installed  KB890830    1 MB Narzędzie Windows do usuwania złośliwego oprogramowania dla komputerów z procesorem x6...
    
    .EXAMPLE
        Get list of updates without language packs and updatets that's not hidden.

        PS C:\> Install-WUUpdate.ps1 -NotCategory "Language packs" -ListOnly

        X Status KB          Size Title
        - ------ --          ---- -----
        1 ------ KB2640148   8 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2640148)
        1 ------ KB2600217  32 MB Aktualizacja dla programu Microsoft .NET Framework 4 w systemach Windows XP, Se...
        1 ------ KB2679255   6 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2679255)
        1 ------ KB915597    3 MB Definition Update for Windows Defender - KB915597 (Definition 1.125.146.0)
        
    .LINK
        http://gallery.technet.microsoft.com/scriptcenter/2d191bcd-3308-4edd-9de2-88dff796b0bc
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=vs.85).aspx
        http://msdn.microsoft.com/en-us/library/windows/desktop/aa386099(v=vs.85).aspx
        http://msdn.microsoft.com/en-us/library/ff357803(VS.85).aspx

    .LINK
        Get-WUHistory
        Get-WUList
    #>

# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------

[CmdletBinding(	)]
Param(
    #Pre search criteria
    [parameter(ValueFromPipelineByPropertyName=$true)]	[ValidateSet("Driver", "Software")]	[String]$UpdateType="",
    [parameter(ValueFromPipelineByPropertyName=$true)]	[Switch]$IsHidden,
    [parameter(ValueFromPipelineByPropertyName=$true)]	[Switch]$WithHidden,

    #Post search criteria
    [parameter(ValueFromPipelineByPropertyName=$true)]	[String[]]$Category="",
    [parameter(ValueFromPipelineByPropertyName=$true)]	[String[]]$KBArticleID,
    [parameter(ValueFromPipelineByPropertyName=$true)]	[String]$Title,

    [parameter(ValueFromPipelineByPropertyName=$true)]	[String[]]$NotCategory="",
    [parameter(ValueFromPipelineByPropertyName=$true)]	[String[]]$NotKBArticleID,
    [parameter(ValueFromPipelineByPropertyName=$true)]	[String]$NotTitle,
    
    [parameter(ValueFromPipelineByPropertyName=$true)]	[Alias("Silent")]	[Switch]$IgnoreUserInput,
    [parameter(ValueFromPipelineByPropertyName=$true)]	[Switch]$IgnoreRebootRequired,

    #Mode options
    [Switch]$ListOnly,
    [Switch]$DownloadOnly,
    [Alias("All")] [Switch]$AcceptAll,
    [Switch]$AutoReboot,
    [Switch]$IgnoreReboot,
    [Switch]$AutoSelectOnly,

    [switch]$version,
    [switch]$help
)

$UpdateCollection = @()
$defaultProperties = @("KB","status","size","Title")
$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

$Scriptversion=1.0


# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$logPath="D:\logs\claranet"
$scriptName = $myInvocation.MyCommand.Name       # get-WUHistory.ps1
$scriptLongName = $myInvocation.MyCommand.path   # D:\sources\claranet\get-WUHistory.ps1
$scriptshortName= $scriptName.substring(0,$scriptName.length-4)  # get-WUHistory

$Stamp=(get-date).ToString("yyMMdd_HHmm")
$Logfile= "$LogPath\$scriptshortname-$stamp.log"

# -----------------------------------------------------------------
# Fonctions utilisables
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

if ($help) {
    Get-Help $scriptLongName -full 
    exit 0
}

$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if(!$Role) {
    Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."	
}

Start-transcript -path "$LogFile" -append |Out-Null
# -----------------------------------------------------------------
#       traitements
# -----------------------------------------------------------------

# --------------------- Stage 0 --------------------------------------------
Write-verbose "STAGE 0: Prepare environment"


#Check reboot status and reboot if autoreboot
$objSystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"
If($objSystemInfo.RebootRequired) {
    Write-Warning "Reboot is required to continue"
    If($AutoReboot) {
        Restart-Computer -Force
    }
    Return
}

Write-Debug "Set number of stage"
If($ListOnly) {
    $NumberOfStage = 2
}ElseIf($DownloadOnly){
    $NumberOfStage = 3
} Else{
    $NumberOfStage = 4
}

# -------------------- STAGE 1 ---------------------------------------------
Write-verbose "STAGE 1: Get updates list"
Write-Debug "Create Microsoft.Update.ServiceManager object"
$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 

Write-Debug "Create Microsoft.Update.Session object"
$objSession = New-Object -ComObject "Microsoft.Update.Session"

Write-Debug "Create Microsoft.Update.Session.Searcher object"
$objSearcher = $objSession.CreateUpdateSearcher()


Foreach ($objService in $objServiceManager.Services) {
    If($objService.IsDefaultAUService -eq $True){
        $serviceName = $objService.Name
        Break
    }
}

Write-host "Connecting to $serviceName server. Please wait..."
Try{
    $search = "IsInstalled = 0"	
    Write-verbose "Set pre search criteria: IsInstalled = 0"
        
    If($UpdateType -ne ""){
        Write-Verbose "Set pre search criteria: Type = $UpdateType"
        $search += " and Type = '$UpdateType'"
    }

    If($IsHidden) {
        Write-Verbose "Set pre search criteria: IsHidden = 1"
        $search += " and IsHidden = 1"
    }ElseIf($WithHidden){
        Write-Verbose "Set pre search criteria: IsHidden = 1 and IsHidden = 0"
    }Else {
        Write-Verbose "Set pre search criteria: IsHidden = 0"
        $search += " and IsHidden = 0"
    }

    If($IgnoreRebootRequired) {
        Write-Verbose "Set pre search criteria: RebootRequired = 0"
        $search += " and RebootRequired = 0"	
    }

    Write-Verbose "Search criteria is: $search"
    $objResults = $objSearcher.Search($search)
}

Catch	{
    If($_ -match "HRESULT: 0x80072EE2")	{ Write-Warning "Probably you don't have connection to Windows Update server"}
    Return
}

$objCollectionUpdate = New-Object -ComObject "Microsoft.Update.UpdateColl" 

$NumberOfUpdate = 1
$UpdateCollection = @()
$UpdatesExtraDataCollection = @{}
$PreFoundUpdatesToDownload = $objResults.Updates.count
Write-Verbose "Found [$PreFoundUpdatesToDownload] Updates in pre search criteria"

Foreach($Update in $objResults.Updates){
    $UpdateAccess = $true
    Write-Progress -Activity "Post search updates for $Computer" -Status "[$NumberOfUpdate/$PreFoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$PreFoundUpdatesToDownload * 100))
    Write-Verbose "Set post search criteria: $($Update.Title)"

    If($Category -ne "") {
        $UpdateCategories = $Update.Categories | Select-Object Name
        Write-Debug "Set post search criteria: Categories = '$([string]::join(", ", $Category))'"	
        Foreach($Cat in $Category){
            If(!($UpdateCategories -match $Cat)) {
                Write-Debug "UpdateAccess: false"
                $UpdateAccess = $false
            } Else{
                $UpdateAccess = $true
                Break
            }
        }
    }

    If($NotCategory -ne "" -and $UpdateAccess -eq $true)	{
        $UpdateCategories = $Update.Categories | Select-Object Name
        Write-Debug "Set post search criteria: NotCategories = '$([string]::join(", ", $NotCategory))'"	
        Foreach($Cat in $NotCategory) {
            If($UpdateCategories -match $Cat) {
                Write-Debug "UpdateAccess: false"
                $UpdateAccess = $false
                Break
            }
        }
    }

    If($KBArticleID -ne $null -and $UpdateAccess -eq $true) {
        Write-Debug "Set post search criteria: KBArticleIDs = '$([string]::join(", ", $KBArticleID))'"
        If(!($KBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs)) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($NotKBArticleID -ne $null -and $UpdateAccess -eq $true)
    {
        Write-Debug "Set post search criteria: NotKBArticleIDs = '$([string]::join(", ", $NotKBArticleID))'"
        If($NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($Title -and $UpdateAccess -eq $true) {
        Write-Debug "Set post search criteria: Title = '$Title'"
        If($Update.Title -notmatch $Title) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($NotTitle -and $UpdateAccess -eq $true) {
        Write-Debug "Set post search criteria: NotTitle = '$NotTitle'"
        If($Update.Title -match $NotTitle) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($IgnoreUserInput -and $UpdateAccess -eq $true) {
        Write-Debug "Set post search criteria: CanRequestUserInput"
        If($Update.InstallationBehavior.CanRequestUserInput -eq $true) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($IgnoreRebootRequired -and $UpdateAccess -eq $true) {
        Write-Debug "Set post search criteria: RebootBehavior"
        If($Update.InstallationBehavior.RebootBehavior -ne 0) {
            Write-Debug "UpdateAccess: false"
            $UpdateAccess = $false
        }
    }

    If($UpdateAccess -eq $true) {
        Switch($Update.MaxDownloadSize)	{
            {[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
            {[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }
            {[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }
            {[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
            default { $size = $_+"B" }
        }
    
        If($Update.KBArticleIDs -ne "") {
            $KB = "KB"+$Update.KBArticleIDs
        } Else{
            $KB = ""
        }

        If($ListOnly) {
            $Status = ""
            If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
            If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
            If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
            If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
            If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
            If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 

            Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME
            Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
            Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
            Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
            Add-Member -InputObject $Update -MemberType NoteProperty -Name X -Value 1
            Add-Member -InputObject $Update MemberSet PSStandardMembers $PSStandardMembers	
            
            $Update.PSTypeNames.Clear()
            $Update.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
            $UpdateCollection += $Update
        }Else{
            $objCollectionUpdate.Add($Update) | Out-Null
            $UpdatesExtraDataCollection.Add($Update.Identity.UpdateID,@{KB = $KB; Size = $size})
        }
    }
    $NumberOfUpdate++
}

Write-Progress -Activity "[1/$NumberOfStage] Post search updates" -Status "Completed" -Completed
If($ListOnly) {
    $FoundUpdatesToDownload = $UpdateCollection.count
}Else{
    $FoundUpdatesToDownload = $objCollectionUpdate.count
}
Write-host "Found [$FoundUpdatesToDownload] Updates in post search criteria"

If($FoundUpdatesToDownload -eq 0) {Return}
If($ListOnly) { Return $UpdateCollection }


# -------------------- STAGE 2 ---------------------------------------------
Write-Verbose "STAGE 2: Choose updates"
$NumberOfUpdate = 1
$logCollection = @()

$objCollectionChoose = New-Object -ComObject "Microsoft.Update.UpdateColl"

Foreach($Update in $objCollectionUpdate){
    $size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
    Write-Progress -Activity "[2/$NumberOfStage] Choose updates" -Status "[$NumberOfUpdate/$FoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$FoundUpdatesToDownload * 100))

    If($AcceptAll){
        $Status = "Accepted"
        If($Update.EulaAccepted -eq 0){ $Update.AcceptEula() }

        Write-Debug "Add update to collection"
        $objCollectionChoose.Add($Update) | Out-Null
    } ElseIf($AutoSelectOnly){
        If($Update.AutoSelectOnWebsites) {
            $Status = "Accepted"
            If($Update.EulaAccepted -eq 0) { $Update.AcceptEula() }
            $objCollectionChoose.Add($Update) | Out-Null
        } Else {
            $Status = "Rejected"
        }
    } Else {
        $Status = "Accepted"

        If($Update.EulaAccepted -eq 0) { $Update.AcceptEula() }
        $objCollectionChoose.Add($Update) | Out-Null 
    } Else {
        $Status = "Rejected"
    }

    Write-Debug "Add to log collection"
    $log = New-Object PSObject -Property @{
        Title = $Update.Title
        KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
        Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
        Status = $Status
    }
    $log.PSTypeNames.Clear()
    $log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
    $logCollection += $log
    $NumberOfUpdate++
}

Write-Progress -Activity "[2/$NumberOfStage] Choose updates" -Status "Completed" -Completed
Write-Debug "Show log collection"
$logCollection

$AcceptUpdatesToDownload = $objCollectionChoose.count
Write-Verbose "Accept [$AcceptUpdatesToDownload] Updates to Download"

If($AcceptUpdatesToDownload -eq 0) { Return}

# -------------------- STAGE 3 ---------------------------------------------
Write-Verbose "STAGE 3: Download updates"
$NumberOfUpdate = 1
$objCollectionDownload = New-Object -ComObject "Microsoft.Update.UpdateColl" 

Foreach($Update in $objCollectionChoose)  {
    Write-Progress -Activity "[3/$NumberOfStage] Downloading updates" -Status "[$NumberOfUpdate/$AcceptUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$AcceptUpdatesToDownload * 100))
    
    $objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
    $objCollectionTmp.Add($Update) | Out-Null
        
    $Downloader = $objSession.CreateUpdateDownloader() 
    $Downloader.Updates = $objCollectionTmp
    Try {
        Write-verbose "Try download update $($Update.Title)"
        $DownloadResult = $Downloader.Download()
    }
    Catch {
        If($_ -match "HRESULT: 0x80240044") {
            Write-Warning "Your security policy don't allow a non-administator identity to perform this task"
        }
        Return
    }
    
    Switch -exact ($DownloadResult.ResultCode) {
        0   { $Status = "NotStarted" }
        1   { $Status = "InProgress" }
        2   { $Status = "Downloaded" }
        3   { $Status = "DownloadedWithErrors" }
        4   { $Status = "Failed" }
        5   { $Status = "Aborted" }
    }

    $log = New-Object PSObject -Property @{
        Title = $Update.Title
        KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
        Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
        Status = $Status
    }

    $log

    If($DownloadResult.ResultCode -eq 2) { $objCollectionDownload.Add($Update) | Out-Null }
    $NumberOfUpdate++

}

Write-Progress -Activity "[3/$NumberOfStage] Downloading updates" -Status "Completed" -Completed

$ReadyUpdatesToInstall = $objCollectionDownload.count
Write-Verbose "Downloaded [$ReadyUpdatesToInstall] Updates to Install"

If($ReadyUpdatesToInstall -eq 0) { Return }

# -------------------- STAGE 4 ---------------------------------------------
If(!$DownloadOnly) {
    Write-Verbose "STAGE 4: Install updates"
    $NeedsReboot = $false
    $NumberOfUpdate = 1

    #install updates
    Foreach($Update in $objCollectionDownload) {
        Write-Progress -Activity "[4/$NumberOfStage] Installing updates" -Status "[$NumberOfUpdate/$ReadyUpdatesToInstall] $($Update.Title)" -PercentComplete ([int]($NumberOfUpdate/$ReadyUpdatesToInstall * 100))
        
        $objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
        $objCollectionTmp.Add($Update) | Out-Null
        
        $objInstaller = $objSession.CreateUpdateInstaller()
        $objInstaller.Updates = $objCollectionTmp

        Try {
            Write-Debug "Try install update $($Update.Title)"
            $InstallResult = $objInstaller.Install()
        }
        Catch {
            If($_ -match "HRESULT: 0x80240044") { Write-Warning "Your security policy don't allow a non-administator identity to perform this task"}
            Return
        }
        
        $NeedsReboot = $NeedsReboot -bor $installResult.RebootRequired 

        Switch -exact ($InstallResult.ResultCode)	{
            0   { $Status = "NotStarted"}
            1   { $Status = "InProgress"}
            2   { $Status = "Installed"}
            3   { $Status = "InstalledWithErrors"}
            4   { $Status = "Failed"}
            5   { $Status = "Aborted"}
        }

        Write-Debug "Add to log collection"
        $log = New-Object PSObject -Property @{
            Title = $Update.Title
            KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
            Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
            Status = $Status
        }
        $log.PSTypeNames.Clear()
        $log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
        $log

        $NumberOfUpdate++
    }
    Write-Progress -Activity "[4/$NumberOfStage] Installing updates" -Status "Completed" -Completed

    If($NeedsReboot){
        If($AutoReboot) {
            Restart-Computer -Force
        } ElseIf($IgnoreReboot) {
            Return "Reboot is required, but do it manually."
        } Else {
            $Reboot = Read-Host "Reboot is required. Do it now ? [Y/N]"
            If($Reboot -eq "Y")	{
                Restart-Computer -Force
            }
        }
    }
}

# -----------------------------------------------------------------
#       fin du script
# -----------------------------------------------------------------
Stop-Transcript | Out-Null
