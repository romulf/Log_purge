<#
.NOTES
Objet:           WSUS

Auteur:          Erwan Le Tiec
Creation:        17/05/2017
Modif:


.SYNOPSIS
    Get list of available updates meeting the criteria.

	.DESCRIPTION
	    Use Get-WUList to get list of available or installed updates meeting specific criteria.
		There are two types of filtering update: Pre search criteria, Post search criteria.
		- Pre search works on server side, like example: (IsInstalled = 0 and IsHidden = 0 )
		- Post search work on client side after downloading the pre-filtered list of updates, like example $KBArticleID -match $Update.KBArticleIDs

		Status list:
        D - IsDownloaded, I - IsInstalled, M - IsMandatory, H - IsHidden, U - IsUninstallable, B - IsBeta

	.PARAMETER UpdateType
		Pre search criteria. Finds updates of a specific type, such as 'Driver' and 'Software'. Default value contains all updates.

	.PARAMETER IsInstalled
		Pre search criteria. Finds updates that are installed on the destination computer.

	.PARAMETER IsHidden
		Pre search criteria. Finds updates that are marked as hidden on the destination computer.
	
	.PARAMETER IsNotHidden
		Pre search criteria. Finds updates that are not marked as hidden on the destination computer. Overwrite IsHidden param.
			
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
	
	.PARAMETER AutoSelectOnly  
		Post search criteria. finds updates that have are flagged to be automatically selected by Windows Update.

 .PARAMETER version
	display script version and exit

 .PARAMETER help
	display this help message

	.EXAMPLE
		Get list of available updates from default Update Server.
	
PS C:\>Get-WUList.ps1
Fetching list from Windows Server Update Service.  please wait...
Found [1] Updates in pre search criteria
Found [1] Updates in post search criteria

KB                            size                          Title                         status
--                            ----                          -----                         ------
KB3172729                     10 MB                         Security Update for Window... ------

	.EXAMPLE
		find only updates with Monthly in title
PS C:\> Get-WUList.ps1  -Title "Monthly" | ft  -AutoSize
Fetching list from Windows Server Update Service.  please wait...
Found [80] Updates in pre search criteria
Found [1] Updates in post search criteria

KB        status size   Title
--        ------ ----   -----
KB4015549 ------ 160 MB April, 2017 Security Monthly Quality Rollup for Windows Server 2008 R2 for x64-based Systems (KB4015549)

.EXAMPLE
		Get list of updates without language packs and updatets that's not hidden.
	
		PS C:\> Get-WUList -NotCategory "Language packs" -IsNotHidden

		ComputerName Status KB          Size Title
		------------ ------ --          ---- -----
		G1           ------ KB2640148   8 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2640148)
		G1           ------ KB2600217  32 MB Aktualizacja dla programu Microsoft .NET Framework 4 w systemach Windows XP, Se...
		G1           ------ KB2679255   6 MB Aktualizacja systemu Windows 7 dla komputerów z procesorami x64 (KB2679255)
		G1           ------ KB915597    3 MB Definition Update for Windows Defender - KB915597 (Definition 1.125.146.0)

#>


# -----------------------------------------------------------------
# Gestion des parametres et variables propres au script
# -----------------------------------------------------------------

[CmdletBinding()]	
Param (
	#Pre search criteria
	[ValidateSet("Driver", "Software")]	[String]$UpdateType="",
	[Switch]$IsInstalled,
	[Switch]$IsHidden,
	[Switch]$IsNotHidden,
	
	#Post search criteria
	[String[]]$Category="",
	[String[]]$KBArticleID,
	[String]$Title,
	
	[String[]]$NotCategory="",
	[String[]]$NotKBArticleID,
	[String]$NotTitle,	
	
	[Alias("Silent")]	[Switch]$IgnoreUserInput,
	[Switch]$IgnoreRebootRequired,
	[Switch]$AutoSelectOnly,

	[switch] $version= $false,
	[switch] $help=$false
)

$UpdateCollection = @()
$defaultProperties = @("KB","status","size","Title")
$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

$Scriptversion=1.0


# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$scriptName = $myInvocation.MyCommand.Name       # get-WUHistory.ps1
$scriptLongName = $myInvocation.MyCommand.path   # D:\sources\claranet\get-WUHistory.ps1

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
    
if ($help) {
  Get-Help $scriptLongName -full 
  exit 0
}

$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

if(!$Role) {
	Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."	
}


# -----------------------------------------------------------------
#       traitements
# -----------------------------------------------------------------
$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
$objSession = New-Object -ComObject "Microsoft.Update.Session"
$objSearcher = $objSession.CreateUpdateSearcher()

Foreach ($objService in $objServiceManager.Services) {
	If($objService.IsDefaultAUService -eq $True) {
		$serviceName = $objService.Name
		Break
	}
} 

Write-Verbose "Set source of updates to $serviceName"
Try	{
	$search = ""

	If($IsInstalled) {
		$search = "IsInstalled = 1"
		Write-Verbose "Set pre search criteria: IsInstalled = 1"
	} Else {
		$search = "IsInstalled = 0"	
		Write-Verbose "Set pre search criteria: IsInstalled = 0"
		}

	If($UpdateType -ne "") {
		Write-Verbose "Set pre search criteria: Type = $UpdateType"
		$search += " and Type = '$UpdateType'"
	}

	If($CategoryIDs) {
		Write-Verbose "Set pre search criteria: CategoryIDs = '$([string]::join(", ", $CategoryIDs))'"
		$tmp = $search
		$search = ""
		$LoopCount =0
		Foreach($ID in $CategoryIDs) {
			If($LoopCount -gt 0) { $search += " or "}
			$search += "($tmp and CategoryIDs contains '$ID')"
			$LoopCount++
		}
	}

	If($IsNotHidden) {
		Write-Verbose "Set pre search criteria: IsHidden = 0"
		$search += " and IsHidden = 0"	
	} ElseIf($IsHidden) {
		Write-Verbose "Set pre search criteria: IsHidden = 1"
		$search += " and IsHidden = 1"	
	} #End ElseIf $IsHidden

	If($IgnoreRebootRequired) {
		Write-Verbose "Set pre search criteria: RebootRequired = 0"
		$search += " and RebootRequired = 0"	
	}
			
	Write-Verbose "Search criteria is: $search"
	write-host "Fetching list from $serviceName.  please wait..."		
	$objResults = $objSearcher.Search($search)
} #End Try

Catch {
	If($_ -match "HRESULT: 0x80072EE2")	{
		Write-Warning "Probably you don't have connection to Windows Update server"
	} #End If $_ -match "HRESULT: 0x80072EE2"
	Return
} #End Catch


$NumberOfUpdate = 1
$PreFoundUpdatesToDownload = $objResults.Updates.count
Write-host "Found [$PreFoundUpdatesToDownload] Updates in pre search criteria"
If($PreFoundUpdatesToDownload -eq 0) { Continue }

Foreach($Update in $objResults.Updates) {
	$UpdateAccess = $true
	Write-Progress -Activity "Post search updates" -Status "[$NumberOfUpdate/$PreFoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$PreFoundUpdatesToDownload * 100))
	Write-Verbose "Set post search criteria: $($Update.Title)"

	If($Category -ne "") {
		$UpdateCategories = $Update.Categories | Select-Object Name
		Write-Verbose "Set post search criteria: Categories = '$([string]::join(", ", $Category))'"	
		Foreach($Cat in $Category) {
			If(!($UpdateCategories -match $Cat)) {
				Write-Verbose "UpdateAccess: false"
				$UpdateAccess = $false
			} Else {
				$UpdateAccess = $true
				Break
			}
		}
	}

	If($NotCategory -ne "" -and $UpdateAccess -eq $true) {
		$UpdateCategories = $Update.Categories | Select-Object Name
		Write-Verbose "Set post search criteria: NotCategories = '$([string]::join(", ", $NotCategory))'"	
		Foreach($Cat in $NotCategory) {
			If($UpdateCategories -match $Cat) {
				Write-Verbose "UpdateAccess: false"
				$UpdateAccess = $false
				Break
			}
		}
	}

	If($KBArticleID -ne $null -and $UpdateAccess -eq $true)	{
		Write-Verbose "Set post search criteria: KBArticleIDs = '$([string]::join(", ", $KBArticleID))'"
		If(!($KBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs)) {
			Write-Verbose "UpdateAccess: false"
			$UpdateAccess = $false
		}
	}

	If($NotKBArticleID -ne $null -and $UpdateAccess -eq $true){
		Write-Verbose "Set post search criteria: NotKBArticleIDs = '$([string]::join(", ", $NotKBArticleID))'"
		If($NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs) {
				Write-Verbose "UpdateAccess: false"
				$UpdateAccess = $false
		}
	}

	If($Title -and $UpdateAccess -eq $true) {
		Write-Verbose "Set post search criteria: Title = '$Title'"
		If($Update.Title -notmatch $Title) {
			Write-Verbose "UpdateAccess: false"
			$UpdateAccess = $false
		}
	}

	If($NotTitle -and $UpdateAccess -eq $true)	{
		Write-Verbose "Set post search criteria: NotTitle = '$NotTitle'"
		If($Update.Title -match $NotTitle) {
			Write-Verbose "UpdateAccess: false"
			$UpdateAccess = $false
		}
	}

	If($IgnoreUserInput -and $UpdateAccess -eq $true) {
		Write-Verbose "Set post search criteria: CanRequestUserInput"
		If($Update.InstallationBehavior.CanRequestUserInput -eq $true) {
			Write-Verbose "UpdateAccess: false"
			$UpdateAccess = $false
		}
	}

	If($IgnoreRebootRequired -and $UpdateAccess -eq $true) {
		Write-Verbose "Set post search criteria: RebootBehavior"
		If($Update.InstallationBehavior.RebootBehavior -ne 0) {
				Write-Verbose "UpdateAccess: false"
				$UpdateAccess = $false
		}
	}

	If($AutoSelectOnly -and $UpdateAccess -eq $true) {
		Write-Verbose "Set post search criteria: AutoSelectOnWebsites"
		If($Update.AutoSelectOnWebsites -ne $true) {
			Write-Verbose "UpdateAccess: false"
			$UpdateAccess = $false
		}
	}


	If($UpdateAccess -eq $true)	{
		Switch($Update.MaxDownloadSize)	{
			{[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
			{[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
			{[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
			{[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
			default { $size = $_+"B" }
		}
			
		If($Update.KBArticleIDs -ne "")	{
			$KB = "KB"+$Update.KBArticleIDs
		} Else {
			$KB = ""
		}

		$Status = ""
		If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
		If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
		If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
		If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
		If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
		If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 

		Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $Computer
		Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
		Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
		Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
		Add-Member -InputObject $Update MemberSet PSStandardMembers $PSStandardMembers	

		$UpdateCollection += $Update
	}

	$NumberOfUpdate++
}
Write-Progress -Activity "Post search updates" -Status "Completed" -Completed

$FoundUpdatesToDownload = $UpdateCollection.count
Write-host "Found [$FoundUpdatesToDownload] Updates in post search criteria"

Return $UpdateCollection

# -----------------------------------------------------------------
#       fin du script
# -----------------------------------------------------------------