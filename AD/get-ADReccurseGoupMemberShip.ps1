 <#
.NOTES

Auteur :     Erwan Le Tiec
Creation:   26/02/2019
Modif:

 .SYNOPSIS
affiche les groupes d'appartenance d'un utilisateur ActiveDir de façon reccursive

 .DESCRIPTION
affiche les groupes d'appartenance d'un utilisateur ActiveDir de façon reccursive.
les groupes en doublon apparaissent en rouge.

 .PARAMETER version
 affiche le numéro de version et se termine sans rien faire

.PARAMETER user
utilisateur activedirectory 
#>

# -----------------------------------------------------------------
# Gestion des paramètres et variables propres au script
# -----------------------------------------------------------------
[cmdletBinding()]
param (
    [string] $User,
    [switch] $version= $false
)

$scriptversion = 1.0

$script:listeGroup = @()

# -----------------------------------------------------------------
# Fonction utilisables
# -----------------------------------------------------------------
function Affiche-Groupe {
    param ($g, $niveau)

    if ($listeGroup -contains $g.name) {
        $couleur="red"
    }else{
        $couleur="green"
    }

    $script:listeGroup += $g.name
    $n="  "* $niveau 
    $chaine = "{0}|___ {1,-30}" -f $n,$g.name
    write-host $chaine -ForegroundColor $couleur

    $niveau+=1
    foreach ( $subg in $(Get-ADPrincipalGroupMembership $g.name))  {
        Affiche-groupe $subg $niveau
    }
}
 

# -----------------------------------------------------------------
# Variables génériques
# -----------------------------------------------------------------
$scriptName = $myInvocation.MyCommand.Name
$scriptLongName = $myInvocation.MyCommand.path

# -----------------------------------------------------------------
#	pre-traitements
# -----------------------------------------------------------------
if ($version) {
    Write-Host "`nnom                   : $scriptName"
    write-host "version               : $scriptversion"
    write-host "chemin d'installation : $scriptLongName `n"
    exit 0
    }

# -----------------------------------------------------------------
#	traitements
# -----------------------------------------------------------------


foreach ( $g in $(Get-ADPrincipalGroupMembership $User) ) {
    Affiche-groupe $g 1
   }

# -----------------------------------------------------------------
#	fin du script
# -----------------------------------------------------------------
