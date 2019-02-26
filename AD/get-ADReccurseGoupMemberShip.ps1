param (
    $Utilisateur
)

$script:listeGroup = @()

function Affiche-Groupe {
    param ($g, $niveau)

    if ($listeGroup -contains $g.name) {
        $couleur="red"
    }else{
        $couleur="green"
    }

    $script:listeGroup += $g.name
    $n="  "* $niveau 
    $chaine = "{0}|__ {1,-30}" -f $n,$g.name
    write-host $chaine -ForegroundColor $couleur

    $niveau+=1
    foreach ( $subg in $(Get-ADPrincipalGroupMembership $g.name))  {
        Affiche-groupe $subg $niveau
    }
}

foreach ( $g in $(Get-ADPrincipalGroupMembership $Utilisateur) ) {
 Affiche-groupe $g 1
}
