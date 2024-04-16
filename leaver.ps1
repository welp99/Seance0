# Définition des scopes et connexion à Graph
$scopes = "Sites.ReadWrite.All"
Connect-MgGraph -Scopes $scopes -NoWelcome

# Identifiants du site et des listes
$siteId = "d8110319-1192-43e0-a400-451dfb2b4f97"
$listIdJOINER = "e5a1a0dc-5f9b-41c1-abd0-5d5c50c33529"
$listIdLEAVER = "7245c8cd-4fd4-4af3-a7ff-c5a79858880a"

# Récupération des utilisateurs à désactiver
$LEAVER = Get-MgSiteListItem -SiteId $siteId -ListId $listIdLEAVER -ExpandProperty Fields
foreach ($user in $LEAVER)
{
    # Désactivation de l'utilisateur dans Active Directory
    $userSamAccountName = $user.Fields.AdditionalProperties.NOM + "." + $user.Fields.AdditionalProperties.PRENOM
    Set-ADUser -Identity $userSamAccountName -Enabled $false
    Write-Host "Utilisateur désactivé : $($userSamAccountName)`n"
}
