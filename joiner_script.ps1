Install-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Sites

Install-Module Microsoft.Graph.Users.Actions
Import-Module Microsoft.Graph.Users.Actions

$scopes = "Sites.ReadWrite.All", "Mail.Send"
Connect-MgGraph -Scopes $scopes -NoWelcome
$siteId = "d8110319-1192-43e0-a400-451dfb2b4f97"
$listIdJOINER = "e5a1a0dc-5f9b-41c1-abd0-5d5c50c33529"
$listIdLEAVER = "7245c8cd-4fd4-4af3-a7ff-c5a79858880a"

$JOINER = Get-MgSiteListItem -SiteId $siteId -ListId $listIdJOINER -ExpandProperty Fields
foreach ($user in $JOINER)
{
    if($user.Fields.AdditionalProperties.TRAITEPOWERSHELL -eq "0")
    {
        Write-Host "Nom de l'utilisateur : $($user.Fields.AdditionalProperties.NOM)"
        Update-MgSiteListItemField -SiteId $siteId -ListId $listIdJOINER -ListItemId $user.Id -BodyParameter @{"TRAITEPOWERSHELL"="1"}

        $user.ID
        $user.Fields.AdditionalProperties.NOM
        $user.Fields.AdditionalProperties.PRENOM
        $user.Fields.AdditionalProperties.MAIL
        $user.Fields.AdditionalProperties.DEBUTCONTRAT
        $user.Fields.AdditionalProperties.POSTE
        $user.Fields.AdditionalProperties.SERVICE
        $user.Fields.AdditionalProperties.VILLE
        $user.Fields.AdditionalProperties.MAILSUPERIEUR
        $user.Fields.AdditionalProperties.INFOSCOMPL
        $user.Fields.AdditionalProperties.TRAITEPOWERSHELL
        Write-Host("`n`n")

        New-ADUser -Name ($user.Fields.AdditionalProperties.NOM + " " + $user.Fields.AdditionalProperties.PRENOM) `
            -GivenName $user.Fields.AdditionalProperties.NOM `
            -Surname $user.Fields.AdditionalProperties.PRENOM `
            -Path "OU=New_User,OU=PRODUCTION,DC=ecoles-epsi,DC=net" `
            -SamAccountNam ($user.Fields.AdditionalProperties.NOM + "." + $user.Fields.AdditionalProperties.PRENOM) `
            -EmailAddress $user.Fields.AdditionalProperties.MAIL `
            -Department $user.Fields.AdditionalProperties.SERVICE `
            -Title $user.Fields.AdditionalProperties.POSTE

        # Définir les informations du nouvel arrivant, du support et du N+1
        $nouvelArrivant = @{
            Prenom = $user.Fields.AdditionalProperties.PRENOM
            Nom = $user.Fields.AdditionalProperties.NOM
            Email = $user.Fields.AdditionalProperties.MAIL
            Poste = $user.Fields.AdditionalProperties.POSTE
            DateArrivee = $user.Fields.AdditionalProperties.DEBUTCONTRAT
            Telephone = $user.Fields.AdditionalProperties.INFOSCOMPL
        }

        $support = @{
            Email = "emmanuel.savadogo@ecoles-epsi.net"
        }

        $nPlusUn = @{
            Email = $user.Fields.AdditionalProperties.MAILSUPERIEUR
        }

        # Définir les paramètres du mail
        $params = @{
            Message = @{
                Subject = "Notification PowerAutomate : Compte créé pour $($nouvelArrivant.Prenom) $($nouvelArrivant.Nom)"
                Body = @{
                    ContentType = "HTML"
                    Content = @"
Bonjour,<br><br>

Le support technique vous informe que le compte pour <b>$($nouvelArrivant.Prenom) $($nouvelArrivant.Nom)</b> a été créé avec succès.
Voici ses coordonnées et informations :
<ul>
<li>Poste : $($nouvelArrivant.Poste)</li>
<li>Date d'arrivée : $($nouvelArrivant.DateArrivee)</li>
<li>E-mail : <a href="mailto:$($nouvelArrivant.Email)">$($nouvelArrivant.Email)</a></li>
<li>Téléphone : $($nouvelArrivant.Telephone)</li>
</ul>

Nous vous prions de bien vouloir l'accueillir chaleureusement et de lui fournir tout le support nécessaire pour faciliter son intégration.

Cordialement,<br>
Le Support Technique
"@
                }
                From = @{
                    EmailAddress = @{
                        Address = $($support.Email)
                    }
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $($support.Email)
                        }
                    }
                )
                CcRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $($nPlusUn.Email)
                        }
                    }
                )
            }
        }

        # Envoyer le mail
        Send-MgUserMail -UserId $support.Email -BodyParameter $params
    }
}