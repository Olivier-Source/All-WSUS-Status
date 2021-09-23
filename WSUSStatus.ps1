#############################################
# Author: Olivier CHALLET
# Creation Date: 21/9/2021
# Last Modified: 21/9/2021
# Version: 1.0
# Description: Script de génération de rapport WSUS
# Attention: Le script est conçu pour être exécuter toutes les semaines via tache plannifiée
#############################################

# Initialisation des variables de base habituelles
Param (
[string]$UpdateServer = 'WSUS-ServerName', # Nom du serveur WSUS.
[int]$Port = 8530, # Port TCP du WSUS. 8530 par defaut.
[bool]$Secure = $False # TRUE si on utilise le HTTPS.
)

# Initialisation des variables pour envoi de Mail
$SMTPServer = "SMTP-Server" # Nom du serveur SMTP
$EmailRecipient = "myemail@domain.com"
$EmailSender = $UpdateServer + "<" + $UpdateServer + "@domain.com>"

# MsgBody est la variable contenant le corps du mail
$MsgBody = "<HTML>"
$MsgBody = $MsgBody + "<HEAD>"
$MsgBody = $MsgBody + '<META http-equiv="Content-Type" content="text/html; charset=UTF-8" />'
$MsgBody = $MsgBody + "<title>Tous les rapports du serveur WSUS :" + $UpdateServer + "</title>"
$MsgBody = $MsgBody + "</HEAD>"
$MsgBody = $MsgBody + "<BODY style=""font-family:'Courier New', Courier, monospace"">"

$MsgBody= $MsgBody + "<h1>Rapport 'Ordinateur' du serveur WSUS : " + $UpdateServer + "</h1>"
$intLineCounter = 0 # Variable qui compte les ordinateurs.

Remove-Item -force ./Logs/PcStatusReport.txt # Petit fichier de log.

If (-Not (Import-Module UpdateServices -PassThru)) {
Add-Type -Path "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll" -PassThru
}

$Wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($UpdateServer,$Secure,$Port) # Connexion au serveur WSUS.

$CTScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope #Scope visé, incluant tous les ordinateurs.

# On continu en écrivant les résultats dans le tableau.
$MsgBody = $MsgBody + "<table border=""0"" cellspacing=""2"" cellpadding=""2"" style=""font-family:'Courier New', Courier, monospace"">"
$MsgBody = $MsgBody + "<tr>"
$MsgBody = $MsgBody + "<th>Index</th>"
$MsgBody = $MsgBody + "<th>Statut</th>"
$MsgBody = $MsgBody + "<th>Nom serveur</th>"
$MsgBody = $MsgBody + "<th>Addresse IP</th>"
$MsgBody = $MsgBody + "<th>Dernier Contact</th>"
$MsgBody = $MsgBody + "<th>Total updates</th>"
$MsgBody = $MsgBody + "<th bgcolor=""#ff8a65"">Attente reboot</th>"
$MsgBody = $MsgBody + "<th bgcolor=""#fb8c00"">Prêt pour installation</th>"
$MsgBody = $MsgBody + "<th bgcolor=""#fdd835"">Téléchargement en attente</th>"
$MsgBody = $MsgBody + "<th bgcolor=""#B00020"">En erreur</th>"
$MsgBody = $MsgBody + "<th bgcolor=""#e0e0e0"">Statut inconnu</th>"
$MsgBody = $MsgBody + "</tr>"

$NbOk=0; $NbRR=0; $NbPI=0; $NbTA=0; $NbErr=0; $NbNR=0; $NbNT=0

# Grosse partie : Ici on tri les ordinateurs par nom et on récupère les détails de chacun d'entre eux.
$wsus.GetComputerTargets($CTScope) | Sort -Property FullDomainName | ForEach {

    $objSummary = $_.GetUpdateInstallationSummary() # Objet intermédiaire contenant les détails.
    $Down = $objSummary.DownloadedCount # This is the amount of updates that has been downloaded already.
    $Fail = $objSummary.FailedCount # This is the count for the failed updates.
    $Pend = $objSummary.InstalledPendingRebootCount # This is the number of updates that need to reboot to complete installation.
    $NotI = $objSummary.NotInstalledCount # These are the needed updates for this computer.
    $Unkn = $objSummary.UnknownCount # These are the updates that are waiting for detection on the first search.
    $Total = $Down + $Fail + $Pend + $NotI + $Unkn # Total amount of updates for this computer.

    $intLineCounter = $intLineCounter + 1 # On incrémente le compteur de ligne.
    $IntStr = [Convert]::ToString($intLineCounter) # On le converti en chaine de caractère pour le code HTML.

    if ($Total -eq 0) {$Estado="OK"; $bgcolor="#8bc34a"; $NbOk=$NbOk+1}
    elseif ($Pend -ne 0) {$Estado="Redémarrage requis"; $bgcolor="#ff8a65"; $NbRR=$NbRR+1}
    elseif ($Down -ne 0) {$Estado="Prêt pour installation"; $bgcolor="#fb8c00"; $NbPI=$NbPI+1}
    elseif ($NotI -ne 0) {$Estado="Téléchargement en attente"; $bgcolor="#fdd835"; $NbTA=$NbTA+1}
    elseif ($Fail -ne 0) {$Estado="Erreur"; $bgcolor="#B00020"; $NbErr=$NbErr+1}
    elseif ($Unkn -ne 0) {$Estado="Pas de rapport"; $bgcolor="#e0e0e0"; $NbNR=$NbNR+1}
    else {$Estado="Erreur de statut"; $bgcolor="White"; $NbNT=$NbNT+1}

    Write-Verbose ($IntStr + " : " + $_.FullDomainName) -Verbose # Affichage d'une barre de progression pour la beauté du geste.

    $LastContact = $_.LastReportedStatusTime # Dernière fois d'un ordinateur à fait un rapport au serveur WSUS.
    $days = [Math]::Ceiling((New-TimeSpan -Start $LastContact).TotalDays) # Nombre de jours depuis.

    if ($days -gt 27) {$Color="#B00020"} # Ordinateur absent depuis trop longtemps. (28 jours)
    elseif ($days -gt 13) {$Color="#ff8a65"} # Ordinateur ayant potentiellement un problème. (pas de rapport depuis 14 jours)
    elseif ($days -gt 2) {$Color="#fdd835"} # Ordinateur ayant potentiellement un problème. (pas de rapport depuis 2 jours)
    else { # Ordinateur OK
        if ($intLineCounter%2) {
            $Color="#eeeeee"
        } else {
            $Color="White"
        }
    } 

    # Formatage de la date.
    if ($days -eq 0) {$Dias="Aujourd'hui"}
    elseif ($days -eq 1) {$Dias="Hier"}
    else {$Dias="Depuis " + $days + " jours."}

    if ($LastContact -eq [DateTime]::MinValue) {$Dias="Jamais"; $Color="Silver"}

    # Et on ecrit les infos dans le tableau.

    if ($intLineCounter%2) {
        $MsgBody = $MsgBody + "<tr style=""background-color:#eeeeee"">"
    } else {
        $MsgBody = $MsgBody + "<tr>"
    }

    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle""> " + $IntStr +" </td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"" bgcolor=""" + $bgcolor + """> " + $Estado + " </td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle""> " + $_.FullDomainName+ " </td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle""> " + $_.IPAddress + " </td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"" bgcolor=""" + $Color + """> " + $Dias +"</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $Total + "</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $Pend + "</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $Down + "</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $NotI + "</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $Fail + "</td>"
    $MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">" + $Unkn + "</td>"
    $MsgBody = $MsgBody + "</tr>"

    $_.FullDomainName >> ./Logs/PcStatusReport.txt # Ajout de la ligne dans le fichier de log
}

$MsgBody = $MsgBody + "</table><br>" # On fini le tableau.

# Tableau récapitulatif des statuts et du pourcentage de chaque statut
$MsgBody = $MsgBody + "<br><center>"

$MsgBody = $MsgBody + "<table border=""0"" cellspacing=""2"" cellpadding=""2"" style=""font-family:'Courier New', Courier, monospace"">"

$MsgBody = $MsgBody + "<tr>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle""></td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">OK</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"" style=""background-color:#eeeeee"">Redémarrage requis</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">Prêt pour installation</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"" style=""background-color:#eeeeee"">Téléchargement en attente</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">Erreur</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"" style=""background-color:#eeeeee"">Pas de rapport</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">Total calculé / Total recensé</td>"
$MsgBody = $MsgBody + "</tr>"

$MsgBody = $MsgBody + "<tr style=""background-color:#eeeeee"">"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">Nombre</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbOk</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbRR</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbPI</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbTA</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbErr</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">$NbNR</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ ($NbOk + $NbRR + $NbPI + $NbTA + $NbErr + $NbNR) +" / "+ $intLineCounter +"</td>"
$MsgBody = $MsgBody + "</tr>"

$MsgBody = $MsgBody + "<tr>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">Pourcentage</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbOk*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbRR*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbPI*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbTA*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbErr*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round(($NbNR*100)/$intLineCounter,2) +"%</td>"
$MsgBody = $MsgBody + "<td align=""center"" valign=""middle"">"+ [Math]::Round((($NbOk + $NbRR + $NbPI + $NbTA + $NbErr + $NbNR)*100)/$intLineCounter,2) +"% / 100%</td>"
$MsgBody = $MsgBody + "</tr>"

$MsgBody = $MsgBody + "</table></center><br>" # On fini le tableau.

if ($intLineCounter -eq 0) {
    Write-Verbose ("You must run this script from as administrator to read WSUS database.") -Verbose # Affichage d'un message d'alerte pour prévenir les boulets sans droits.
    $MsgBody = $MsgBody + " Vous devez exécuter le script en administrateur afin de pouvoir lire la base de données. <hr>"
} else {
    $MsgBody = $MsgBody + "<hr>"
}

#Petite note de pied de page.
$MsgBody = $MsgBody + "<p><h2>Note: </h2>Les mises à jour sont appliquées selon un processus en trois parties : Recherche, Téléchargement & Installation.<br> Chaque ordinateur doit passer par ces trois étapes pour être mis à jour.<ul>"
$MsgBody = $MsgBody + "<li>Avant que la première recherche ne soit effectuée le statut '<strong>Inconnu</strong>' est affecté par défaut par l'ordinateur ne sait pas s'il est à jour ou non.</li>"
$MsgBody = $MsgBody + "<li>Une fois la recherche complète les updates apparaissent avec le statut <strong>Téléchargement en attente</strong> avec chaque update en attente d'installation.</li>"
$MsgBody = $MsgBody + "<li>Quand l'ordinateur a téléchargé les MAJ WSUS il est considéré comme <strong>Prêt pour installation</strong>.</li>"
$MsgBody = $MsgBody + "<li>Et à la fin de l'installation pour le stage en cours, vous pouvez avoir les statuts suivants :</li><ul>"
$MsgBody = $MsgBody + "<li><strong>OK</strong>: Toutes les MAJ WSUS sont installées correctement.</li>"
$MsgBody = $MsgBody + "<li><strong>ERREUR</strong>: Une ou plusieurs MAJ ne sont pas correctement installés.</li>"
$MsgBody = $MsgBody + "<li><strong>Redémarrage requis</strong>: Les installations sont faites, mais un redémarrage est nécessaire.</li>"
$MsgBody = $MsgBody + "</ul></ul></p>"
$MsgBody = $MsgBody + "Si un ordinateur ne s'est pas connecté au serveur WSUS depuis plus de 28 jours il est souligné en <font style=""background-color:#B00020"">rouge</font>.<br>"
$MsgBody = $MsgBody + "Si un ordinateur ne s'est pas connecté au serveur WSUS depuis plus de 14 jours et jusqu'à 28 jours il est souligné en <font style=""background-color:#ff8a65"">orange</font>.<br>"
$MsgBody = $MsgBody + "Si un ordinateur ne s'est pas connecté au serveur WSUS depuis plus de 2 jours et jusqu'à 14 jours il est souligné en <font style=""background-color:#fdd835"">jaune</font>.<br>"
$MsgBody = $MsgBody + "<hr>"

$MsgBody = $MsgBody + "<center><strong>Généré le : " + [System.DateTime]::Now + "</strong></center>" # This is a timestamp at the end of the message, taking into account the SMTP delivery delay.

$IntStr = [Convert]::ToString($intLineCounter) # convert the line counter to string
$subject = "Rapport de " + $IntStr + " ordinateurs enregistrés sur le serveur " + $UpdateServer # This will be the subject of the message.

# Closing Body and HTML tags
$MsgBody = $MsgBody + "</BODY>"
$MsgBody = $MsgBody + "</HTML>"

# Now send the email message and thats all.
send-MailMessage -SmtpServer $SMTPServer -To $EmailRecipient -From $EmailSender -Subject $subject -Body $MsgBody -BodyAsHtml -Priority high -Encoding UTF8
