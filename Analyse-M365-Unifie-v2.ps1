# Script unifie d'analyse Microsoft 365 avec gestion complete des connexions
# Version corrigee - tout-en-un : connexion, analyse, deconnexion optionnelle

param(
    [string]$CheminRapport = "",
    [string]$TenantCible = "",
    [string]$EmailAdmin = ""
)

# Configuration de l'encodage
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "======================================================" -ForegroundColor Magenta
Write-Host "    ANALYSE STOCKAGE MICROSOFT 365 - VERSION UNIFIEE" -ForegroundColor Magenta
Write-Host "======================================================" -ForegroundColor Magenta

# Fonction pour convertir la taille en GB
function Convert-SizeToGB {
    param($SizeString)
    
    if ($null -eq $SizeString -or $SizeString -eq "") {
        return 0
    }
    
    $SizeStr = $SizeString.ToString()
    
    if ($SizeStr -match '([\d,\.]+)\s*([KMGT]?B)') {
        $Value = [double]($matches[1] -replace ',', '')
        $Unit = $matches[2]
        
        switch ($Unit) {
            "B" { return [math]::Round($Value / 1GB, 4) }
            "KB" { return [math]::Round($Value / 1MB, 4) }
            "MB" { return [math]::Round($Value / 1KB, 4) }
            "GB" { return [math]::Round($Value, 4) }
            "TB" { return [math]::Round($Value * 1KB, 4) }
            default { return 0 }
        }
    }
    
    try {
        $bytes = [double]$SizeStr.Split('(')[1].Split(' ')[0].Replace(',', '')
        return [math]::Round($bytes / 1GB, 4)
    }
    catch {
        return 0
    }
}

# Fonction de deconnexion PowerShell uniquement - VERSION CORRIGEE
function Disconnect-PowerShellSessions {
    param([bool]$Confirm = $true)
    
    if ($Confirm) {
        Write-Host "`nDeconnexion des sessions PowerShell (applications locales inchangees)" -ForegroundColor Yellow
        $Response = Read-Host "Confirmer la deconnexion PowerShell? (O/N)"
        if ($Response.ToUpper() -ne 'O') {
            Write-Host "Deconnexion annulee" -ForegroundColor Gray
            return $false
        }
    }
    
    Write-Host "Deconnexion des sessions PowerShell en cours..." -ForegroundColor Yellow
    
    # 1. Deconnexion Exchange Online - METHODE RENFORCEE
    try {
        # Verifier s'il y a des sessions Exchange
        $ExoSessions = Get-PSSession | Where-Object {
            $_.ConfigurationName -eq "Microsoft.Exchange" -or 
            $_.ComputerName -like "*outlook.office365.com*" -or
            $_.ComputerName -like "*protection.outlook.com*"
        }
        
        if ($ExoSessions) {
            Write-Host "  Sessions Exchange detectees: $($ExoSessions.Count)" -ForegroundColor Cyan
            foreach ($Session in $ExoSessions) {
                Write-Host "    Suppression session: $($Session.ComputerName)" -ForegroundColor Gray
                Remove-PSSession $Session -ErrorAction SilentlyContinue
            }
        }
        
        # Forcer la deconnexion Exchange
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            Write-Host "  Exchange Online: DECONNECTE (commande)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Exchange Online: Deconnexion forcee (sessions supprimees)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Exchange: Erreur lors de la deconnexion: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # 2. Deconnexion SharePoint Online
    try {
        Disconnect-SPOService -ErrorAction Stop
        Write-Host "  SharePoint Online: DECONNECTE" -ForegroundColor Green
    }
    catch {
        Write-Host "  SharePoint: Deja deconnecte ou erreur mineure" -ForegroundColor Gray
    }
    
    # 3. Suppression FORCEE de toutes les sessions PowerShell distantes
    try {
        $AllRemoteSessions = Get-PSSession
        if ($AllRemoteSessions) {
            Write-Host "  Suppression de toutes les sessions distantes: $($AllRemoteSessions.Count)" -ForegroundColor Yellow
            foreach ($Session in $AllRemoteSessions) {
                Remove-PSSession $Session -ErrorAction SilentlyContinue
                Write-Host "    Session supprimee: $($Session.Name) - $($Session.ComputerName)" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "  Erreur lors du nettoyage des sessions: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # 4. VERIFICATION FINALE RENFORCEE
    Write-Host "`nVerification finale des deconnexions:" -ForegroundColor Cyan
    
    # Test Exchange - doit echouer si bien deconnecte
    $ExchangeDeconnecte = $false
    try {
        $TestOrg = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "  PROBLEME: Exchange semble encore connecte!" -ForegroundColor Red
        Write-Host "    Organisation: $($TestOrg.DisplayName)" -ForegroundColor Red
        
        # Tentative de deconnexion forcee supplementaire
        Write-Host "  Tentative de deconnexion forcee..." -ForegroundColor Yellow
        try {
            # Supprimer toutes les sessions Exchange residuelles
            Get-PSSession | Where-Object {
                $_.ConfigurationName -eq "Microsoft.Exchange" -or
                $_.State -eq "Opened"
            } | Remove-PSSession -Force
            
            # Re-tenter la deconnexion
            Disconnect-ExchangeOnline -Confirm:$false -Force -ErrorAction SilentlyContinue
            
            # Re-test
            $null = Get-OrganizationConfig -ErrorAction Stop
            Write-Host "  ATTENTION: Exchange toujours connecte apres deconnexion forcee!" -ForegroundColor Red
        }
        catch {
            Write-Host "  Exchange: DECONNECTE apres tentative forcee" -ForegroundColor Green
            $ExchangeDeconnecte = $true
        }
    }
    catch {
        Write-Host "  Exchange: DECONNECTE" -ForegroundColor Green
        $ExchangeDeconnecte = $true
    }
    
    # Test SharePoint
    $SharePointDeconnecte = $false
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        Write-Host "  ATTENTION: SharePoint semble encore connecte!" -ForegroundColor Yellow
    }
    catch {
        Write-Host "  SharePoint: DECONNECTE" -ForegroundColor Green
        $SharePointDeconnecte = $true
    }
    
    # Test sessions restantes
    $SessionsRestantes = Get-PSSession
    if ($SessionsRestantes) {
        Write-Host "  ATTENTION: $($SessionsRestantes.Count) session(s) PowerShell encore active(s)" -ForegroundColor Yellow
        foreach ($Session in $SessionsRestantes) {
            Write-Host "    Session active: $($Session.Name) - $($Session.ComputerName) - Etat: $($Session.State)" -ForegroundColor Red
        }
    } else {
        Write-Host "  Sessions PowerShell: TOUTES FERMEES" -ForegroundColor Green
    }
    
    # Resultat final
    if ($ExchangeDeconnecte -and $SharePointDeconnecte -and -not $SessionsRestantes) {
        Write-Host "`nDeconnexion PowerShell REUSSIE!" -ForegroundColor Green
        return $true
    } else {
        Write-Host "`nDeconnexion PowerShell PARTIELLE - Certains services restent connectes" -ForegroundColor Yellow
        return $false
    }
}

try {
    # === ETAPE 1: VERIFICATION ET CONNEXION AUTOMATIQUE ===
    Write-Host "`n=== VERIFICATION ET CONNEXION MICROSOFT 365 ===" -ForegroundColor Cyan
    
    $ExchangeConnecte = $false
    $SharePointConnecte = $false
    $TenantActuel = ""
    $OrganisationActuelle = ""
    
    # Test connexion Exchange
    Write-Host "Test de la connexion Exchange..." -ForegroundColor Yellow
    try {
        $OrgConfig = Get-OrganizationConfig -ErrorAction Stop
        $TenantActuel = $OrgConfig.Name.Split('.')[0]
        $OrganisationActuelle = $OrgConfig.DisplayName
        $ExchangeConnecte = $true
        
        Write-Host "Exchange: CONNECTE" -ForegroundColor Green
        Write-Host "  Tenant: $TenantActuel" -ForegroundColor Cyan
        Write-Host "  Organisation: $OrganisationActuelle" -ForegroundColor Cyan
        
        # Verifier si c'est le bon tenant
        if ($TenantCible -and $TenantCible -ne $TenantActuel) {
            Write-Host "`nTenant cible ($TenantCible) different du tenant connecte ($TenantActuel)" -ForegroundColor Yellow
            $Response = Read-Host "Changer de tenant? (O/N)"
            
            if ($Response.ToUpper() -eq 'O') {
                Write-Host "Deconnexion du tenant actuel..." -ForegroundColor Red
                Disconnect-PowerShellSessions -Confirm $false
                $ExchangeConnecte = $false
                $SharePointConnecte = $false
                $TenantActuel = ""
            }
        }
    }
    catch {
        Write-Host "Exchange: NON CONNECTE" -ForegroundColor Red
    }
    
    # Si Exchange non connecte, etablir la connexion
    if (-not $ExchangeConnecte) {
        Write-Host "`n--- CONNEXION EXCHANGE REQUISE ---" -ForegroundColor Yellow
        
        # Determiner les informations de connexion
        if (-not $EmailAdmin) {
            if ($TenantCible) {
                Write-Host "Tenant cible: $TenantCible" -ForegroundColor Cyan
                
                # Proposer des formats d'email courants
                Write-Host "Formats d'email suggeres:" -ForegroundColor Gray
                Write-Host "  1. admin@$TenantCible.com" -ForegroundColor White
                Write-Host "  2. admin@$TenantCible.onmicrosoft.com" -ForegroundColor White
                Write-Host "  3. Saisie manuelle" -ForegroundColor White
                
                $EmailChoice = Read-Host "Choisir le format d'email (1/2/3)"
                switch ($EmailChoice) {
                    "1" { $EmailAdmin = "admin@$TenantCible.com" }
                    "2" { $EmailAdmin = "admin@$TenantCible.onmicrosoft.com" }
                    "3" { $EmailAdmin = Read-Host "Entrez l'email administrateur" }
                    default { $EmailAdmin = Read-Host "Entrez l'email administrateur" }
                }
            } else {
                Write-Host "Informations de connexion requises:" -ForegroundColor Yellow
                $EmailAdmin = Read-Host "Email administrateur (ex: admin@entreprise.com)"
                
                # Extraire le tenant depuis l'email si pas specifie
                if ($EmailAdmin -match '@(.+?)\.(com|onmicrosoft\.com)') {
                    $TenantCible = $matches[1]
                    Write-Host "Tenant detecte depuis l'email: $TenantCible" -ForegroundColor Cyan
                }
            }
        }
        
        # Validation de l'email
        if (-not $EmailAdmin -or $EmailAdmin -notmatch '^[^@]+@[^@]+\.[^@]+$') {
            Write-Error "Email administrateur invalide ou manquant"
            return
        }
        
        # Connexion Exchange
        Write-Host "`nConnexion a Exchange Online..." -ForegroundColor Yellow
        Write-Host "Email: $EmailAdmin" -ForegroundColor Cyan
        
        try {
            Connect-ExchangeOnline -UserPrincipalName $EmailAdmin -ShowProgress $true
            
            # Valider la connexion
            $OrgConfig = Get-OrganizationConfig
            $TenantActuel = $OrgConfig.Name.Split('.')[0]
            $OrganisationActuelle = $OrgConfig.DisplayName
            $ExchangeConnecte = $true
            
            Write-Host "Exchange connecte avec succes!" -ForegroundColor Green
            Write-Host "  Tenant: $TenantActuel" -ForegroundColor Cyan
            Write-Host "  Organisation: $OrganisationActuelle" -ForegroundColor Cyan
        }
        catch {
            Write-Error "Echec de la connexion Exchange: $($_.Exception.Message)"
            Write-Host "Verifiez:" -ForegroundColor Yellow
            Write-Host "  - L'email administrateur: $EmailAdmin" -ForegroundColor White
            Write-Host "  - Vos identifiants" -ForegroundColor White
            Write-Host "  - Votre connexion internet" -ForegroundColor White
            return
        }
    }
    
    # Test connexion SharePoint
    Write-Host "`nTest de la connexion SharePoint..." -ForegroundColor Yellow
    try {
        $TenantInfo = Get-SPOTenant -ErrorAction Stop
        $SharePointConnecte = $true
        Write-Host "SharePoint: CONNECTE" -ForegroundColor Green
    }
    catch {
        Write-Host "SharePoint: NON CONNECTE" -ForegroundColor Red
        
        # Proposer la connexion SharePoint
        if ($TenantActuel) {
            $SPOUrl = "https://$TenantActuel-admin.sharepoint.com"
            Write-Host "`n--- CONNEXION SHAREPOINT OPTIONNELLE ---" -ForegroundColor Yellow
            Write-Host "URL SharePoint detectee: $SPOUrl" -ForegroundColor Cyan
            Write-Host "SharePoint permet d'analyser OneDrive et les sites d'equipe" -ForegroundColor Gray
            
            $ConnectSPO = Read-Host "Connecter SharePoint maintenant? (O/N)"
            
            if ($ConnectSPO.ToUpper() -eq 'O') {
                Write-Host "Connexion a SharePoint Online..." -ForegroundColor Yellow
                try {
                    Connect-SPOService -Url $SPOUrl
                    $SharePointConnecte = $true
                    Write-Host "SharePoint connecte avec succes!" -ForegroundColor Green
                }
                catch {
                    Write-Warning "Echec connexion SharePoint: $($_.Exception.Message)"
                    Write-Host "L'analyse continuera sans SharePoint (Exchange uniquement)" -ForegroundColor Yellow
                    Write-Host "Pour vous connecter manuellement: Connect-SPOService -Url $SPOUrl" -ForegroundColor Cyan
                }
            } else {
                Write-Host "SharePoint ignore - Analyse Exchange uniquement" -ForegroundColor Gray
            }
        }
    }
    
    # Resume des connexions etablies
    Write-Host "`n--- RESUME DES CONNEXIONS ---" -ForegroundColor Cyan
    if ($ExchangeConnecte) {
        Write-Host "Exchange: CONNECTE ($TenantActuel - $OrganisationActuelle)" -ForegroundColor Green
    } else {
        Write-Host "Exchange: ECHEC - ARRET DU SCRIPT" -ForegroundColor Red
        return
    }
    
    if ($SharePointConnecte) {
        Write-Host "SharePoint: CONNECTE (OneDrive + Sites analyses)" -ForegroundColor Green
    } else {
        Write-Host "SharePoint: NON CONNECTE (Exchange uniquement)" -ForegroundColor Yellow
    }
    
    # Validation finale
    if (-not $ExchangeConnecte) {
        Write-Error "Connexion Exchange requise pour continuer l'analyse"
        return
    }
    
    # === ETAPE 2: CONFIGURATION DU RAPPORT ===
    if (-not $CheminRapport) {
        $DateStr = Get-Date -Format "yyyyMMdd_HHmm"
        $CheminRapport = "C:\Temp\M365_${TenantActuel}_${DateStr}.csv"
    }
    
    Write-Host "`n=== DEBUT DE L ANALYSE ===" -ForegroundColor Green
    Write-Host "Tenant: $TenantActuel" -ForegroundColor Cyan
    Write-Host "Organisation: $OrganisationActuelle" -ForegroundColor Cyan
    Write-Host "Rapport: $CheminRapport" -ForegroundColor Cyan
    
    $Report = @()
    
    # === ETAPE 3: ANALYSE EXCHANGE ===
    Write-Host "`nAnalyse Exchange..." -ForegroundColor Yellow
    
    $MailboxTypes = @(
        'UserMailbox',           # Boites utilisateurs
        'SharedMailbox',         # Boites partagees
        'RoomMailbox',          # Salles de reunion
        'EquipmentMailbox',     # Equipements
        'DiscoveryMailbox'      # Boites de decouverte
    )
    
    $AllMailboxes = @()
    foreach ($Type in $MailboxTypes) {
        try {
            Write-Host "  Recuperation des boites: $Type..." -ForegroundColor Gray
            $TypeMailboxes = Get-Mailbox -RecipientTypeDetails $Type -ResultSize Unlimited -ErrorAction SilentlyContinue
            if ($TypeMailboxes) {
                $AllMailboxes += $TypeMailboxes
                Write-Host "    Trouvees: $($TypeMailboxes.Count)" -ForegroundColor Green
            } else {
                Write-Host "    Trouvees: 0" -ForegroundColor Gray
            }
        }
        catch {
            Write-Warning "Erreur pour le type $Type : $($_.Exception.Message)"
        }
    }
    
    $MailboxCount = if ($AllMailboxes) { $AllMailboxes.Count } else { 0 }
    Write-Host "Total boites aux lettres: $MailboxCount" -ForegroundColor Cyan
    
    if ($MailboxCount -eq 0) {
        Write-Warning "Aucune boite aux lettres trouvee"
    } else {
        $Counter = 0
        foreach ($Mailbox in $AllMailboxes) {
            $Counter++
            if ($MailboxCount -gt 0) {
                Write-Progress -Activity "Analyse Exchange" -Status "Traitement: $($Mailbox.DisplayName)" -PercentComplete (($Counter / $MailboxCount) * 100)
            }
            
            try {
                $Stats = Get-MailboxStatistics -Identity $Mailbox.PrimarySmtpAddress -ErrorAction Stop
                $SizeGB = Convert-SizeToGB -SizeString $Stats.TotalItemSize
                
                $Report += [PSCustomObject]@{
                    Tenant = $TenantActuel
                    Service = "Exchange"
                    TypeBoite = $Mailbox.RecipientTypeDetails
                    Utilisateur = $Mailbox.DisplayName
                    Email = $Mailbox.PrimarySmtpAddress
                    TailleGB = $SizeGB
                    NombreElements = if ($Stats.ItemCount) { $Stats.ItemCount } else { 0 }
                    DernierAcces = $Stats.LastLogonTime
                }
            }
            catch {
                Write-Warning "Erreur pour $($Mailbox.DisplayName): $($_.Exception.Message)"
                $Report += [PSCustomObject]@{
                    Tenant = $TenantActuel
                    Service = "Exchange"
                    TypeBoite = $Mailbox.RecipientTypeDetails
                    Utilisateur = $Mailbox.DisplayName
                    Email = $Mailbox.PrimarySmtpAddress
                    TailleGB = 0
                    NombreElements = 0
                    DernierAcces = "Inaccessible"
                }
            }
        }
        Write-Progress -Activity "Analyse Exchange" -Completed
    }
    
    # === ETAPE 4: ANALYSE SHAREPOINT/ONEDRIVE ===
    if ($SharePointConnecte) {
        Write-Host "Analyse OneDrive..." -ForegroundColor Yellow
        try {
            $OneDrives = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop
            Write-Host "OneDrive trouves: $($OneDrives.Count)" -ForegroundColor Cyan
            
            $OneDriveCounter = 0
            foreach ($OneDrive in $OneDrives) {
                $OneDriveCounter++
                Write-Progress -Activity "Analyse OneDrive" -Status "Traitement: $($OneDrive.Title)" -PercentComplete (($OneDriveCounter / $OneDrives.Count) * 100)
                
                $SizeGB = [math]::Round($OneDrive.StorageUsageCurrent / 1024, 4)
                
                $Report += [PSCustomObject]@{
                    Tenant = $TenantActuel
                    Service = "OneDrive"
                    TypeBoite = "OneDrive"
                    Utilisateur = $OneDrive.Title
                    Email = $OneDrive.Owner
                    TailleGB = $SizeGB
                    NombreElements = "N/A"
                    DernierAcces = $OneDrive.LastContentModifiedDate
                }
            }
            Write-Progress -Activity "Analyse OneDrive" -Completed
        }
        catch {
            Write-Warning "Erreur OneDrive: $($_.Exception.Message)"
        }
        
        Write-Host "Analyse SharePoint..." -ForegroundColor Yellow
        try {
            $SharePointSites = Get-SPOSite -Limit All | Where-Object {$_.Url -notlike "*-my.sharepoint.com/personal/*"}
            Write-Host "Sites SharePoint trouves: $($SharePointSites.Count)" -ForegroundColor Cyan
            
            $SharePointCounter = 0
            foreach ($Site in $SharePointSites) {
                $SharePointCounter++
                Write-Progress -Activity "Analyse SharePoint" -Status "Traitement: $($Site.Title)" -PercentComplete (($SharePointCounter / $SharePointSites.Count) * 100)
                
                $SizeGB = [math]::Round($Site.StorageUsageCurrent / 1024, 4)
                
                $Report += [PSCustomObject]@{
                    Tenant = $TenantActuel
                    Service = "SharePoint"
                    TypeBoite = "Site SharePoint"
                    Utilisateur = $Site.Title
                    Email = $Site.Url
                    TailleGB = $SizeGB
                    NombreElements = "N/A"
                    DernierAcces = $Site.LastContentModifiedDate
                }
            }
            Write-Progress -Activity "Analyse SharePoint" -Completed
        }
        catch {
            Write-Warning "Erreur SharePoint: $($_.Exception.Message)"
        }
    } else {
        Write-Host "SharePoint non connecte - OneDrive et Sites ignores" -ForegroundColor Yellow
        if ($TenantActuel) {
            Write-Host "Pour analyser SharePoint: Connect-SPOService -Url https://$TenantActuel-admin.sharepoint.com" -ForegroundColor Cyan
        }
    }
    
    # === ETAPE 5: GENERATION DU RAPPORT ===
    Write-Host "`n=== GENERATION DU RAPPORT ===" -ForegroundColor Magenta
    
    # Creer le repertoire
    $RepertoireRapport = Split-Path $CheminRapport -Parent
    if (!(Test-Path $RepertoireRapport)) {
        New-Item -ItemType Directory -Path $RepertoireRapport -Force | Out-Null
        Write-Host "Repertoire cree: $RepertoireRapport" -ForegroundColor Gray
    }
    
    # Exporter
    if ($Report.Count -gt 0) {
        $Report | Export-Csv -Path $CheminRapport -NoTypeInformation -Encoding UTF8
        Write-Host "Donnees exportees: $($Report.Count) elements" -ForegroundColor Green
    } else {
        Write-Warning "Aucune donnee a exporter"
    }
    
    # === ETAPE 6: AFFICHAGE DU RESUME ===
    Write-Host "`n=== RESUME DE L ANALYSE ===" -ForegroundColor Green
    Write-Host "TENANT: $TenantActuel" -ForegroundColor Magenta
    Write-Host "ORGANISATION: $OrganisationActuelle" -ForegroundColor Magenta
    
    if ($Report.Count -eq 0) {
        Write-Warning "Aucune donnee collectee"
        return
    }
    
    $Resume = $Report | Group-Object Service | ForEach-Object {
        [PSCustomObject]@{
            Service = $_.Name
            Elements = $_.Count
            TailleTotal_GB = [math]::Round(($_.Group | Measure-Object TailleGB -Sum).Sum, 2)
            TailleMoyenne_GB = [math]::Round(($_.Group | Measure-Object TailleGB -Average).Average, 2)
        }
    }
    
    $Resume | Format-Table -AutoSize
    
    # Detail Exchange par type
    $ExchangeData = $Report | Where-Object {$_.Service -eq "Exchange"}
    if ($ExchangeData) {
        Write-Host "DETAIL EXCHANGE PAR TYPE:" -ForegroundColor Green
        $ResumeExchange = $ExchangeData | Group-Object TypeBoite | ForEach-Object {
            [PSCustomObject]@{
                TypeBoite = $_.Name
                Nombre = $_.Count
                TailleTotal_GB = [math]::Round(($_.Group | Measure-Object TailleGB -Sum).Sum, 2)
                TailleMoyenne_GB = [math]::Round(($_.Group | Measure-Object TailleGB -Average).Average, 2)
            }
        }
        $ResumeExchange | Format-Table -AutoSize
    }
    
    # Totaux et couts
    $TotalGeneral = [math]::Round(($Report | Measure-Object TailleGB -Sum).Sum, 2)
    $CoutEstime = [math]::Round($TotalGeneral * 0.15, 2)
    
    Write-Host "TAILLE TOTALE: $TotalGeneral GB" -ForegroundColor Cyan
    Write-Host "COUT ESTIME SAUVEGARDE (0,15$/GB/mois): $CoutEstime$ USD/mois" -ForegroundColor Yellow
    Write-Host "RAPPORT SAUVEGARDE: $CheminRapport" -ForegroundColor Green
    
    # Top 10 consommateurs
    Write-Host "`nTOP 10 PLUS GROS CONSOMMATEURS:" -ForegroundColor Green
    $Report | Sort-Object TailleGB -Descending | Select-Object -First 10 Service, TypeBoite, Utilisateur, TailleGB | Format-Table -AutoSize
    
} catch {
    Write-Error "Erreur lors de l execution: $($_.Exception.Message)"
    Write-Host "Ligne d erreur: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
} finally {
    # === ETAPE 7: PROPOSITION DE DECONNEXION ===
    Write-Host "`n=== GESTION DES CONNEXIONS ===" -ForegroundColor Cyan
    
    # Afficher l'etat actuel
    Write-Host "Etat actuel des connexions PowerShell:" -ForegroundColor White
    
    try {
        $CurrentOrg = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "  Exchange: CONNECTE ($($CurrentOrg.Name))" -ForegroundColor Green
    }
    catch {
        Write-Host "  Exchange: DECONNECTE" -ForegroundColor Gray
    }
    
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        Write-Host "  SharePoint: CONNECTE" -ForegroundColor Green
    }
    catch {
        Write-Host "  SharePoint: DECONNECTE" -ForegroundColor Gray
    }
    
    Write-Host "`nVoulez-vous deconnecter les sessions PowerShell?" -ForegroundColor Yellow
    Write-Host "(Les applications locales comme OneDrive restent connectees)" -ForegroundColor Gray
    $DeconnexionSouhaitee = Read-Host "Deconnecter PowerShell? (O/N)"
    
    if ($DeconnexionSouhaitee.ToUpper() -eq 'O') {
        $Deconnecte = Disconnect-PowerShellSessions -Confirm $false
        
        # Verification supplementaire apres deconnexion
        Start-Sleep -Seconds 2
        
        Write-Host "`nVerification finale post-deconnexion:" -ForegroundColor Cyan
        
        $StillConnected = @()
        
        # Re-test Exchange
        try {
            $TestExchange = Get-OrganizationConfig -ErrorAction Stop
            $StillConnected += "Exchange ($($TestExchange.Name))"
            Write-Host "  Exchange: ENCORE CONNECTE" -ForegroundColor Red
        }
        catch {
            Write-Host "  Exchange: DECONNECTE" -ForegroundColor Green
        }
        
        # Re-test SharePoint
        try {
            $null = Get-SPOTenant -ErrorAction Stop
            $StillConnected += "SharePoint"
            Write-Host "  SharePoint: ENCORE CONNECTE" -ForegroundColor Yellow
        }
        catch {
            Write-Host "  SharePoint: DECONNECTE" -ForegroundColor Green
        }
        
        # Sessions restantes
        $RemainingSessions = Get-PSSession
        if ($RemainingSessions) {
            $StillConnected += "$($RemainingSessions.Count) session(s) PowerShell"
            Write-Host "  Sessions: $($RemainingSessions.Count) ENCORE ACTIVES" -ForegroundColor Yellow
        } else {
            Write-Host "  Sessions: TOUTES FERMEES" -ForegroundColor Green
        }
        
        if ($StillConnected.Count -eq 0) {
            Write-Host "`nDeconnexion COMPLETE - Toutes les sessions PowerShell sont fermees" -ForegroundColor Green
        } else {
            Write-Host "`nDeconnexion PARTIELLE - Services encore connectes:" -ForegroundColor Yellow
            foreach ($Service in $StillConnected) {
                Write-Host "  - $Service" -ForegroundColor Red
            }
            
            Write-Host "`nPour forcer la deconnexion complete:" -ForegroundColor Cyan
            Write-Host "  Get-PSSession | Remove-PSSession -Force" -ForegroundColor White
            Write-Host "  Disconnect-ExchangeOnline -Confirm:`$false" -ForegroundColor White
        }
    } else {
        Write-Host "Sessions PowerShell conservees" -ForegroundColor Cyan
        Write-Host "Pour vous deconnecter manuellement plus tard:" -ForegroundColor Gray
        Write-Host "  Disconnect-ExchangeOnline -Confirm:`$false" -ForegroundColor White
        Write-Host "  Disconnect-SPOService" -ForegroundColor White
        Write-Host "  Get-PSSession | Remove-PSSession" -ForegroundColor White
    }
}

Write-Host "`n======================================================" -ForegroundColor Magenta
Write-Host "           ANALYSE TERMINEE AVEC SUCCES" -ForegroundColor Magenta
Write-Host "======================================================" -ForegroundColor Magenta

# Exemples d utilisation
<#
EXEMPLES D UTILISATION:

# Execution simple (detecte connexions existantes)
.\Analyse-M365-Unifie.ps1

# Analyser un tenant specifique
.\Analyse-M365-Unifie.ps1 -TenantCible "monautretenant" -EmailAdmin "admin@monautretenant.com"

# Avec chemin de rapport personnalise
.\Analyse-M365-Unifie.ps1 -CheminRapport "D:\Rapports\MonAnalyse.csv"

# Combinaison complete
.\Analyse-M365-Unifie.ps1 -TenantCible "client1" -EmailAdmin "admin@client1.com" -CheminRapport "C:\Rapports\Client1.csv"
#>