# ============================================================
# GUI - Gestionnaire Active Directory
# Auteur : Gaël RAUTUREAU
# Description : Interface graphique de provisioning AD
#               Création OUs, groupes et utilisateurs depuis CSV
# ============================================================

# ------------------------------------------------------------
# FONCTION : Génération de mot de passe complexe
# Respecte la politique de complexité Active Directory
# ------------------------------------------------------------
function New-MotDePasse {
    $longueur = 12
    $majuscules = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $minuscules = "abcdefghijklmnopqrstuvwxyz"
    $chiffres   = "0123456789"
    $speciaux   = "!@#$%^&*"
    $tousCaracteres = $majuscules + $minuscules + $chiffres + $speciaux
    $mdp  = $majuscules[(Get-Random -Maximum $majuscules.Length)]
    $mdp += $minuscules[(Get-Random -Maximum $minuscules.Length)]
    $mdp += $chiffres[(Get-Random -Maximum $chiffres.Length)]
    $mdp += $speciaux[(Get-Random -Maximum $speciaux.Length)]
    for ($i = 4; $i -lt $longueur; $i++) {
        $mdp += $tousCaracteres[(Get-Random -Maximum $tousCaracteres.Length)]
    }
    # Mélange les caractères pour éviter un pattern prévisible
    $mdp = ($mdp.ToCharArray() | Get-Random -Count $mdp.Length) -join ""
    return $mdp
}

# ------------------------------------------------------------
# CHARGEMENT DES ASSEMBLIES WINDOWS FORMS
# ------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
# CONSTRUCTION DU FORMULAIRE PRINCIPAL
# ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Gestionnaire AD - TechCorp"
$form.Size = New-Object System.Drawing.Size(600, 450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Text = "Gestionnaire de comptes Active Directory"
$label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(560, 30)
$label.TextAlign = "MiddleCenter"
$form.Controls.Add($label)

# --- Titre ---
$labelCSV = New-Object System.Windows.Forms.Label
$labelCSV.Text = "Aucun fichier sélectionné"
$labelCSV.Location = New-Object System.Drawing.Point(10, 90)
$labelCSV.Size = New-Object System.Drawing.Size(460, 25)
$form.Controls.Add($labelCSV)

$btnCSV = New-Object System.Windows.Forms.Button
$btnCSV.Text = "Choisir le fichier CSV"
$btnCSV.Location = New-Object System.Drawing.Point(10, 55)
$btnCSV.Size = New-Object System.Drawing.Size(200, 30)
$form.Controls.Add($btnCSV)

# --- Bouton export rapport (désactivé jusqu'à la fin de la création) ---
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exporter le rapport"
$btnExport.Location = New-Object System.Drawing.Point(220, 130)
$btnExport.Size = New-Object System.Drawing.Size(200, 30)
$btnExport.Enabled = $false

$form.Controls.Add($btnExport)

# --- Nom de l'entreprise (modifiable) ---
$labelEntreprise = New-Object System.Windows.Forms.Label
$labelEntreprise.Text = "Nom de l'entreprise :"
$labelEntreprise.Location = New-Object System.Drawing.Point(10, 170)
$labelEntreprise.Size = New-Object System.Drawing.Size(150, 25)
$form.Controls.Add($labelEntreprise)

$textEntreprise = New-Object System.Windows.Forms.TextBox
$textEntreprise.Text = "TechCorp"
$textEntreprise.Location = New-Object System.Drawing.Point(165, 170)
$textEntreprise.Size = New-Object System.Drawing.Size(200, 25)
$form.Controls.Add($textEntreprise)

# ------------------------------------------------------------
# EVENEMENTS DES BOUTONS
# ------------------------------------------------------------

# --- Ouverture du fichier CSV ---
$btnCSV.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    $dialog.Title = "Choisir le fichier CSV"
    if ($dialog.ShowDialog() -eq "OK") {
        $script:cheminCSV = $dialog.FileName
        $labelCSV.Text = "Fichier : $($dialog.FileName)"
    }
})

# --- Bouton lancer la création ---
$btnLancer = New-Object System.Windows.Forms.Button
$btnLancer.Text = "Lancer la création"
$btnLancer.Location = New-Object System.Drawing.Point(10, 130)
$btnLancer.Size = New-Object System.Drawing.Size(200, 30)
$form.Controls.Add($btnLancer)

# --- Textbox log (fond noir, texte vert) ---
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 210)
$textBox.Size = New-Object System.Drawing.Size(560, 160)
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.ReadOnly = $true
$textBox.BackColor = "Black"
$textBox.ForeColor = "Lime"
$form.Controls.Add($textBox)

# --- Export du rapport CSV ---
$btnExport.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    $dialog.FileName = "rapport_$(Get-Date -Format 'dd_MM_yyyy_HH-mm').csv"
    if ($dialog.ShowDialog() -eq "OK") {
        $script:rapport_mdp | Out-File -FilePath $dialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Rapport exporté avec succès !", "Export")
    }
})

# ------------------------------------------------------------
# LOGIQUE PRINCIPALE - CREATION AD
# ------------------------------------------------------------

# Vérification du fichier CSV
$btnLancer.Add_Click({
    if (-not $script:cheminCSV) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez choisir un fichier CSV d'abord !", "Erreur")
        return
       
    }

    $textBox.Clear()
    $textBox.AppendText("Démarrage du script...`r`n")
    [System.Windows.Forms.Application]::DoEvents()

    # Chargement du CSV et détection automatique du domaine
    $utilisateurs = Import-Csv -Path $script:cheminCSV
    $domaine      = (Get-ADDomain).DNSRoot
    $dcPath       = ($domaine.Split(".") | ForEach-Object { "DC=$_" }) -join ","
    $nomEntreprise = $textEntreprise.Text
    $ouRacine = "OU=$nomEntreprise,$dcPath"

    $textBox.AppendText("Fichier CSV chargé : $($utilisateurs.Count) utilisateurs trouvés.`r`n")
    [System.Windows.Forms.Application]::DoEvents()

    # Création des groupes
    $textBox.AppendText("`r`nCréation des groupes...`r`n")
    [System.Windows.Forms.Application]::DoEvents()
   
    # --- Création des groupes de sécurité ---
    foreach ($groupe in @("GRP-Informatique", "GRP-Administration", "GRP-Comptabilite", "GRP-PDG", "GRP-Direction")) {
        if (Get-ADGroup -Filter "Name -eq '$groupe'" -ErrorAction SilentlyContinue) {
            $textBox.AppendText("  [EXISTE] $groupe`r`n")
        } else {
            New-ADGroup -Name $groupe -GroupScope Global -GroupCategory Security
            $textBox.AppendText("  [CREE] $groupe`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

# Création des OUs
    $textBox.AppendText("`r`nCréation des OUs...`r`n")
    [System.Windows.Forms.Application]::DoEvents()

     # --- Création des OUs ---
    if (-not (Get-ADOrganizationalUnit -Filter "Name -eq '$nomEntreprise'" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $nomEntreprise -Path $dcPath
        $textBox.AppendText("  [CRÉÉ] OU $nomEntreprise`r`n")
    } else {
        $textBox.AppendText("  [EXISTE] OU $nomEntreprise`r`n")
    }
    foreach ($ou in @("Informatique", "Administration", "Comptabilite", "Direction")) {
        if (-not (Get-ADOrganizationalUnit -Filter "Name -eq '$ou'" -SearchBase $ouRacine -ErrorAction SilentlyContinue)) {
            New-ADOrganizationalUnit -Name $ou -Path $ouRacine
            $textBox.AppendText("  [CRÉÉ] OU $ou`r`n")
        } else {
            $textBox.AppendText("  [EXISTE] OU $ou`r`n")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Création des utilisateurs
    $textBox.AppendText("`r`nCréation des utilisateurs...`r`n")
    [System.Windows.Forms.Application]::DoEvents()
    $script:rapport_mdp = @()

    # --- Création des utilisateurs et affectation aux groupes ---
    foreach ($user in $utilisateurs) {
        $login = "$($user.Prenom.ToLower()).$($user.Nom.ToLower())"
        $email = "$login@$($nomEntreprise.ToLower()).com"
        if (Get-ADUser -Filter "SamAccountName -eq '$login'" -ErrorAction SilentlyContinue) {
            $textBox.AppendText("  [EXISTE] $login`r`n")
            
        } else {
    $mdp = New-MotDePasse
    $ou = "OU=$($user.Departement),OU=$nomEntreprise,$dcPath"
    try {
        New-ADUser `
            -Name "$($user.Prenom) $($user.Nom)" `
            -GivenName $user.Prenom `
            -Surname $user.Nom `
            -SamAccountName $login `
            -UserPrincipalName "$login@$domaine" `
            -Department $user.Departement `
            -Title $user.Titre `
            -Company $user.Societe `
            -City $user.Ville `
            -Path $ou `
            -AccountPassword (ConvertTo-SecureString $mdp -AsPlainText -Force) `
            -EmailAddress $email `
            -Enabled $true `
            -ChangePasswordAtLogon $true
        $script:rapport_mdp += "$login | $email | MDP : $mdp"
        $textBox.AppendText("  [CRÉÉ] $login | MDP : $mdp`r`n")
        $groupe = "GRP-$($user.Departement)"
        Add-ADGroupMember -Identity $groupe -Members $login
        $textBox.AppendText("  [GROUPE] $login → $groupe`r`n")
    } catch {
        $textBox.AppendText("  [ERREUR] $login : $($_.Exception.Message)`r`n")
    }
}
        [System.Windows.Forms.Application]::DoEvents()
    }
        $textBox.AppendText("`r`nTerminé !`r`n")
        
        # Activation du bouton export une fois la création terminée
        $btnExport.Enabled = $true

})

# ------------------------------------------------------------
# LANCEMENT DU FORMULAIRE
# ------------------------------------------------------------
$form.ShowDialog()