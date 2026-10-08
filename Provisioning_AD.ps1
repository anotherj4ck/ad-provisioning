# ============================================================
#  PROVISIONING AD - GUI Windows Forms v5
#  CSV unique : OUs / Groupes / Utilisateurs
#  Auteur    : anotherj4ck
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# ============================================================
# CONSTANTES & CHEMINS
# ============================================================
$script:SEP        = " -> "
$script:configPath = "$PSScriptRoot\config_ad.csv"
$script:logPath    = "C:\Logs\provisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ============================================================
# DONNEES EN MEMOIRE
# ============================================================
$script:listeOUs     = [System.Collections.ArrayList]@()
$script:listeGroupes = [System.Collections.ArrayList]@()

# ============================================================
# FONCTIONS
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    if ($null -ne $txtLog -and $txtLog.IsHandleCreated) {
        $txtLog.AppendText("$line`r`n")
        $txtLog.ScrollToCaret()
    }
    $dir = Split-Path $script:logPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -Path $script:logPath -Value $line
}

function Test-NomValide {
    param([string]$Nom)
    if ($Nom -match '[,\\#+<>;"=]') {
        [System.Windows.Forms.MessageBox]::Show(
            "Nom invalide : caracteres interdits detectes.`nEvitez : , \ # + < > ; `" =",
            "Validation", "OK", "Warning") | Out-Null
        return $false
    }
    if ($Nom.Length -gt 64) {
        [System.Windows.Forms.MessageBox]::Show(
            "Nom trop long (max 64 caracteres).", "Validation", "OK", "Warning") | Out-Null
        return $false
    }
    return $true
}

function Update-ComboboxOUs {
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    foreach ($o in $script:listeOUs) {
        $cmbOUGroupe.Items.Add($o) | Out-Null
        $cmbOUDefaut.Items.Add($o) | Out-Null
    }
}

function Update-ComboboxGroupes {
    $cmbGroupeDefaut.Items.Clear()
    foreach ($entry in $script:listeGroupes) {
        $idx = $entry.IndexOf($script:SEP)
        $nom = $entry.Substring(0, $idx).Trim()
        $cmbGroupeDefaut.Items.Add($nom) | Out-Null
    }
}

function Sync-GridColumns {
    $colOU     = $grid.Columns["OU"]
    $colGroupe = $grid.Columns["Groupe"]
    $colOU.Items.Clear()
    $colGroupe.Items.Clear()
    foreach ($o in $script:listeOUs) { $colOU.Items.Add($o) | Out-Null }
    foreach ($entry in $script:listeGroupes) {
        $idx = $entry.IndexOf($script:SEP)
        $nom = $entry.Substring(0, $idx).Trim()
        $colGroupe.Items.Add($nom) | Out-Null
    }
}

function Nettoyer-OrphansOU {
    param([string]$OUSupprimee)
    foreach ($row in $grid.Rows) {
        if ($row.Cells["OU"].Value -eq $OUSupprimee) { $row.Cells["OU"].Value = $null }
    }
}

function Nettoyer-OrphansGroupe {
    param([string]$GroupeSupprime)
    foreach ($row in $grid.Rows) {
        if ($row.Cells["Groupe"].Value -eq $GroupeSupprime) { $row.Cells["Groupe"].Value = $null }
    }
}

function Export-Config {
    $lignes = [System.Collections.ArrayList]@()
    $lignes.Add("Type;Nom;OU;Login;Prenom;Nom2;MotDePasse;Groupe") | Out-Null

    foreach ($ou in $script:listeOUs) { $lignes.Add("OU;$ou;;;;;;;") | Out-Null }

    foreach ($entry in $script:listeGroupes) {
        $idx = $entry.IndexOf($script:SEP)
        $nom = $entry.Substring(0, $idx).Trim()
        $ou  = $entry.Substring($idx + $script:SEP.Length).Trim()
        $lignes.Add("GROUPE;$nom;$ou;;;;;;") | Out-Null
    }

    foreach ($row in $grid.Rows) {
        $prenom = $row.Cells["Prenom"].Value
        $nom    = $row.Cells["Nom"].Value
        $login  = $row.Cells["Login"].Value
        $mdp    = $row.Cells["MotDePasse"].Value
        $ou     = $row.Cells["OU"].Value
        $grp    = $row.Cells["Groupe"].Value
        if ($login) { $lignes.Add("USER;;$ou;$login;$prenom;$nom;$mdp;$grp") | Out-Null }
    }

    $lignes | Set-Content -Path $script:configPath -Encoding UTF8
}

function Import-Config {
    param([string]$Chemin)
    if (-not (Test-Path $Chemin)) { return }

    $data = Import-Csv -Path $Chemin -Delimiter ";"

    $script:listeOUs.Clear()
    $script:listeGroupes.Clear()
    $lstOU.Items.Clear()
    $lstGroupe.Items.Clear()
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    $cmbGroupeDefaut.Items.Clear()
    $grid.Rows.Clear()

    foreach ($row in $data) {
        if ($row.Type -eq "OU") {
            $val = $row.Nom.Trim()
            if ($val -ne "" -and $script:listeOUs -notcontains $val) {
                $script:listeOUs.Add($val) | Out-Null
                $lstOU.Items.Add($val) | Out-Null
                $cmbOUGroupe.Items.Add($val) | Out-Null
                $cmbOUDefaut.Items.Add($val) | Out-Null
            }
        }
    }

    foreach ($row in $data) {
        if ($row.Type -eq "GROUPE") {
            $nom   = $row.Nom.Trim()
            $ou    = $row.OU.Trim()
            $entry = "$nom$($script:SEP)$ou"
            if ($nom -ne "" -and $ou -ne "" -and $script:listeGroupes -notcontains $entry) {
                $script:listeGroupes.Add($entry) | Out-Null
                $lstGroupe.Items.Add($entry) | Out-Null
                $cmbGroupeDefaut.Items.Add($nom) | Out-Null
            }
        }
    }

    Sync-GridColumns

    foreach ($row in $data) {
        if ($row.Type -eq "USER") {
            $prenom = $row.Prenom.Trim()
            $nom2   = $row.Nom2.Trim()
            $login  = $row.Login.Trim()
            $mdp    = $row.MotDePasse.Trim()
            $ou     = $row.OU.Trim()
            $grp    = $row.Groupe.Trim()
            if ($login -ne "") {
                $grid.Rows.Add($prenom, $nom2, $login, $mdp, $ou, $grp) | Out-Null
            }
        }
    }
}

function Import-DepuisAD {
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Cette action remplace les donnees de l'interface par celles lues depuis l'AD.`n`nContinuer ?",
        "Lecture depuis l'AD", "YesNo", "Question")
    if ($confirm -ne "Yes") { return }

    try {
        $domaine = Get-ADDomain -ErrorAction Stop
        $dcRoot  = $domaine.DistinguishedName
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Impossible de contacter l'AD :`n$_", "Erreur AD", "OK", "Error")
        return
    }

    $script:listeOUs.Clear()
    $script:listeGroupes.Clear()
    $lstOU.Items.Clear()
    $lstGroupe.Items.Clear()
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    $cmbGroupeDefaut.Items.Clear()
    $grid.Rows.Clear()

    # OUs de niveau 1 sous la racine du domaine
    $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $dcRoot -SearchScope OneLevel -ErrorAction SilentlyContinue
    foreach ($ou in $ous) {
        $nom = $ou.Name
        if ($script:listeOUs -notcontains $nom) {
            $script:listeOUs.Add($nom) | Out-Null
            $lstOU.Items.Add($nom) | Out-Null
            $cmbOUGroupe.Items.Add($nom) | Out-Null
            $cmbOUDefaut.Items.Add($nom) | Out-Null
        }
    }

    # Groupes dans les OUs importees
    $groupes = Get-ADGroup -Filter * -SearchBase $dcRoot -Properties DistinguishedName -ErrorAction SilentlyContinue
    foreach ($grp in $groupes) {
        $nom     = $grp.Name
        $dn      = $grp.DistinguishedName
        $ouMatch = [regex]::Match($dn, '(?<=,OU=)[^,]+')
        $ouNom   = if ($ouMatch.Success) { $ouMatch.Value } else { "" }
        if ($ouNom -ne "" -and $script:listeOUs -contains $ouNom) {
            $entry = "$nom$($script:SEP)$ouNom"
            if ($script:listeGroupes -notcontains $entry) {
                $script:listeGroupes.Add($entry) | Out-Null
                $lstGroupe.Items.Add($entry) | Out-Null
                $cmbGroupeDefaut.Items.Add($nom) | Out-Null
            }
        }
    }

    Sync-GridColumns

    # Utilisateurs dans les OUs connues (mot de passe laisse vide — illisible depuis l'AD)
    $users = Get-ADUser -Filter * -SearchBase $dcRoot `
        -Properties GivenName, Surname, DistinguishedName, MemberOf `
        -ErrorAction SilentlyContinue
    foreach ($user in $users) {
        $login   = $user.SamAccountName
        $dn      = $user.DistinguishedName
        $ouMatch = [regex]::Match($dn, '(?<=,OU=)[^,]+')
        $ouNom   = if ($ouMatch.Success) { $ouMatch.Value } else { "" }
        if ($ouNom -eq "" -or $script:listeOUs -notcontains $ouNom) { continue }

        $grpNom = ""
        foreach ($memberDN in $user.MemberOf) {
            $grpMatch = [regex]::Match($memberDN, 'CN=([^,]+)')
            if ($grpMatch.Success) {
                $candidate = $grpMatch.Groups[1].Value
                foreach ($entry in $script:listeGroupes) {
                    $idx = $entry.IndexOf($script:SEP)
                    if ($entry.Substring(0, $idx).Trim() -eq $candidate) {
                        $grpNom = $candidate; break
                    }
                }
                if ($grpNom -ne "") { break }
            }
        }

        $grid.Rows.Add($user.GivenName, $user.Surname, $login, "", $ouNom, $grpNom) | Out-Null
    }

    [System.Windows.Forms.MessageBox]::Show(
        "Lecture AD terminee.`n$($script:listeOUs.Count) OU(s)   |   $($script:listeGroupes.Count) groupe(s)   |   $($grid.Rows.Count) utilisateur(s)`n`nNote : les mots de passe ne sont pas recuperables depuis l'AD.",
        "Lecture AD", "OK", "Information")
}

# ============================================================
# FENETRE PRINCIPALE
# ============================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Provisioning AD v5"
$form.Size            = New-Object System.Drawing.Size(820, 860)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ============================================================
# TABCONTROL
# ============================================================
$tabControl           = New-Object System.Windows.Forms.TabControl
$tabControl.Location  = New-Object System.Drawing.Point(10, 10)
$tabControl.Size      = New-Object System.Drawing.Size(780, 520)
$tabControl.Font      = New-Object System.Drawing.Font("Segoe UI", 9)

$tabOU                = New-Object System.Windows.Forms.TabPage; $tabOU.Text     = "OUs"
$tabGroupe            = New-Object System.Windows.Forms.TabPage; $tabGroupe.Text = "Groupes"
$tabUser              = New-Object System.Windows.Forms.TabPage; $tabUser.Text   = "Utilisateurs"

$tabControl.TabPages.AddRange(@($tabOU, $tabGroupe, $tabUser))
$form.Controls.Add($tabControl)

# ============================================================
# ONGLET OUs
# ============================================================

$lblOU          = New-Object System.Windows.Forms.Label
$lblOU.Text     = "Nom de l'OU :"
$lblOU.Location = New-Object System.Drawing.Point(15, 20)
$lblOU.Size     = New-Object System.Drawing.Size(120, 22)
$tabOU.Controls.Add($lblOU)

$txtOU          = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(140, 18)
$txtOU.Size     = New-Object System.Drawing.Size(200, 22)
$tabOU.Controls.Add($txtOU)

$btnAjouterOU          = New-Object System.Windows.Forms.Button
$btnAjouterOU.Text     = "Ajouter"
$btnAjouterOU.Location = New-Object System.Drawing.Point(355, 16)
$btnAjouterOU.Size     = New-Object System.Drawing.Size(80, 26)
$tabOU.Controls.Add($btnAjouterOU)

$btnSupprimerOU          = New-Object System.Windows.Forms.Button
$btnSupprimerOU.Text     = "Supprimer"
$btnSupprimerOU.Location = New-Object System.Drawing.Point(445, 16)
$btnSupprimerOU.Size     = New-Object System.Drawing.Size(80, 26)
$tabOU.Controls.Add($btnSupprimerOU)

$lstOU          = New-Object System.Windows.Forms.ListBox
$lstOU.Location = New-Object System.Drawing.Point(15, 55)
$lstOU.Size     = New-Object System.Drawing.Size(510, 380)
$lstOU.Font     = New-Object System.Drawing.Font("Consolas", 9)
$tabOU.Controls.Add($lstOU)

$btnAjouterOU.Add_Click({
    $val = $txtOU.Text.Trim()
    if ($val -ne "" -and (Test-NomValide $val) -and $script:listeOUs -notcontains $val) {
        $script:listeOUs.Add($val) | Out-Null
        $lstOU.Items.Add($val) | Out-Null
        Update-ComboboxOUs
        Sync-GridColumns
        $txtOU.Clear()
    }
})

$btnSupprimerOU.Add_Click({
    if ($lstOU.SelectedItem) {
        $val = $lstOU.SelectedItem
        $script:listeOUs.Remove($val)
        $lstOU.Items.Remove($val)
        Update-ComboboxOUs
        Sync-GridColumns
        Nettoyer-OrphansOU -OUSupprimee $val
    }
})

# ============================================================
# ONGLET GROUPES
# ============================================================

$lblGroupe          = New-Object System.Windows.Forms.Label
$lblGroupe.Text     = "Nom du groupe :"
$lblGroupe.Location = New-Object System.Drawing.Point(15, 20)
$lblGroupe.Size     = New-Object System.Drawing.Size(120, 22)
$tabGroupe.Controls.Add($lblGroupe)

$txtGroupe          = New-Object System.Windows.Forms.TextBox
$txtGroupe.Location = New-Object System.Drawing.Point(140, 18)
$txtGroupe.Size     = New-Object System.Drawing.Size(160, 22)
$tabGroupe.Controls.Add($txtGroupe)

$lblOUGroupe          = New-Object System.Windows.Forms.Label
$lblOUGroupe.Text     = "OU cible :"
$lblOUGroupe.Location = New-Object System.Drawing.Point(315, 20)
$lblOUGroupe.Size     = New-Object System.Drawing.Size(65, 22)
$tabGroupe.Controls.Add($lblOUGroupe)

$cmbOUGroupe               = New-Object System.Windows.Forms.ComboBox
$cmbOUGroupe.Location      = New-Object System.Drawing.Point(385, 17)
$cmbOUGroupe.Size          = New-Object System.Drawing.Size(140, 22)
$cmbOUGroupe.DropDownStyle = "DropDownList"
$tabGroupe.Controls.Add($cmbOUGroupe)

$btnAjouterGroupe          = New-Object System.Windows.Forms.Button
$btnAjouterGroupe.Text     = "Ajouter"
$btnAjouterGroupe.Location = New-Object System.Drawing.Point(540, 15)
$btnAjouterGroupe.Size     = New-Object System.Drawing.Size(75, 26)
$tabGroupe.Controls.Add($btnAjouterGroupe)

$btnSupprimerGroupe          = New-Object System.Windows.Forms.Button
$btnSupprimerGroupe.Text     = "Supprimer"
$btnSupprimerGroupe.Location = New-Object System.Drawing.Point(625, 15)
$btnSupprimerGroupe.Size     = New-Object System.Drawing.Size(80, 26)
$tabGroupe.Controls.Add($btnSupprimerGroupe)

$lstGroupe          = New-Object System.Windows.Forms.ListBox
$lstGroupe.Location = New-Object System.Drawing.Point(15, 55)
$lstGroupe.Size     = New-Object System.Drawing.Size(720, 380)
$lstGroupe.Font     = New-Object System.Drawing.Font("Consolas", 9)
$tabGroupe.Controls.Add($lstGroupe)

$btnAjouterGroupe.Add_Click({
    $nom = $txtGroupe.Text.Trim()
    $ou  = $cmbOUGroupe.SelectedItem
    if ($nom -ne "" -and $ou -and (Test-NomValide $nom)) {
        $entry = "$nom$($script:SEP)$ou"
        if ($script:listeGroupes -notcontains $entry) {
            $script:listeGroupes.Add($entry) | Out-Null
            $lstGroupe.Items.Add($entry) | Out-Null
            $cmbGroupeDefaut.Items.Add($nom) | Out-Null
            Sync-GridColumns
            $txtGroupe.Clear()
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Remplis le nom et choisis une OU.", "Attention")
    }
})

$btnSupprimerGroupe.Add_Click({
    if ($lstGroupe.SelectedItem) {
        $entry = $lstGroupe.SelectedItem
        $idx   = $entry.IndexOf($script:SEP)
        $nom   = $entry.Substring(0, $idx).Trim()
        $script:listeGroupes.Remove($entry)
        $lstGroupe.Items.Remove($entry)
        Update-ComboboxGroupes
        Sync-GridColumns
        Nettoyer-OrphansGroupe -GroupeSupprime $nom
    }
})

# ============================================================
# ONGLET UTILISATEURS
# ============================================================

$lblDefautOU          = New-Object System.Windows.Forms.Label
$lblDefautOU.Text     = "OU par defaut :"
$lblDefautOU.Location = New-Object System.Drawing.Point(15, 15)
$lblDefautOU.Size     = New-Object System.Drawing.Size(105, 22)
$tabUser.Controls.Add($lblDefautOU)

$cmbOUDefaut               = New-Object System.Windows.Forms.ComboBox
$cmbOUDefaut.Location      = New-Object System.Drawing.Point(125, 13)
$cmbOUDefaut.Size          = New-Object System.Drawing.Size(150, 22)
$cmbOUDefaut.DropDownStyle = "DropDownList"
$tabUser.Controls.Add($cmbOUDefaut)

$lblDefautGroupe          = New-Object System.Windows.Forms.Label
$lblDefautGroupe.Text     = "Groupe par defaut :"
$lblDefautGroupe.Location = New-Object System.Drawing.Point(295, 15)
$lblDefautGroupe.Size     = New-Object System.Drawing.Size(130, 22)
$tabUser.Controls.Add($lblDefautGroupe)

$cmbGroupeDefaut               = New-Object System.Windows.Forms.ComboBox
$cmbGroupeDefaut.Location      = New-Object System.Drawing.Point(430, 13)
$cmbGroupeDefaut.Size          = New-Object System.Drawing.Size(150, 22)
$cmbGroupeDefaut.DropDownStyle = "DropDownList"
$tabUser.Controls.Add($cmbGroupeDefaut)

$btnAjouterUser          = New-Object System.Windows.Forms.Button
$btnAjouterUser.Text     = "+ Ajouter user"
$btnAjouterUser.Location = New-Object System.Drawing.Point(600, 11)
$btnAjouterUser.Size     = New-Object System.Drawing.Size(120, 26)
$tabUser.Controls.Add($btnAjouterUser)

# Panneau saisie manuelle
$pnlSaisie          = New-Object System.Windows.Forms.Panel
$pnlSaisie.Location = New-Object System.Drawing.Point(15, 45)
$pnlSaisie.Size     = New-Object System.Drawing.Size(735, 35)
$pnlSaisie.Visible  = $false
$tabUser.Controls.Add($pnlSaisie)

foreach ($info in @(
    @{Name="txtSaisiePrenom"; PlaceHolder="Prenom";     X=0},
    @{Name="txtSaisieNom";    PlaceHolder="Nom";        X=110},
    @{Name="txtSaisieLogin";  PlaceHolder="Login";      X=220},
    @{Name="txtSaisieMdp";    PlaceHolder="MotDePasse"; X=330}
)) {
    $t              = New-Object System.Windows.Forms.TextBox
    $t.Location     = New-Object System.Drawing.Point($info.X, 5)
    $t.Size         = New-Object System.Drawing.Size(100, 22)
    $t.Text         = $info.PlaceHolder
    $t.ForeColor    = [System.Drawing.Color]::Gray
    $t.Name         = $info.Name
    $t.Tag          = $info.PlaceHolder
    $t.Add_Enter({
        if ($this.ForeColor -eq [System.Drawing.Color]::Gray) {
            $this.Text = ""; $this.ForeColor = [System.Drawing.Color]::Black
            if ($this.Name -eq "txtSaisieMdp") { $this.PasswordChar = [char]9679 }
        }
    })
    $t.Add_Leave({
        if ($this.Name -eq "txtSaisieMdp") { $this.PasswordChar = [char]0 }
        if ($this.Text -eq "") { $this.ForeColor = [System.Drawing.Color]::Gray; $this.Text = $this.Tag }
    })
    $pnlSaisie.Controls.Add($t)
}

$btnConfirmerUser          = New-Object System.Windows.Forms.Button
$btnConfirmerUser.Text     = "OK"
$btnConfirmerUser.Location = New-Object System.Drawing.Point(440, 3)
$btnConfirmerUser.Size     = New-Object System.Drawing.Size(50, 26)
$pnlSaisie.Controls.Add($btnConfirmerUser)

$btnAnnulerUser          = New-Object System.Windows.Forms.Button
$btnAnnulerUser.Text     = "Annuler"
$btnAnnulerUser.Location = New-Object System.Drawing.Point(498, 3)
$btnAnnulerUser.Size     = New-Object System.Drawing.Size(65, 26)
$pnlSaisie.Controls.Add($btnAnnulerUser)

$btnAjouterUser.Add_Click({ $pnlSaisie.Visible = $true })

$btnAnnulerUser.Add_Click({
    foreach ($ctrl in $pnlSaisie.Controls) {
        if ($ctrl -is [System.Windows.Forms.TextBox]) {
            $ctrl.PasswordChar  = [char]0
            $ctrl.Text          = $ctrl.Tag
            $ctrl.ForeColor     = [System.Drawing.Color]::Gray
        }
    }
    $pnlSaisie.Visible = $false
})

$btnConfirmerUser.Add_Click({
    $prenom = $pnlSaisie.Controls["txtSaisiePrenom"].Text.Trim()
    $nom2   = $pnlSaisie.Controls["txtSaisieNom"].Text.Trim()
    $login  = $pnlSaisie.Controls["txtSaisieLogin"].Text.Trim()
    $mdp    = $pnlSaisie.Controls["txtSaisieMdp"].Text.Trim()
    $ou     = if ($cmbOUDefaut.SelectedItem)     { $cmbOUDefaut.SelectedItem }     else { "" }
    $grp    = if ($cmbGroupeDefaut.SelectedItem) { $cmbGroupeDefaut.SelectedItem } else { "" }

    $placeholders = @("Prenom","Nom","Login","MotDePasse")
    if ($login -ne "" -and $login -notin $placeholders) {
        $grid.Rows.Add($prenom, $nom2, $login, $mdp, $ou, $grp) | Out-Null
        foreach ($ctrl in $pnlSaisie.Controls) {
            if ($ctrl -is [System.Windows.Forms.TextBox]) {
                $ctrl.PasswordChar = [char]0
                $ctrl.Text         = $ctrl.Tag
                $ctrl.ForeColor    = [System.Drawing.Color]::Gray
            }
        }
        $pnlSaisie.Visible = $false
    } else {
        [System.Windows.Forms.MessageBox]::Show("Login obligatoire.", "Attention")
    }
})

# DataGridView
$grid                              = New-Object System.Windows.Forms.DataGridView
$grid.Location                     = New-Object System.Drawing.Point(15, 88)
$grid.Size                         = New-Object System.Drawing.Size(735, 355)
$grid.AllowUserToAddRows           = $false
$grid.AllowUserToDeleteRows        = $true
$grid.AutoSizeColumnsMode          = "Fill"
$grid.ColumnHeadersHeightSizeMode  = "AutoSize"
$grid.Font                         = New-Object System.Drawing.Font("Segoe UI", 8.5)
$grid.SelectionMode                = "FullRowSelect"

foreach ($col in @("Prenom","Nom","Login","MotDePasse")) {
    $c            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText = $col; $c.Name = $col
    $grid.Columns.Add($c) | Out-Null
}
foreach ($col in @("OU","Groupe")) {
    $c           = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $c.HeaderText = $col; $c.Name = $col
    $c.FlatStyle  = "Flat"
    $grid.Columns.Add($c) | Out-Null
}
$tabUser.Controls.Add($grid)

# Masquage mot de passe dans la grille
$grid.Add_CellFormatting({
    param($s, $e)
    if ($e.ColumnIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -eq "MotDePasse") {
        if ($e.Value -ne $null -and $e.Value.ToString() -ne "") {
            $e.Value = [string]::new([char]9679, $e.Value.ToString().Length)
            $e.FormattingApplied = $true
        }
    }
})

$grid.Add_EditingControlShowing({
    param($s, $e)
    if ($grid.CurrentCell -ne $null -and $grid.Columns[$grid.CurrentCell.ColumnIndex].Name -eq "MotDePasse") {
        if ($e.Control -is [System.Windows.Forms.TextBox]) {
            ([System.Windows.Forms.TextBox]$e.Control).PasswordChar = [char]9679
        }
    } elseif ($e.Control -is [System.Windows.Forms.TextBox]) {
        ([System.Windows.Forms.TextBox]$e.Control).PasswordChar = [char]0
    }
})

# Supprime le dialog d'erreur DataGridView pour les combobox orphelines
$grid.Add_DataError({
    param($s, $e)
    $e.Cancel = $true
})

# ============================================================
# BARRE CONFIG CSV
# ============================================================
$grpBoxConfig          = New-Object System.Windows.Forms.GroupBox
$grpBoxConfig.Text     = "Fichier de configuration (config_ad.csv)"
$grpBoxConfig.Location = New-Object System.Drawing.Point(10, 540)
$grpBoxConfig.Size     = New-Object System.Drawing.Size(780, 65)
$grpBoxConfig.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxConfig)

$lblConfigActif           = New-Object System.Windows.Forms.Label
$lblConfigActif.Text      = "Config : $($script:configPath)"
$lblConfigActif.Location  = New-Object System.Drawing.Point(15, 30)
$lblConfigActif.Size      = New-Object System.Drawing.Size(290, 18)
$lblConfigActif.ForeColor = [System.Drawing.Color]::Gray
$grpBoxConfig.Controls.Add($lblConfigActif)

$btnLireAD          = New-Object System.Windows.Forms.Button
$btnLireAD.Text     = "Lire depuis AD"
$btnLireAD.Location = New-Object System.Drawing.Point(315, 22)
$btnLireAD.Size     = New-Object System.Drawing.Size(105, 26)
$grpBoxConfig.Controls.Add($btnLireAD)

$btnExporterConfig          = New-Object System.Windows.Forms.Button
$btnExporterConfig.Text     = "Exporter config"
$btnExporterConfig.Location = New-Object System.Drawing.Point(430, 22)
$btnExporterConfig.Size     = New-Object System.Drawing.Size(105, 26)
$grpBoxConfig.Controls.Add($btnExporterConfig)

$btnChargerConfig          = New-Object System.Windows.Forms.Button
$btnChargerConfig.Text     = "Charger config"
$btnChargerConfig.Location = New-Object System.Drawing.Point(545, 22)
$btnChargerConfig.Size     = New-Object System.Drawing.Size(105, 26)
$grpBoxConfig.Controls.Add($btnChargerConfig)

$btnOuvrirConfig          = New-Object System.Windows.Forms.Button
$btnOuvrirConfig.Text     = "Ouvrir config"
$btnOuvrirConfig.Location = New-Object System.Drawing.Point(660, 22)
$btnOuvrirConfig.Size     = New-Object System.Drawing.Size(100, 26)
$grpBoxConfig.Controls.Add($btnOuvrirConfig)

$btnLireAD.Add_Click({ Import-DepuisAD })

$btnExporterConfig.Add_Click({
    Export-Config
    $lblConfigActif.Text      = "Config sauvegardee : $($script:configPath)"
    $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
    [System.Windows.Forms.MessageBox]::Show("Config exportee :`n$($script:configPath)", "Export OK")
})

$btnChargerConfig.Add_Click({
    $ofd          = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter   = "CSV|*.csv"
    $ofd.FileName = "config_ad.csv"
    if ($ofd.ShowDialog() -eq "OK") {
        $script:configPath        = $ofd.FileName
        $lblConfigActif.Text      = "Config : $($script:configPath)"
        $lblConfigActif.ForeColor = [System.Drawing.Color]::Gray
        Import-Config -Chemin $script:configPath
        [System.Windows.Forms.MessageBox]::Show("Config chargee :`n$($script:configPath)", "Chargement OK")
    }
})

$btnOuvrirConfig.Add_Click({
    if (Test-Path $script:configPath) {
        Start-Process notepad.exe $script:configPath
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Aucune config trouvee.`nExporte d'abord une config.", "Introuvable")
    }
})

# ============================================================
# OPTION PLACEMENT DES GROUPES
# ============================================================
$grpBoxOption          = New-Object System.Windows.Forms.GroupBox
$grpBoxOption.Text     = "Placement des groupes dans l'AD"
$grpBoxOption.Location = New-Object System.Drawing.Point(10, 615)
$grpBoxOption.Size     = New-Object System.Drawing.Size(780, 80)
$grpBoxOption.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxOption)

$radioAvecOU          = New-Object System.Windows.Forms.RadioButton
$radioAvecOU.Text     = "Dans l'OU des utilisateurs"
$radioAvecOU.Location = New-Object System.Drawing.Point(15, 22)
$radioAvecOU.Size     = New-Object System.Drawing.Size(250, 22)
$radioAvecOU.Checked  = $true
$grpBoxOption.Controls.Add($radioAvecOU)

$radioDediee          = New-Object System.Windows.Forms.RadioButton
$radioDediee.Text     = "Dans une OU dediee :"
$radioDediee.Location = New-Object System.Drawing.Point(15, 48)
$radioDediee.Size     = New-Object System.Drawing.Size(160, 22)
$grpBoxOption.Controls.Add($radioDediee)

$txtOUDediee          = New-Object System.Windows.Forms.TextBox
$txtOUDediee.Text     = "Groupes"
$txtOUDediee.Location = New-Object System.Drawing.Point(180, 46)
$txtOUDediee.Size     = New-Object System.Drawing.Size(130, 22)
$txtOUDediee.Enabled  = $false
$grpBoxOption.Controls.Add($txtOUDediee)

$lblOUDedieeInfo           = New-Object System.Windows.Forms.Label
$lblOUDedieeInfo.Text      = "(creee automatiquement si elle n'existe pas)"
$lblOUDedieeInfo.Location  = New-Object System.Drawing.Point(320, 50)
$lblOUDedieeInfo.Size      = New-Object System.Drawing.Size(380, 18)
$lblOUDedieeInfo.ForeColor = [System.Drawing.Color]::Gray
$grpBoxOption.Controls.Add($lblOUDedieeInfo)

$radioDediee.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })
$radioAvecOU.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })

# ============================================================
# BARRE DE PROGRESSION
# ============================================================
$progressBar               = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location      = New-Object System.Drawing.Point(10, 703)
$progressBar.Size          = New-Object System.Drawing.Size(780, 16)
$progressBar.Minimum       = 0
$progressBar.Maximum       = 100
$progressBar.Value         = 0
$form.Controls.Add($progressBar)

# ============================================================
# BOUTON LANCER + LOG
# ============================================================
$btnLancer              = New-Object System.Windows.Forms.Button
$btnLancer.Text         = "Lancer le provisioning"
$btnLancer.Location     = New-Object System.Drawing.Point(10, 727)
$btnLancer.Size         = New-Object System.Drawing.Size(200, 34)
$btnLancer.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnLancer.ForeColor    = [System.Drawing.Color]::White
$btnLancer.FlatStyle    = "Flat"
$btnLancer.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnLancer)

$txtLog               = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location      = New-Object System.Drawing.Point(220, 725)
$txtLog.Size          = New-Object System.Drawing.Size(570, 90)
$txtLog.ReadOnly      = $true
$txtLog.BackColor     = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor     = [System.Drawing.Color]::LightGreen
$txtLog.Font          = New-Object System.Drawing.Font("Consolas", 8)
$txtLog.ScrollBars    = "Vertical"
$form.Controls.Add($txtLog)

# ============================================================
# LOGIQUE PROVISIONING
# ============================================================
$btnLancer.Add_Click({

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Lancer le provisioning sur l'AD ?`nCette operation cree / modifie des comptes et des groupes.",
        "Confirmation", "YesNo", "Question")
    if ($confirm -ne "Yes") { return }

    $btnLancer.Enabled   = $false
    $progressBar.Value   = 0

    try { $domaine = Get-ADDomain -ErrorAction Stop }
    catch {
        Write-Log "Impossible de contacter l'AD : $_" "ERR"
        $btnLancer.Enabled = $true
        return
    }
    $dcRoot      = $domaine.DistinguishedName
    $dnsDomain   = $domaine.DNSRoot
    $modeDediee  = $radioDediee.Checked
    $nomOUDediee = $txtOUDediee.Text.Trim()

    # Calcul du total pour la barre de progression
    $totalSteps = $script:listeOUs.Count + $script:listeGroupes.Count + $grid.Rows.Count
    if ($modeDediee) { $totalSteps++ }
    $step = 0

    Write-Log "=== DEBUT PROVISIONING ===" "INFO"
    Write-Log "Domaine : $dcRoot" "INFO"
    if ($modeDediee) { Write-Log "Mode : groupes dans OU dediee '$nomOUDediee'" "INFO" }
    else             { Write-Log "Mode : groupes dans l'OU de leurs utilisateurs" "INFO" }

    # -- ETAPE 1 : OUs
    Write-Log "--- Etape 1 : OUs ---" "INFO"
    foreach ($ou in $script:listeOUs) {
        $ouDN = "OU=$ou,$dcRoot"
        try {
            Get-ADOrganizationalUnit -Identity $ouDN -ErrorAction Stop | Out-Null
            Write-Log "OU existe deja : $ou" "WARN"
        } catch {
            try {
                New-ADOrganizationalUnit -Name $ou -Path $dcRoot
                Write-Log "OU creee : $ou" "OK"
            } catch { Write-Log "Erreur OU $ou : $_" "ERR" }
        }
        $step++
        $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
        [System.Windows.Forms.Application]::DoEvents()
    }

    # -- ETAPE 2 : OU dediee groupes
    if ($modeDediee) {
        Write-Log "--- Etape 2 : OU dediee groupes ---" "INFO"
        $ouDedieDN = "OU=$nomOUDediee,$dcRoot"
        try {
            Get-ADOrganizationalUnit -Identity $ouDedieDN -ErrorAction Stop | Out-Null
            Write-Log "OU dediee existe deja : $nomOUDediee" "WARN"
        } catch {
            try {
                New-ADOrganizationalUnit -Name $nomOUDediee -Path $dcRoot
                Write-Log "OU dediee creee : $nomOUDediee" "OK"
            } catch { Write-Log "Erreur OU dediee : $_" "ERR" }
        }
        $step++
        $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
        [System.Windows.Forms.Application]::DoEvents()
    }

    # -- ETAPE 3 : Groupes
    Write-Log "--- Etape 3 : Groupes ---" "INFO"
    foreach ($entry in $script:listeGroupes) {
        $idx    = $entry.IndexOf($script:SEP)
        $nom    = $entry.Substring(0, $idx).Trim()
        $ouUser = $entry.Substring($idx + $script:SEP.Length).Trim()
        $ouDN   = if ($modeDediee) { "OU=$nomOUDediee,$dcRoot" } else { "OU=$ouUser,$dcRoot" }
        try {
            Get-ADGroup -Identity $nom -ErrorAction Stop | Out-Null
            Write-Log "Groupe existe deja : $nom" "WARN"
        } catch {
            try {
                New-ADGroup -Name $nom -SamAccountName $nom -GroupScope Global -GroupCategory Security -Path $ouDN
                Write-Log "Groupe cree : $nom" "OK"
            } catch { Write-Log "Erreur groupe $nom : $_" "ERR" }
        }
        $step++
        $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
        [System.Windows.Forms.Application]::DoEvents()
    }

    # -- ETAPE 4 : Utilisateurs
    Write-Log "--- Etape 4 : Utilisateurs ---" "INFO"
    foreach ($row in $grid.Rows) {
        $prenom = $row.Cells["Prenom"].Value
        $nom    = $row.Cells["Nom"].Value
        $login  = $row.Cells["Login"].Value
        $mdp    = $row.Cells["MotDePasse"].Value
        $ou     = $row.Cells["OU"].Value
        $grp    = $row.Cells["Groupe"].Value

        if (-not $login -or -not $ou -or -not $grp) {
            Write-Log "Ligne incomplete, ignoree (login=$login)." "WARN"
            $step++
            $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
            [System.Windows.Forms.Application]::DoEvents()
            continue
        }

        $ouDN     = "OU=$ou,$dcRoot"
        $fullName = "$prenom $nom"
        $existing = Get-ADUser -Filter { SamAccountName -eq $login } `
                               -Properties DistinguishedName -ErrorAction SilentlyContinue

        if ($existing) {
            # Deplacer si l'OU a change
            $currentParent = $existing.DistinguishedName.Substring(
                $existing.DistinguishedName.IndexOf(',') + 1)
            if ($currentParent -ne $ouDN) {
                try {
                    Move-ADObject -Identity $existing.DistinguishedName -TargetPath $ouDN
                    Write-Log "  -> $login deplace vers $ou" "OK"
                } catch { Write-Log "  -> Erreur deplacement $login : $_" "ERR" }
            }
            try {
                Set-ADUser -Identity $login -GivenName $prenom -Surname $nom -DisplayName $fullName
                if (-not [string]::IsNullOrWhiteSpace($mdp)) {
                    $securePwd = ConvertTo-SecureString $mdp -AsPlainText -Force
                    Set-ADAccountPassword -Identity $login -NewPassword $securePwd -Reset
                }
                Write-Log "User mis a jour : $login" "WARN"
            } catch { Write-Log "Erreur MAJ $login : $_" "ERR"; continue }
        } else {
            if ([string]::IsNullOrWhiteSpace($mdp)) {
                Write-Log "User $login ignore : mot de passe manquant pour la creation." "WARN"
                $step++
                $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }
            $securePwd = ConvertTo-SecureString $mdp -AsPlainText -Force
            try {
                New-ADUser -SamAccountName $login `
                           -UserPrincipalName "$login@$dnsDomain" `
                           -GivenName $prenom -Surname $nom -DisplayName $fullName -Name $fullName `
                           -AccountPassword $securePwd -Path $ouDN -Enabled $true
                Write-Log "User cree : $login" "OK"
            } catch { Write-Log "Erreur creation $login : $_" "ERR"; continue }
        }

        try {
            Add-ADGroupMember -Identity $grp -Members $login -ErrorAction Stop
            Write-Log "  -> $login ajoute a $grp" "OK"
        } catch { Write-Log "  -> Erreur groupe $grp pour $login : $_" "ERR" }

        $step++
        $progressBar.Value = [int]([Math]::Min(100, ($step / $totalSteps * 100)))
        [System.Windows.Forms.Application]::DoEvents()
    }

    $progressBar.Value = 100
    Write-Log "=== FIN PROVISIONING ===" "INFO"
    Write-Log "Log : $script:logPath" "INFO"
    Export-Config
    Write-Log "Config sauvegardee : $($script:configPath)" "INFO"
    $btnLancer.Enabled = $true
})

# ============================================================
# CHARGEMENT AUTO AU DEMARRAGE
# ============================================================
$form.Add_Shown({
    if (Test-Path $script:configPath) {
        Import-Config -Chemin $script:configPath
        $lblConfigActif.Text      = "Config auto-chargee : $($script:configPath)"
        $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
    }
})

# ============================================================
$form.ShowDialog() | Out-Null
