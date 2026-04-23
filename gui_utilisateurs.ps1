# ============================================================
#  PROVISIONING AD - GUI Windows Forms v4
#  CSV unique : OUs / Groupes / Utilisateurs
#  Auteur    : anotherj4ck
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# ============================================================
# CHEMINS
# ============================================================
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
    $txtLog.AppendText("$line`r`n")
    $txtLog.ScrollToCaret()
    $dir = Split-Path $script:logPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -Path $script:logPath -Value $line
}

function Rafraichir-ComboboxOUs {
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    foreach ($o in $script:listeOUs) {
        $cmbOUGroupe.Items.Add($o) | Out-Null
        $cmbOUDefaut.Items.Add($o) | Out-Null
    }
}

function Rafraichir-ComboboxGroupes {
    $cmbGroupeDefaut.Items.Clear()
    foreach ($entry in $script:listeGroupes) {
        $separateur = " -> "
        $idx = $entry.IndexOf($separateur)
        $nom = $entry.Substring(0, $idx).Trim()
        $cmbGroupeDefaut.Items.Add($nom) | Out-Null
    }
}

function Exporter-Config {
    $lignes = [System.Collections.ArrayList]@()
    $lignes.Add("Type;Nom;OU;Login;Prenom;Nom2;MotDePasse;Groupe") | Out-Null

    foreach ($ou in $script:listeOUs) {
        $lignes.Add("OU;$ou;;;;;;;") | Out-Null
    }

    foreach ($entry in $script:listeGroupes) {
        $separateur = " -> "
        $idx = $entry.IndexOf($separateur)
        $nom = $entry.Substring(0, $idx).Trim()
        $ou  = $entry.Substring($idx + $separateur.Length).Trim()
        $lignes.Add("GROUPE;$nom;$ou;;;;;;") | Out-Null
    }

    foreach ($row in $grid.Rows) {
        $prenom = $row.Cells["Prenom"].Value
        $nom    = $row.Cells["Nom"].Value
        $login  = $row.Cells["Login"].Value
        $mdp    = $row.Cells["MotDePasse"].Value
        $ou     = $row.Cells["OU"].Value
        $grp    = $row.Cells["Groupe"].Value
        if ($login) {
            $lignes.Add("USER;;$ou;$login;$prenom;$nom;$mdp;$grp") | Out-Null
        }
    }

    $lignes | Set-Content -Path $script:configPath -Encoding UTF8
}

function Charger-Config {
    param([string]$Chemin)
    if (-not (Test-Path $Chemin)) { return }

    $data = Import-Csv -Path $Chemin -Delimiter ";"

    # Reset
    $script:listeOUs.Clear()
    $script:listeGroupes.Clear()
    $lstOU.Items.Clear()
    $lstGroupe.Items.Clear()
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    $cmbGroupeDefaut.Items.Clear()
    $grid.Rows.Clear()

    # Passe 1 : OUs
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

    # Passe 2 : Groupes
    foreach ($row in $data) {
        if ($row.Type -eq "GROUPE") {
            $nom   = $row.Nom.Trim()
            $ou    = $row.OU.Trim()
            $entry = "$nom -> $ou"
            if ($nom -ne "" -and $ou -ne "" -and $script:listeGroupes -notcontains $entry) {
                $script:listeGroupes.Add($entry) | Out-Null
                $lstGroupe.Items.Add($entry) | Out-Null
                $cmbGroupeDefaut.Items.Add($nom) | Out-Null
            }
        }
    }

    # Passe 3 : Users — remplir combobox colonnes d'abord
    $colOU     = $grid.Columns["OU"]
    $colGroupe = $grid.Columns["Groupe"]
    $colOU.Items.Clear()
    $colGroupe.Items.Clear()
    foreach ($o in $script:listeOUs)     { $colOU.Items.Add($o) | Out-Null }
    foreach ($entry in $script:listeGroupes) {
        $separateur = " -> "
        $idx = $entry.IndexOf($separateur)
        $nom = $entry.Substring(0, $idx).Trim()
        $colGroupe.Items.Add($nom) | Out-Null
    }

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

# ============================================================
# FENETRE PRINCIPALE
# ============================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Provisioning AD v4"
$form.Size            = New-Object System.Drawing.Size(820, 820)
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

$btnAjouterOU           = New-Object System.Windows.Forms.Button
$btnAjouterOU.Text      = "Ajouter"
$btnAjouterOU.Location  = New-Object System.Drawing.Point(355, 16)
$btnAjouterOU.Size      = New-Object System.Drawing.Size(80, 26)
$tabOU.Controls.Add($btnAjouterOU)

$btnSupprimerOU         = New-Object System.Windows.Forms.Button
$btnSupprimerOU.Text    = "Supprimer"
$btnSupprimerOU.Location= New-Object System.Drawing.Point(445, 16)
$btnSupprimerOU.Size    = New-Object System.Drawing.Size(80, 26)
$tabOU.Controls.Add($btnSupprimerOU)

$lstOU          = New-Object System.Windows.Forms.ListBox
$lstOU.Location = New-Object System.Drawing.Point(15, 55)
$lstOU.Size     = New-Object System.Drawing.Size(510, 380)
$lstOU.Font     = New-Object System.Drawing.Font("Consolas", 9)
$tabOU.Controls.Add($lstOU)

$btnAjouterOU.Add_Click({
    $val = $txtOU.Text.Trim()
    if ($val -ne "" -and $script:listeOUs -notcontains $val) {
        $script:listeOUs.Add($val) | Out-Null
        $lstOU.Items.Add($val) | Out-Null
        Rafraichir-ComboboxOUs
        # Sync colonnes grid
        $colOU = $grid.Columns["OU"]
        if ($colOU.Items -notcontains $val) { $colOU.Items.Add($val) | Out-Null }
        $txtOU.Clear()
    }
})

$btnSupprimerOU.Add_Click({
    if ($lstOU.SelectedItem) {
        $val = $lstOU.SelectedItem
        $script:listeOUs.Remove($val)
        $lstOU.Items.Remove($val)
        Rafraichir-ComboboxOUs
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

$lblOUGroupe            = New-Object System.Windows.Forms.Label
$lblOUGroupe.Text       = "OU cible :"
$lblOUGroupe.Location   = New-Object System.Drawing.Point(315, 20)
$lblOUGroupe.Size       = New-Object System.Drawing.Size(65, 22)
$tabGroupe.Controls.Add($lblOUGroupe)

$cmbOUGroupe            = New-Object System.Windows.Forms.ComboBox
$cmbOUGroupe.Location   = New-Object System.Drawing.Point(385, 17)
$cmbOUGroupe.Size       = New-Object System.Drawing.Size(140, 22)
$cmbOUGroupe.DropDownStyle = "DropDownList"
$tabGroupe.Controls.Add($cmbOUGroupe)

$btnAjouterGroupe           = New-Object System.Windows.Forms.Button
$btnAjouterGroupe.Text      = "Ajouter"
$btnAjouterGroupe.Location  = New-Object System.Drawing.Point(540, 15)
$btnAjouterGroupe.Size      = New-Object System.Drawing.Size(75, 26)
$tabGroupe.Controls.Add($btnAjouterGroupe)

$btnSupprimerGroupe         = New-Object System.Windows.Forms.Button
$btnSupprimerGroupe.Text    = "Supprimer"
$btnSupprimerGroupe.Location= New-Object System.Drawing.Point(625, 15)
$btnSupprimerGroupe.Size    = New-Object System.Drawing.Size(80, 26)
$tabGroupe.Controls.Add($btnSupprimerGroupe)

$lstGroupe          = New-Object System.Windows.Forms.ListBox
$lstGroupe.Location = New-Object System.Drawing.Point(15, 55)
$lstGroupe.Size     = New-Object System.Drawing.Size(720, 380)
$lstGroupe.Font     = New-Object System.Drawing.Font("Consolas", 9)
$tabGroupe.Controls.Add($lstGroupe)

$btnAjouterGroupe.Add_Click({
    $nom = $txtGroupe.Text.Trim()
    $ou  = $cmbOUGroupe.SelectedItem
    if ($nom -ne "" -and $ou) {
        $entry = "$nom -> $ou"
        if ($script:listeGroupes -notcontains $entry) {
            $script:listeGroupes.Add($entry) | Out-Null
            $lstGroupe.Items.Add($entry) | Out-Null
            $cmbGroupeDefaut.Items.Add($nom) | Out-Null
            # Sync colonne grid
            $colGroupe = $grid.Columns["Groupe"]
            if ($colGroupe.Items -notcontains $nom) { $colGroupe.Items.Add($nom) | Out-Null }
            $txtGroupe.Clear()
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Remplis le nom et choisis une OU.", "Attention")
    }
})

$btnSupprimerGroupe.Add_Click({
    if ($lstGroupe.SelectedItem) {
        $entry = $lstGroupe.SelectedItem
        $separateur = " -> "
        $idx = $entry.IndexOf($separateur)
        $nom = $entry.Substring(0, $idx).Trim()
        $script:listeGroupes.Remove($entry)
        $lstGroupe.Items.Remove($entry)
        Rafraichir-ComboboxGroupes
    }
})

# ============================================================
# ONGLET UTILISATEURS
# ============================================================

$lblDefautOU            = New-Object System.Windows.Forms.Label
$lblDefautOU.Text       = "OU par défaut :"
$lblDefautOU.Location   = New-Object System.Drawing.Point(15, 15)
$lblDefautOU.Size       = New-Object System.Drawing.Size(105, 22)
$tabUser.Controls.Add($lblDefautOU)

$cmbOUDefaut            = New-Object System.Windows.Forms.ComboBox
$cmbOUDefaut.Location   = New-Object System.Drawing.Point(125, 13)
$cmbOUDefaut.Size       = New-Object System.Drawing.Size(150, 22)
$cmbOUDefaut.DropDownStyle = "DropDownList"
$tabUser.Controls.Add($cmbOUDefaut)

$lblDefautGroupe            = New-Object System.Windows.Forms.Label
$lblDefautGroupe.Text       = "Groupe par défaut :"
$lblDefautGroupe.Location   = New-Object System.Drawing.Point(295, 15)
$lblDefautGroupe.Size       = New-Object System.Drawing.Size(130, 22)
$tabUser.Controls.Add($lblDefautGroupe)

$cmbGroupeDefaut            = New-Object System.Windows.Forms.ComboBox
$cmbGroupeDefaut.Location   = New-Object System.Drawing.Point(430, 13)
$cmbGroupeDefaut.Size       = New-Object System.Drawing.Size(150, 22)
$cmbGroupeDefaut.DropDownStyle = "DropDownList"
$tabUser.Controls.Add($cmbGroupeDefaut)

# Bouton ajouter user manuellement
$btnAjouterUser         = New-Object System.Windows.Forms.Button
$btnAjouterUser.Text    = "+ Ajouter user"
$btnAjouterUser.Location= New-Object System.Drawing.Point(600, 11)
$btnAjouterUser.Size    = New-Object System.Drawing.Size(120, 26)
$tabUser.Controls.Add($btnAjouterUser)

# Champs saisie manuelle user
$pnlSaisie              = New-Object System.Windows.Forms.Panel
$pnlSaisie.Location     = New-Object System.Drawing.Point(15, 45)
$pnlSaisie.Size         = New-Object System.Drawing.Size(735, 35)
$pnlSaisie.Visible      = $false
$tabUser.Controls.Add($pnlSaisie)

foreach ($info in @(
    @{Name="txtSaisiePrenom"; PlaceHolder="Prenom"; X=0},
    @{Name="txtSaisieNom";    PlaceHolder="Nom";    X=110},
    @{Name="txtSaisieLogin";  PlaceHolder="Login";  X=220},
    @{Name="txtSaisieMdp";    PlaceHolder="MotDePasse"; X=330}
)) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = New-Object System.Drawing.Point($info.X, 5)
    $t.Size        = New-Object System.Drawing.Size(100, 22)
    $t.Text        = $info.PlaceHolder
    $t.ForeColor   = [System.Drawing.Color]::Gray
    $t.Name        = $info.Name
    $t.Add_Enter({ if ($this.ForeColor -eq [System.Drawing.Color]::Gray) { $this.Text = ""; $this.ForeColor = [System.Drawing.Color]::Black } })
    $t.Add_Leave({ if ($this.Text -eq "") { $this.ForeColor = [System.Drawing.Color]::Gray; $this.Text = $this.Name -replace "txtSaisie","" } })
    $pnlSaisie.Controls.Add($t)
}

$btnConfirmerUser       = New-Object System.Windows.Forms.Button
$btnConfirmerUser.Text  = "OK"
$btnConfirmerUser.Location = New-Object System.Drawing.Point(440, 3)
$btnConfirmerUser.Size  = New-Object System.Drawing.Size(50, 26)
$pnlSaisie.Controls.Add($btnConfirmerUser)

$btnAjouterUser.Add_Click({ $pnlSaisie.Visible = $true })

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
                $ctrl.Text = $ctrl.Name -replace "txtSaisie",""
                $ctrl.ForeColor = [System.Drawing.Color]::Gray
            }
        }
        $pnlSaisie.Visible = $false
    } else {
        [System.Windows.Forms.MessageBox]::Show("Login obligatoire.", "Attention")
    }
})

# DataGridView
$grid                   = New-Object System.Windows.Forms.DataGridView
$grid.Location          = New-Object System.Drawing.Point(15, 88)
$grid.Size              = New-Object System.Drawing.Size(735, 355)
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $true
$grid.AutoSizeColumnsMode   = "Fill"
$grid.ColumnHeadersHeightSizeMode = "AutoSize"
$grid.Font              = New-Object System.Drawing.Font("Segoe UI", 8.5)
$grid.SelectionMode     = "FullRowSelect"

foreach ($col in @("Prenom","Nom","Login","MotDePasse")) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText = $col; $c.Name = $col
    $grid.Columns.Add($c) | Out-Null
}
foreach ($col in @("OU","Groupe")) {
    $c = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $c.HeaderText = $col; $c.Name = $col
    $c.FlatStyle  = "Flat"
    $grid.Columns.Add($c) | Out-Null
}
$tabUser.Controls.Add($grid)

# ============================================================
# BARRE CONFIG CSV
# ============================================================
$grpBoxConfig           = New-Object System.Windows.Forms.GroupBox
$grpBoxConfig.Text      = "Fichier de configuration (config_ad.csv)"
$grpBoxConfig.Location  = New-Object System.Drawing.Point(10, 540)
$grpBoxConfig.Size      = New-Object System.Drawing.Size(780, 65)
$grpBoxConfig.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxConfig)

$lblConfigActif             = New-Object System.Windows.Forms.Label
$lblConfigActif.Text        = "Config : $($script:configPath)"
$lblConfigActif.Location    = New-Object System.Drawing.Point(15, 30)
$lblConfigActif.Size        = New-Object System.Drawing.Size(360, 18)
$lblConfigActif.ForeColor   = [System.Drawing.Color]::Gray
$grpBoxConfig.Controls.Add($lblConfigActif)

$btnExporterConfig          = New-Object System.Windows.Forms.Button
$btnExporterConfig.Text     = "Exporter config"
$btnExporterConfig.Location = New-Object System.Drawing.Point(385, 22)
$btnExporterConfig.Size     = New-Object System.Drawing.Size(120, 26)
$grpBoxConfig.Controls.Add($btnExporterConfig)

$btnChargerConfig           = New-Object System.Windows.Forms.Button
$btnChargerConfig.Text      = "Charger config"
$btnChargerConfig.Location  = New-Object System.Drawing.Point(515, 22)
$btnChargerConfig.Size      = New-Object System.Drawing.Size(120, 26)
$grpBoxConfig.Controls.Add($btnChargerConfig)

$btnOuvrirConfig            = New-Object System.Windows.Forms.Button
$btnOuvrirConfig.Text       = "Ouvrir config"
$btnOuvrirConfig.Location   = New-Object System.Drawing.Point(645, 22)
$btnOuvrirConfig.Size       = New-Object System.Drawing.Size(110, 26)
$grpBoxConfig.Controls.Add($btnOuvrirConfig)

$btnExporterConfig.Add_Click({
    Exporter-Config
    $lblConfigActif.Text      = "Config sauvegardée : $($script:configPath)"
    $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
    [System.Windows.Forms.MessageBox]::Show("Config exportée :`n$($script:configPath)", "Export OK")
})

$btnChargerConfig.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter   = "CSV|*.csv"
    $ofd.FileName = "config_ad.csv"
    if ($ofd.ShowDialog() -eq "OK") {
        $script:configPath        = $ofd.FileName
        $lblConfigActif.Text      = "Config : $($script:configPath)"
        $lblConfigActif.ForeColor = [System.Drawing.Color]::Gray
        Charger-Config -Chemin $script:configPath
        [System.Windows.Forms.MessageBox]::Show("Config chargée :`n$($script:configPath)", "Chargement OK")
    }
})

$btnOuvrirConfig.Add_Click({
    if (Test-Path $script:configPath) {
        Start-Process notepad.exe $script:configPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Aucune config trouvée.`nExporte d'abord une config.", "Introuvable")
    }
})

# ============================================================
# OPTION PLACEMENT DES GROUPES
# ============================================================
$grpBoxOption           = New-Object System.Windows.Forms.GroupBox
$grpBoxOption.Text      = "Placement des groupes dans l'AD"
$grpBoxOption.Location  = New-Object System.Drawing.Point(10, 615)
$grpBoxOption.Size      = New-Object System.Drawing.Size(780, 80)
$grpBoxOption.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxOption)

$radioAvecOU            = New-Object System.Windows.Forms.RadioButton
$radioAvecOU.Text       = "Dans l'OU des utilisateurs"
$radioAvecOU.Location   = New-Object System.Drawing.Point(15, 22)
$radioAvecOU.Size       = New-Object System.Drawing.Size(250, 22)
$radioAvecOU.Checked    = $true
$grpBoxOption.Controls.Add($radioAvecOU)

$radioDediee            = New-Object System.Windows.Forms.RadioButton
$radioDediee.Text       = "Dans une OU dédiée :"
$radioDediee.Location   = New-Object System.Drawing.Point(15, 48)
$radioDediee.Size       = New-Object System.Drawing.Size(160, 22)
$grpBoxOption.Controls.Add($radioDediee)

$txtOUDediee            = New-Object System.Windows.Forms.TextBox
$txtOUDediee.Text       = "Groupes"
$txtOUDediee.Location   = New-Object System.Drawing.Point(180, 46)
$txtOUDediee.Size       = New-Object System.Drawing.Size(130, 22)
$txtOUDediee.Enabled    = $false
$grpBoxOption.Controls.Add($txtOUDediee)

$lblOUDedieeInfo        = New-Object System.Windows.Forms.Label
$lblOUDedieeInfo.Text   = "(créée automatiquement si elle n'existe pas)"
$lblOUDedieeInfo.Location = New-Object System.Drawing.Point(320, 50)
$lblOUDedieeInfo.Size   = New-Object System.Drawing.Size(380, 18)
$lblOUDedieeInfo.ForeColor = [System.Drawing.Color]::Gray
$grpBoxOption.Controls.Add($lblOUDedieeInfo)

$radioDediee.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })
$radioAvecOU.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })

# ============================================================
# BOUTON LANCER + LOG
# ============================================================
$btnLancer              = New-Object System.Windows.Forms.Button
$btnLancer.Text         = "Lancer le provisioning"
$btnLancer.Location     = New-Object System.Drawing.Point(10, 708)
$btnLancer.Size         = New-Object System.Drawing.Size(200, 34)
$btnLancer.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnLancer.ForeColor    = [System.Drawing.Color]::White
$btnLancer.FlatStyle    = "Flat"
$btnLancer.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnLancer)

$txtLog                 = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location        = New-Object System.Drawing.Point(220, 706)
$txtLog.Size            = New-Object System.Drawing.Size(570, 90)
$txtLog.ReadOnly        = $true
$txtLog.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor       = [System.Drawing.Color]::LightGreen
$txtLog.Font            = New-Object System.Drawing.Font("Consolas", 8)
$txtLog.ScrollBars      = "Vertical"
$form.Controls.Add($txtLog)

# ============================================================
# LOGIQUE PROVISIONING
# ============================================================
$btnLancer.Add_Click({

    $dcRoot      = (Get-ADDomain).DistinguishedName
    $modeDediee  = $radioDediee.Checked
    $nomOUDediee = $txtOUDediee.Text.Trim()

    Write-Log "=== DEBUT PROVISIONING ===" "INFO"
    Write-Log "Domaine : $dcRoot" "INFO"
    if ($modeDediee) {
        Write-Log "Mode : groupes dans OU dédiée '$nomOUDediee'" "INFO"
    } else {
        Write-Log "Mode : groupes dans l'OU de leurs utilisateurs" "INFO"
    }

    # -- ETAPE 1 : OUs
    Write-Log "--- Etape 1 : OUs ---" "INFO"
    foreach ($ou in $script:listeOUs) {
        $ouDN = "OU=$ou,$dcRoot"
        try {
            Get-ADOrganizationalUnit -Identity $ouDN -ErrorAction Stop | Out-Null
            Write-Log "OU existe : $ou" "WARN"
        } catch {
            try {
                New-ADOrganizationalUnit -Name $ou -Path $dcRoot
                Write-Log "OU créée : $ou" "OK"
            } catch { Write-Log "Erreur OU $ou : $_" "ERR" }
        }
    }

    # -- ETAPE 2 : OU dédiée
    if ($modeDediee) {
        Write-Log "--- Etape 2 : OU dédiée groupes ---" "INFO"
        $ouDedieDN = "OU=$nomOUDediee,$dcRoot"
        try {
            Get-ADOrganizationalUnit -Identity $ouDedieDN -ErrorAction Stop | Out-Null
            Write-Log "OU dédiée existe : $nomOUDediee" "WARN"
        } catch {
            try {
                New-ADOrganizationalUnit -Name $nomOUDediee -Path $dcRoot
                Write-Log "OU dédiée créée : $nomOUDediee" "OK"
            } catch { Write-Log "Erreur OU dédiée : $_" "ERR" }
        }
    }

    # -- ETAPE 3 : Groupes
    Write-Log "--- Etape 3 : Groupes ---" "INFO"
    foreach ($entry in $script:listeGroupes) {
        $separateur = " -> "
        $idx    = $entry.IndexOf($separateur)
        $nom    = $entry.Substring(0, $idx).Trim()
        $ouUser = $entry.Substring($idx + $separateur.Length).Trim()
        $ouDN   = if ($modeDediee) { "OU=$nomOUDediee,$dcRoot" } else { "OU=$ouUser,$dcRoot" }
        try {
            Get-ADGroup -Identity $nom -ErrorAction Stop | Out-Null
            Write-Log "Groupe existe : $nom" "WARN"
        } catch {
            try {
                New-ADGroup -Name $nom -SamAccountName $nom -GroupScope Global -GroupCategory Security -Path $ouDN
                Write-Log "Groupe créé : $nom" "OK"
            } catch { Write-Log "Erreur groupe $nom : $_" "ERR" }
        }
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
            Write-Log "Ligne incomplète, ignorée." "WARN"
            continue
        }

        $ouDN      = "OU=$ou,$dcRoot"
        $securePwd = ConvertTo-SecureString $mdp -AsPlainText -Force
        $fullName  = "$prenom $nom"
        $existing  = Get-ADUser -Filter { SamAccountName -eq $login } -ErrorAction SilentlyContinue

        if ($existing) {
            try {
                Set-ADUser -Identity $login -GivenName $prenom -Surname $nom -DisplayName $fullName
                Set-ADAccountPassword -Identity $login -NewPassword $securePwd -Reset
                Write-Log "User mis à jour : $login" "WARN"
            } catch { Write-Log "Erreur MAJ $login : $_" "ERR"; continue }
        } else {
            try {
                New-ADUser -SamAccountName $login `
                           -UserPrincipalName "$login@$((Get-ADDomain).DNSRoot)" `
                           -GivenName $prenom -Surname $nom -DisplayName $fullName -Name $fullName `
                           -AccountPassword $securePwd -Path $ouDN -Enabled $true
                Write-Log "User créé : $login" "OK"
            } catch { Write-Log "Erreur création $login : $_" "ERR"; continue }
        }

        try {
            Add-ADGroupMember -Identity $grp -Members $login -ErrorAction Stop
            Write-Log "  -> $login ajouté à $grp" "OK"
        } catch { Write-Log "  -> Erreur groupe $grp pour $login : $_" "ERR" }
    }

    Write-Log "=== FIN PROVISIONING ===" "INFO"
    Write-Log "Log : $script:logPath" "INFO"

    # Sauvegarde automatique config
    Exporter-Config
    Write-Log "Config sauvegardée : $($script:configPath)" "INFO"
})

# ============================================================
# CHARGEMENT AUTO AU DEMARRAGE
# ============================================================
$form.Add_Shown({
    if (Test-Path $script:configPath) {
        Charger-Config -Chemin $script:configPath
        $lblConfigActif.Text      = "Config auto-chargée : $($script:configPath)"
        $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
    }
})

# ============================================================
$form.ShowDialog() | Out-Null
