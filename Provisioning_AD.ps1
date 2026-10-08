#Requires -Version 5.1
# ============================================================
#  PROVISIONING AD - GUI Windows Forms v6
#  CSV unique : OUs / Groupes / Utilisateurs
#  Auteur    : anotherj4ck
#  Licence   : MIT
# ============================================================
#  Securite (details dans le README, section "Choix de securite") :
#  - mots de passe provisoires generes en memoire, jamais saisis, jamais
#    ecrits sur disque ni dans le journal, affiches une seule fois ;
#  - changement de mot de passe force a la premiere connexion ;
#  - mode simulation par defaut, saisies validees, recapitulatif et
#    confirmation avant toute ecriture reelle dans l'AD.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Logique sans AD ni interface (validation, mots de passe, CSV), testee par Pester
. "$PSScriptRoot\src\ProvisioningCore.ps1"

# ============================================================
# MODULE ACTIVEDIRECTORY (RSAT)
# ============================================================
# Pas de "#Requires -Modules ActiveDirectory" : il arreterait le script sur
# une erreur console, sans dire comment installer le module.
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Le module PowerShell ActiveDirectory est introuvable.`n`n" +
        "Installez les outils RSAT depuis une console PowerShell lancee en administrateur :`n`n" +
        "Windows 10 / 11 :`n" +
        "Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`n`n" +
        "Windows Server :`n" +
        "Install-WindowsFeature RSAT-AD-PowerShell",
        "Module ActiveDirectory manquant", "OK", "Error") | Out-Null
    return
}
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Impossible de charger le module ActiveDirectory :`n$_", "Module ActiveDirectory", "OK", "Error") | Out-Null
    return
}

# ============================================================
# CONSTANTES & CHEMINS
# ============================================================
$script:SEP        = " -> "
$script:configPath = "$PSScriptRoot\config_ad.csv"
# Journal horodate a cote du script (logs\ est exclu du depot par .gitignore)
$script:logPath    = "$PSScriptRoot\logs\provisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ============================================================
# DONNEES EN MEMOIRE
# ============================================================
$script:listeOUs     = [System.Collections.ArrayList]@()   # noms d'OU
$script:listeGroupes = [System.Collections.ArrayList]@()   # objets { Nom ; OU }

# ============================================================
# FONCTIONS
# ============================================================

function Write-Journal {
    # Journal horodate : zone de texte de l'interface + fichier dans logs\.
    # Ne jamais y passer de mot de passe.
    param([string]$Message, [string]$Level = "INFO")
    $heure = Get-Date
    if ($null -ne $txtLog -and $txtLog.IsHandleCreated) {
        $txtLog.AppendText("[$($heure.ToString('HH:mm:ss'))] [$Level] $Message`r`n")
        $txtLog.ScrollToCaret()
    }
    try {
        $dir = Split-Path $script:logPath
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $script:logPath -Value "[$($heure.ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message" -Encoding UTF8
    } catch {
        if ($null -ne $txtLog -and $txtLog.IsHandleCreated) {
            $txtLog.AppendText("[ERR] Ecriture du journal impossible ($($script:logPath)) : $_`r`n")
        }
    }
}

function Sync-ListeOU {
    # Repercute $script:listeOUs dans la liste, les listes deroulantes et la colonne OU de la grille.
    $lstOU.Items.Clear()
    $cmbOUGroupe.Items.Clear()
    $cmbOUDefaut.Items.Clear()
    $grid.Columns["OU"].Items.Clear()
    foreach ($ou in $script:listeOUs) {
        $lstOU.Items.Add($ou) | Out-Null
        $cmbOUGroupe.Items.Add($ou) | Out-Null
        $cmbOUDefaut.Items.Add($ou) | Out-Null
        $grid.Columns["OU"].Items.Add($ou) | Out-Null
    }
}

function Sync-ListeGroupe {
    # Repercute $script:listeGroupes dans la liste, la liste deroulante et la colonne Groupe.
    # Entree vide en tete : le groupe d'un utilisateur est facultatif.
    $lstGroupe.Items.Clear()
    $cmbGroupeDefaut.Items.Clear()
    $grid.Columns["Groupe"].Items.Clear()
    $cmbGroupeDefaut.Items.Add("") | Out-Null
    $grid.Columns["Groupe"].Items.Add("") | Out-Null
    foreach ($g in $script:listeGroupes) {
        $lstGroupe.Items.Add("$($g.Nom)$($script:SEP)$($g.OU)") | Out-Null
        $cmbGroupeDefaut.Items.Add($g.Nom) | Out-Null
        $grid.Columns["Groupe"].Items.Add($g.Nom) | Out-Null
    }
}

function Clear-ReferenceOrpheline {
    # Vide les cellules de la grille qui pointent vers une OU ou un groupe supprime.
    param([string]$Colonne, [string]$Valeur)
    foreach ($row in $grid.Rows) {
        if ($row.Cells[$Colonne].Value -eq $Valeur) { $row.Cells[$Colonne].Value = $null }
    }
}

function Clear-Saisie {
    # Remet les champs de saisie manuelle sur leur texte d'aide (gris).
    foreach ($ctrl in $pnlSaisie.Controls) {
        if ($ctrl -is [System.Windows.Forms.TextBox]) {
            $ctrl.Text      = $ctrl.Tag
            $ctrl.ForeColor = [System.Drawing.Color]::Gray
        }
    }
}

function Get-UtilisateurGrille {
    # Lignes de la grille -> objets { Prenom ; Nom ; Login ; OU ; Groupe }
    foreach ($row in $grid.Rows) {
        [PSCustomObject]@{
            Prenom = ([string]$row.Cells["Prenom"].Value).Trim()
            Nom    = ([string]$row.Cells["Nom"].Value).Trim()
            Login  = ([string]$row.Cells["Login"].Value).Trim()
            OU     = ([string]$row.Cells["OU"].Value).Trim()
            Groupe = ([string]$row.Cells["Groupe"].Value).Trim()
        }
    }
}

function Save-Config {
    # Sauvegarde OUs, groupes et utilisateurs dans le CSV de config (jamais de mot de passe).
    Export-ConfigProvisioning -Chemin $script:configPath -OUs $script:listeOUs `
        -Groupes $script:listeGroupes -Utilisateurs @(Get-UtilisateurGrille)
}

function Import-Config {
    # Charge un CSV de config dans l'interface. Retourne la config lue, ou $null en cas d'erreur.
    param([string]$Chemin)
    try {
        $config = Import-ConfigProvisioning -Chemin $Chemin
    } catch {
        Write-Journal "Lecture de la config impossible : $_" "ERR"
        [System.Windows.Forms.MessageBox]::Show("Lecture de la config impossible :`n$_", "Erreur", "OK", "Error") | Out-Null
        return $null
    }

    $grid.Rows.Clear()
    $script:listeOUs.Clear()
    $script:listeGroupes.Clear()
    foreach ($ou in $config.OUs)    { $script:listeOUs.Add($ou) | Out-Null }
    foreach ($g in $config.Groupes) { $script:listeGroupes.Add($g) | Out-Null }
    Sync-ListeOU
    Sync-ListeGroupe
    foreach ($u in $config.Utilisateurs) {
        $grid.Rows.Add($u.Prenom, $u.Nom, $u.Login, $u.OU, $u.Groupe) | Out-Null
    }

    Write-Journal ("Config chargee : $Chemin ($($config.OUs.Count) OU, $($config.Groupes.Count) groupe(s), " +
               "$($config.Utilisateurs.Count) utilisateur(s))") "INFO"
    if ($config.ColonneMotDePasseIgnoree) {
        Write-Journal ("Colonne MotDePasse (ancien format) ignoree : les mots de passe sont desormais generes. " +
                   "Ce fichier contient des mots de passe en clair : re-exportez la config pour le reecrire sans eux.") "WARN"
    }
    foreach ($a in $config.Avertissements) { Write-Journal $a "WARN" }
    $erreurs = @(Get-ErreurConfiguration -OUs $script:listeOUs -Groupes $script:listeGroupes -Utilisateurs @(Get-UtilisateurGrille))
    foreach ($e in $erreurs) { Write-Journal "A corriger avant lancement : $e" "WARN" }
    return $config
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
            "Impossible de contacter l'AD :`n$_", "Erreur AD", "OK", "Error") | Out-Null
        return
    }

    $grid.Rows.Clear()
    $script:listeOUs.Clear()
    $script:listeGroupes.Clear()

    # OUs de niveau 1 sous la racine du domaine
    $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $dcRoot -SearchScope OneLevel -ErrorAction SilentlyContinue
    foreach ($ou in $ous) {
        if ($script:listeOUs -notcontains $ou.Name) { $script:listeOUs.Add($ou.Name) | Out-Null }
    }

    # Groupes dont l'OU parente fait partie des OUs importees
    $groupes = Get-ADGroup -Filter * -SearchBase $dcRoot -ErrorAction SilentlyContinue
    foreach ($grp in $groupes) {
        $ouMatch = [regex]::Match($grp.DistinguishedName, '(?<=,OU=)[^,]+')
        if ($ouMatch.Success -and $script:listeOUs -contains $ouMatch.Value) {
            $script:listeGroupes.Add([PSCustomObject]@{ Nom = $grp.Name; OU = $ouMatch.Value }) | Out-Null
        }
    }

    Sync-ListeOU
    Sync-ListeGroupe

    # Utilisateurs des OUs importees (pas de mot de passe : il n'est pas lisible dans l'AD)
    $nomsGroupes = @($script:listeGroupes | ForEach-Object { $_.Nom })
    $users = Get-ADUser -Filter * -SearchBase $dcRoot -Properties GivenName, Surname, MemberOf `
        -ErrorAction SilentlyContinue
    foreach ($user in $users) {
        $ouMatch = [regex]::Match($user.DistinguishedName, '(?<=,OU=)[^,]+')
        if (-not $ouMatch.Success -or $script:listeOUs -notcontains $ouMatch.Value) { continue }

        $grpNom = ""
        foreach ($memberDN in $user.MemberOf) {
            $grpMatch = [regex]::Match($memberDN, '^CN=([^,]+)')
            if ($grpMatch.Success -and $nomsGroupes -contains $grpMatch.Groups[1].Value) {
                $grpNom = $grpMatch.Groups[1].Value
                break
            }
        }
        $grid.Rows.Add($user.GivenName, $user.Surname, $user.SamAccountName, $ouMatch.Value, $grpNom) | Out-Null
    }

    $resume = "$($script:listeOUs.Count) OU(s)   |   $($script:listeGroupes.Count) groupe(s)   |   $($grid.Rows.Count) utilisateur(s)"
    Write-Journal "Lecture AD terminee : $resume" "INFO"
    [System.Windows.Forms.MessageBox]::Show(
        "Lecture AD terminee.`n$resume", "Lecture AD", "OK", "Information") | Out-Null
}

function Invoke-Provisioning {
    <#
    .SYNOPSIS
    Ecrit OUs, groupes et utilisateurs dans l'AD, ou simule l'operation.
    .DESCRIPTION
    En simulation, seules des lectures sont faites (l'objet existe-t-il deja ?)
    et chaque ecriture est seulement journalisee. -WhatIf des cmdlets AD n'est
    pas utilise : sa sortie va a la console et non au journal.
    Les mots de passe provisoires sont ajoutes a -MotsDePasse (memoire seulement)
    des que le compte est cree : si le traitement s'interrompt, l'appelant peut
    quand meme afficher ceux des comptes deja crees.
    Retourne le nombre d'erreurs.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '',
        Justification = 'Mot de passe genere en memoire par New-MotDePasseProvisoire, jamais lu depuis un fichier ni une saisie.')]
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][string]$DomaineDN,
        [Parameter(Mandatory)][string]$DomaineDNS,
        [string[]]$OUs = @(),
        [object[]]$Groupes = @(),
        [object[]]$Utilisateurs = @(),
        [string]$OUDediee = "",
        [bool]$Simulation = $true,
        [bool]$ReinitialiserMdpExistants = $false,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$MotsDePasse,
        [scriptblock]$Progression = {}
    )

    $erreurs   = 0
    $etape     = 0
    $ousACreer = @($OUs)
    if ($OUDediee -and $ousACreer -notcontains $OUDediee) { $ousACreer += $OUDediee }
    $total = [Math]::Max(1, $ousACreer.Count + $Groupes.Count + $Utilisateurs.Count)

    # -- ETAPE 1 : OUs (dont l'OU dediee aux groupes)
    Write-Journal "--- Etape 1 : OUs ---" "INFO"
    foreach ($ou in $ousACreer) {
        try {
            $existe = Get-ADOrganizationalUnit -LDAPFilter "(ou=$(ConvertTo-ValeurFiltreLdap $ou))" `
                -SearchBase $DomaineDN -SearchScope OneLevel -ErrorAction Stop
            if ($existe) {
                Write-Journal "OU existe deja : $ou" "INFO"
            } elseif ($Simulation) {
                Write-Journal "[SIMULATION] OU serait creee : $ou" "SIMU"
            } else {
                New-ADOrganizationalUnit -Name $ou -Path $DomaineDN -ErrorAction Stop
                Write-Journal "OU creee : $ou" "OK"
            }
        } catch { $erreurs++; Write-Journal "Erreur OU $ou : $_" "ERR" }
        $etape++; & $Progression $etape $total
    }

    # -- ETAPE 2 : Groupes
    Write-Journal "--- Etape 2 : Groupes ---" "INFO"
    foreach ($g in $Groupes) {
        $ouCible = if ($OUDediee) { $OUDediee } else { $g.OU }
        try {
            $existe = Get-ADGroup -LDAPFilter "(sAMAccountName=$(ConvertTo-ValeurFiltreLdap $g.Nom))" -ErrorAction Stop
            if ($existe) {
                Write-Journal "Groupe existe deja : $($g.Nom)" "INFO"
            } elseif ($Simulation) {
                Write-Journal "[SIMULATION] Groupe serait cree : $($g.Nom) (OU $ouCible)" "SIMU"
            } else {
                New-ADGroup -Name $g.Nom -SamAccountName $g.Nom -GroupScope Global -GroupCategory Security `
                    -Path "OU=$ouCible,$DomaineDN" -ErrorAction Stop
                Write-Journal "Groupe cree : $($g.Nom) (OU $ouCible)" "OK"
            }
        } catch { $erreurs++; Write-Journal "Erreur groupe $($g.Nom) : $_" "ERR" }
        $etape++; & $Progression $etape $total
    }

    # -- ETAPE 3 : Utilisateurs
    Write-Journal "--- Etape 3 : Utilisateurs ---" "INFO"
    foreach ($u in $Utilisateurs) {
        $login      = $u.Login
        $ouDN       = "OU=$($u.OU),$DomaineDN"
        $nomComplet = "$($u.Prenom) $($u.Nom)".Trim()
        if (-not $nomComplet) { $nomComplet = $login }

        try {
            $existant = Get-ADUser -LDAPFilter "(sAMAccountName=$(ConvertTo-ValeurFiltreLdap $login))" -ErrorAction Stop
            if ($existant) {
                # Le GUID ne change pas quand le compte est deplace, contrairement au DN.
                $id = $existant.ObjectGUID

                # OU parente = DN sans son premier RDN (une virgule echappee "\," fait partie du RDN)
                $parent = $existant.DistinguishedName -replace '^(?:\\.|[^,\\])*,', ''
                if ($parent -ne $ouDN) {
                    if ($Simulation) {
                        Write-Journal "[SIMULATION] $login serait deplace vers $($u.OU)" "SIMU"
                    } else {
                        Move-ADObject -Identity $id -TargetPath $ouDN -ErrorAction Stop
                        Write-Journal "  -> $login deplace vers $($u.OU)" "OK"
                    }
                }

                # Une cellule vide n'efface pas l'attribut deja present dans l'AD.
                $attributs = @{}
                if ($u.Prenom) { $attributs.GivenName = $u.Prenom }
                if ($u.Nom)    { $attributs.Surname   = $u.Nom }
                if ($u.Prenom -or $u.Nom) { $attributs.DisplayName = $nomComplet }
                if ($Simulation) {
                    Write-Journal "[SIMULATION] Utilisateur existant, serait mis a jour : $login" "SIMU"
                } elseif ($attributs.Count -gt 0) {
                    Set-ADUser -Identity $id @attributs -ErrorAction Stop
                    Write-Journal "Utilisateur existant mis a jour : $login" "OK"
                } else {
                    Write-Journal "Utilisateur existant, rien a mettre a jour : $login" "INFO"
                }

                # Le mot de passe d'un compte existant n'est jamais touche sans la case explicite.
                if ($ReinitialiserMdpExistants) {
                    if ($Simulation) {
                        Write-Journal "[SIMULATION] Mot de passe de $login serait reinitialise (changement force)" "SIMU"
                    } else {
                        try {
                            $mdp = New-MotDePasseProvisoire
                            Set-ADAccountPassword -Identity $id -Reset -ErrorAction Stop `
                                -NewPassword (ConvertTo-SecureString $mdp -AsPlainText -Force)
                            # Ajoute avant l'etape suivante : si elle echoue, l'admin doit
                            # quand meme connaitre le nouveau mot de passe.
                            $MotsDePasse.Add([PSCustomObject]@{ Login = $login; MotDePasse = $mdp })
                            Set-ADUser -Identity $id -ChangePasswordAtLogon $true -ErrorAction Stop
                            Write-Journal "Mot de passe reinitialise : $login (changement force a la prochaine connexion)" "WARN"
                        } catch { $erreurs++; Write-Journal "Erreur reinitialisation du mot de passe de $login : $_" "ERR" }
                    }
                }
            } else {
                if ($Simulation) {
                    Write-Journal "[SIMULATION] Utilisateur serait cree : $login (OU $($u.OU), mot de passe provisoire, changement force)" "SIMU"
                } else {
                    $mdp = New-MotDePasseProvisoire
                    $params = @{
                        SamAccountName        = $login
                        UserPrincipalName     = "$login@$DomaineDNS"
                        Name                  = $nomComplet
                        DisplayName           = $nomComplet
                        Path                  = $ouDN
                        AccountPassword       = (ConvertTo-SecureString $mdp -AsPlainText -Force)
                        ChangePasswordAtLogon = $true
                        Enabled               = $true
                    }
                    if ($u.Prenom) { $params.GivenName = $u.Prenom }
                    if ($u.Nom)    { $params.Surname   = $u.Nom }
                    New-ADUser @params -ErrorAction Stop
                    $MotsDePasse.Add([PSCustomObject]@{ Login = $login; MotDePasse = $mdp })
                    Write-Journal "Utilisateur cree : $login (changement de mot de passe force a la premiere connexion)" "OK"
                }
            }
        } catch {
            $erreurs++
            Write-Journal "Erreur utilisateur $login : $_" "ERR"
            $etape++; & $Progression $etape $total
            continue
        }

        if ($u.Groupe) {
            if ($Simulation) {
                Write-Journal "[SIMULATION] $login serait ajoute au groupe $($u.Groupe)" "SIMU"
            } else {
                try {
                    Add-ADGroupMember -Identity $u.Groupe -Members $login -ErrorAction Stop
                    Write-Journal "  -> $login ajoute a $($u.Groupe)" "OK"
                } catch { $erreurs++; Write-Journal "  -> Erreur groupe $($u.Groupe) pour $login : $_" "ERR" }
            }
        }
        $etape++; & $Progression $etape $total
    }

    return $erreurs
}

function Copy-TexteSensible {
    # Copie dans le presse-papiers en demandant a Windows de ne garder le contenu ni
    # dans l'historique (Win+V) ni dans le presse-papiers cloud : formats documentes
    # par Microsoft et utilises par les gestionnaires de mots de passe. La valeur
    # attendue (DWORD 0) est fournie sous forme de flux d'octets.
    param([Parameter(Mandatory)][string]$Texte)
    $donnees = New-Object System.Windows.Forms.DataObject
    $donnees.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $Texte)
    foreach ($format in @("ExcludeClipboardContentFromMonitorProcessing", "CanIncludeInClipboardHistory", "CanUploadToCloudClipboard")) {
        $donnees.SetData($format, [System.IO.MemoryStream]::new([byte[]](0, 0, 0, 0)))
    }
    [System.Windows.Forms.Clipboard]::SetDataObject($donnees, $true, 5, 100)
}

function Show-MotDePasseProvisoire {
    # Affiche UNE seule fois les mots de passe provisoires : ils ne sont ecrits
    # ni sur disque ni dans le journal, et sont oublies a la fermeture.
    param([Parameter(Mandatory)][object[]]$MotsDePasse)

    $texte = (@("Login`tMot de passe provisoire") +
              @($MotsDePasse | ForEach-Object { "$($_.Login)`t$($_.MotDePasse)" })) -join "`r`n"

    $fen                 = New-Object System.Windows.Forms.Form
    $fen.Text            = "Mots de passe provisoires ($($MotsDePasse.Count))"
    $fen.Size            = New-Object System.Drawing.Size(560, 430)
    $fen.StartPosition   = "CenterParent"
    $fen.FormBorderStyle = "FixedDialog"
    $fen.MaximizeBox     = $false
    $fen.MinimizeBox     = $false
    $fen.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblAvert           = New-Object System.Windows.Forms.Label
    $lblAvert.Text      = "Affichage unique : ces mots de passe ne sont ecrits ni sur disque ni dans le journal.`n" +
                          "Transmettez-les par un canal sur ; chaque utilisateur devra changer le sien a la premiere connexion.`n" +
                          "Le presse-papiers est vide a la fermeture de cette fenetre."
    $lblAvert.Location  = New-Object System.Drawing.Point(15, 12)
    $lblAvert.Size      = New-Object System.Drawing.Size(515, 70)
    $lblAvert.ForeColor = [System.Drawing.Color]::DarkRed
    $fen.Controls.Add($lblAvert)

    $txtMdp            = New-Object System.Windows.Forms.TextBox
    $txtMdp.Multiline  = $true
    $txtMdp.ReadOnly   = $true
    $txtMdp.WordWrap   = $false
    $txtMdp.ScrollBars = "Vertical"
    $txtMdp.Font       = New-Object System.Drawing.Font("Consolas", 10)
    $txtMdp.Location   = New-Object System.Drawing.Point(15, 88)
    $txtMdp.Size       = New-Object System.Drawing.Size(515, 242)
    $txtMdp.Text       = $texte
    $fen.Controls.Add($txtMdp)

    $btnCopier          = New-Object System.Windows.Forms.Button
    $btnCopier.Text     = "Copier"
    $btnCopier.Location = New-Object System.Drawing.Point(15, 342)
    $btnCopier.Size     = New-Object System.Drawing.Size(100, 30)
    $fen.Controls.Add($btnCopier)

    $lblCopie           = New-Object System.Windows.Forms.Label
    $lblCopie.Location  = New-Object System.Drawing.Point(125, 349)
    $lblCopie.Size      = New-Object System.Drawing.Size(295, 18)
    $lblCopie.ForeColor = [System.Drawing.Color]::Gray
    $fen.Controls.Add($lblCopie)

    $btnFermer          = New-Object System.Windows.Forms.Button
    $btnFermer.Text     = "Fermer"
    $btnFermer.Location = New-Object System.Drawing.Point(430, 342)
    $btnFermer.Size     = New-Object System.Drawing.Size(100, 30)
    $fen.Controls.Add($btnFermer)
    $fen.ActiveControl = $btnCopier   # sinon tout le texte apparait selectionne a l'ouverture

    $btnCopier.Add_Click({
        try {
            Copy-TexteSensible -Texte $texte
            $lblCopie.Text = "Copie dans le presse-papiers."
        } catch {
            $lblCopie.Text = "Copie impossible : $_"
        }
    })
    $btnFermer.Add_Click({ $fen.Close() })
    $fen.Add_FormClosing({
        # $_ : arguments de l'evenement (FormClosingEventArgs)
        $rep = [System.Windows.Forms.MessageBox]::Show(
            "Ces mots de passe ne pourront plus etre affiches.`nLes avez-vous copies ou transmis ?`n`nFermer la fenetre ?",
            "Mots de passe provisoires", "YesNo", "Warning", "Button2")
        if ($rep -ne "Yes") { $_.Cancel = $true }
    })

    $fen.ShowDialog() | Out-Null

    # Nettoyage : on vide le presse-papiers s'il contient encore ces mots de passe.
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText() -and [System.Windows.Forms.Clipboard]::GetText() -eq $texte) {
            [System.Windows.Forms.Clipboard]::Clear()
        }
    } catch {
        Write-Journal "Presse-papiers inaccessible, videz-le manuellement : $_" "WARN"
    }
    $txtMdp.Clear()
    $fen.Dispose()
}

# ============================================================
# FENETRE PRINCIPALE
# ============================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Provisioning AD v6"
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
    $err = Get-ErreurNomAD -Nom $val
    if ($err) {
        [System.Windows.Forms.MessageBox]::Show($err, "Nom d'OU invalide", "OK", "Warning") | Out-Null
        return
    }
    if ($script:listeOUs -contains $val) {
        [System.Windows.Forms.MessageBox]::Show("L'OU '$val' est deja dans la liste.", "Attention") | Out-Null
        return
    }
    $script:listeOUs.Add($val) | Out-Null
    Sync-ListeOU
    $txtOU.Clear()
})

$btnSupprimerOU.Add_Click({
    if ($lstOU.SelectedItem) {
        $val = [string]$lstOU.SelectedItem
        $script:listeOUs.Remove($val)
        Sync-ListeOU
        Clear-ReferenceOrpheline -Colonne "OU" -Valeur $val
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
    $ou  = [string]$cmbOUGroupe.SelectedItem
    if (-not $ou) {
        [System.Windows.Forms.MessageBox]::Show("Choisis l'OU cible du groupe.", "Attention") | Out-Null
        return
    }
    $err = Get-ErreurNomAD -Nom $nom
    if ($err) {
        [System.Windows.Forms.MessageBox]::Show($err, "Nom de groupe invalide", "OK", "Warning") | Out-Null
        return
    }
    if (@($script:listeGroupes | Where-Object { $_.Nom -eq $nom }).Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Le groupe '$nom' est deja dans la liste (un nom de groupe est unique dans le domaine).", "Attention") | Out-Null
        return
    }
    $script:listeGroupes.Add([PSCustomObject]@{ Nom = $nom; OU = $ou }) | Out-Null
    Sync-ListeGroupe
    $txtGroupe.Clear()
})

$btnSupprimerGroupe.Add_Click({
    $i = $lstGroupe.SelectedIndex
    if ($i -ge 0) {
        $nom = $script:listeGroupes[$i].Nom
        $script:listeGroupes.RemoveAt($i)
        Sync-ListeGroupe
        Clear-ReferenceOrpheline -Colonne "Groupe" -Valeur $nom
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

# Panneau saisie manuelle (aucun champ mot de passe : il est genere a la creation)
$pnlSaisie          = New-Object System.Windows.Forms.Panel
$pnlSaisie.Location = New-Object System.Drawing.Point(15, 45)
$pnlSaisie.Size     = New-Object System.Drawing.Size(735, 35)
$pnlSaisie.Visible  = $false
$tabUser.Controls.Add($pnlSaisie)

foreach ($info in @(
    @{Name="txtSaisiePrenom"; PlaceHolder="Prenom"; X=0},
    @{Name="txtSaisieNom";    PlaceHolder="Nom";    X=110},
    @{Name="txtSaisieLogin";  PlaceHolder="Login";  X=220}
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
        }
    })
    $t.Add_Leave({
        if ($this.Text -eq "") { $this.ForeColor = [System.Drawing.Color]::Gray; $this.Text = $this.Tag }
    })
    $pnlSaisie.Controls.Add($t)
}

$btnConfirmerUser          = New-Object System.Windows.Forms.Button
$btnConfirmerUser.Text     = "OK"
$btnConfirmerUser.Location = New-Object System.Drawing.Point(330, 3)
$btnConfirmerUser.Size     = New-Object System.Drawing.Size(50, 26)
$pnlSaisie.Controls.Add($btnConfirmerUser)

$btnAnnulerUser          = New-Object System.Windows.Forms.Button
$btnAnnulerUser.Text     = "Annuler"
$btnAnnulerUser.Location = New-Object System.Drawing.Point(388, 3)
$btnAnnulerUser.Size     = New-Object System.Drawing.Size(65, 26)
$pnlSaisie.Controls.Add($btnAnnulerUser)

$btnAjouterUser.Add_Click({ $pnlSaisie.Visible = $true })

$btnAnnulerUser.Add_Click({
    Clear-Saisie
    $pnlSaisie.Visible = $false
})

$btnConfirmerUser.Add_Click({
    # Un champ gris affiche encore son texte d'aide : il est vide.
    $champs = @{}
    foreach ($nom in @("Prenom", "Nom", "Login")) {
        $ctrl = $pnlSaisie.Controls["txtSaisie$nom"]
        $champs[$nom] = if ($ctrl.ForeColor -eq [System.Drawing.Color]::Gray) { "" } else { $ctrl.Text.Trim() }
    }

    $err = Get-ErreurLogin -Login $champs.Login
    if ($err) {
        [System.Windows.Forms.MessageBox]::Show($err, "Login invalide", "OK", "Warning") | Out-Null
        return
    }
    if (@(Get-UtilisateurGrille | Where-Object { $_.Login -eq $champs.Login }).Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Le login '$($champs.Login)' est deja dans la liste.", "Attention") | Out-Null
        return
    }

    $ou  = if ($cmbOUDefaut.SelectedItem)     { [string]$cmbOUDefaut.SelectedItem }     else { "" }
    $grp = if ($cmbGroupeDefaut.SelectedItem) { [string]$cmbGroupeDefaut.SelectedItem } else { "" }
    $grid.Rows.Add($champs.Prenom, $champs.Nom, $champs.Login, $ou, $grp) | Out-Null
    Clear-Saisie
    $pnlSaisie.Visible = $false
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

foreach ($col in @("Prenom","Nom","Login")) {
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

# Supprime le dialog d'erreur DataGridView pour les combobox orphelines
$grid.Add_DataError({ $_.Cancel = $true })

# ============================================================
# BARRE CONFIG CSV
# ============================================================
$grpBoxConfig          = New-Object System.Windows.Forms.GroupBox
$grpBoxConfig.Text     = "Fichier de configuration (config_ad.csv, sans mot de passe)"
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
    try {
        Save-Config
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Export impossible :`n$_", "Erreur", "OK", "Error") | Out-Null
        return
    }
    Write-Journal "Config exportee (sans mot de passe) : $($script:configPath)" "INFO"
    $lblConfigActif.Text      = "Config sauvegardee : $($script:configPath)"
    $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
    [System.Windows.Forms.MessageBox]::Show("Config exportee :`n$($script:configPath)", "Export OK") | Out-Null
})

$btnChargerConfig.Add_Click({
    $ofd          = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter   = "CSV|*.csv"
    $ofd.FileName = "config_ad.csv"
    if ($ofd.ShowDialog() -eq "OK") {
        $config = Import-Config -Chemin $ofd.FileName
        if ($config) {
            $script:configPath        = $ofd.FileName
            $lblConfigActif.Text      = "Config : $($script:configPath)"
            $lblConfigActif.ForeColor = [System.Drawing.Color]::Gray
            $message = "Config chargee :`n$($script:configPath)"
            if ($config.ColonneMotDePasseIgnoree) {
                $message += "`n`nAncienne colonne MotDePasse ignoree : les mots de passe sont desormais generes. " +
                            "Re-exportez la config pour supprimer ces mots de passe en clair du fichier."
            }
            [System.Windows.Forms.MessageBox]::Show($message, "Chargement OK") | Out-Null
        }
    }
})

$btnOuvrirConfig.Add_Click({
    if (Test-Path $script:configPath) {
        Start-Process notepad.exe $script:configPath
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Aucune config trouvee.`nExporte d'abord une config.", "Introuvable") | Out-Null
    }
})

# ============================================================
# OPTION PLACEMENT DES GROUPES
# ============================================================
$grpBoxOption          = New-Object System.Windows.Forms.GroupBox
$grpBoxOption.Text     = "Placement des groupes dans l'AD"
$grpBoxOption.Location = New-Object System.Drawing.Point(10, 615)
$grpBoxOption.Size     = New-Object System.Drawing.Size(385, 80)
$grpBoxOption.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxOption)

$radioAvecOU          = New-Object System.Windows.Forms.RadioButton
$radioAvecOU.Text     = "Dans l'OU des utilisateurs"
$radioAvecOU.Location = New-Object System.Drawing.Point(15, 22)
$radioAvecOU.Size     = New-Object System.Drawing.Size(250, 22)
$radioAvecOU.Checked  = $true
$grpBoxOption.Controls.Add($radioAvecOU)

$radioDediee          = New-Object System.Windows.Forms.RadioButton
$radioDediee.Text     = "Dans une OU dediee (creee si absente) :"
$radioDediee.Location = New-Object System.Drawing.Point(15, 48)
$radioDediee.Size     = New-Object System.Drawing.Size(245, 22)
$grpBoxOption.Controls.Add($radioDediee)

$txtOUDediee          = New-Object System.Windows.Forms.TextBox
$txtOUDediee.Text     = "Groupes"
$txtOUDediee.Location = New-Object System.Drawing.Point(262, 47)
$txtOUDediee.Size     = New-Object System.Drawing.Size(110, 22)
$txtOUDediee.Enabled  = $false
$grpBoxOption.Controls.Add($txtOUDediee)

$radioDediee.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })
$radioAvecOU.Add_CheckedChanged({ $txtOUDediee.Enabled = $radioDediee.Checked })

# ============================================================
# OPTIONS D'EXECUTION (SECURITE)
# ============================================================
$grpBoxExecution          = New-Object System.Windows.Forms.GroupBox
$grpBoxExecution.Text     = "Execution"
$grpBoxExecution.Location = New-Object System.Drawing.Point(405, 615)
$grpBoxExecution.Size     = New-Object System.Drawing.Size(385, 80)
$grpBoxExecution.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($grpBoxExecution)

# Coche par defaut : un clic sur "Lancer" ne doit jamais ecrire dans l'AD par surprise.
$chkSimulation          = New-Object System.Windows.Forms.CheckBox
$chkSimulation.Text     = "Mode simulation (aucune ecriture dans l'AD)"
$chkSimulation.Location = New-Object System.Drawing.Point(15, 22)
$chkSimulation.Size     = New-Object System.Drawing.Size(360, 22)
$chkSimulation.Checked  = $true
$grpBoxExecution.Controls.Add($chkSimulation)

# Decochee par defaut : le mot de passe d'un compte existant n'est jamais change en silence.
$chkResetMdp          = New-Object System.Windows.Forms.CheckBox
$chkResetMdp.Text     = "Reinitialiser les mots de passe des comptes existants"
$chkResetMdp.Location = New-Object System.Drawing.Point(15, 48)
$chkResetMdp.Size     = New-Object System.Drawing.Size(360, 22)
$chkResetMdp.Checked  = $false
$grpBoxExecution.Controls.Add($chkResetMdp)

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
$btnLancer.Text         = "Lancer la simulation"
$btnLancer.Location     = New-Object System.Drawing.Point(10, 727)
$btnLancer.Size         = New-Object System.Drawing.Size(200, 34)
$btnLancer.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnLancer.ForeColor    = [System.Drawing.Color]::White
$btnLancer.FlatStyle    = "Flat"
$btnLancer.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnLancer)

# Le bouton change de texte et de couleur quand l'ecriture reelle est activee.
$chkSimulation.Add_CheckedChanged({
    if ($chkSimulation.Checked) {
        $btnLancer.Text      = "Lancer la simulation"
        $btnLancer.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    } else {
        $btnLancer.Text      = "Lancer le provisioning"
        $btnLancer.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
    }
})

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
    $simulation   = $chkSimulation.Checked
    $reset        = $chkResetMdp.Checked
    $groupes      = @($script:listeGroupes)
    $utilisateurs = @(Get-UtilisateurGrille)
    $ouDediee     = if ($radioDediee.Checked) { $txtOUDediee.Text.Trim() } else { "" }

    if ($radioDediee.Checked -and -not $ouDediee) {
        [System.Windows.Forms.MessageBox]::Show("Indique le nom de l'OU dediee aux groupes.", "Attention") | Out-Null
        return
    }
    if (($script:listeOUs.Count + $groupes.Count + $utilisateurs.Count) -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Rien a provisionner : ajoute des OUs, groupes ou utilisateurs, ou charge une config.", "Attention") | Out-Null
        return
    }

    # 1. Validation complete, avant tout contact avec l'AD
    $erreurs = @(Get-ErreurConfiguration -OUs $script:listeOUs -Groupes $groupes -Utilisateurs $utilisateurs -OUDediee $ouDediee)
    if ($erreurs.Count -gt 0) {
        foreach ($e in $erreurs) { Write-Journal $e "ERR" }
        $apercu = ($erreurs | Select-Object -First 12) -join "`n"
        if ($erreurs.Count -gt 12) { $apercu += "`n... (liste complete dans le journal)" }
        [System.Windows.Forms.MessageBox]::Show(
            "Configuration invalide : $($erreurs.Count) erreur(s). Rien n'a ete lance.`n`n$apercu",
            "Validation", "OK", "Warning") | Out-Null
        return
    }

    try { $domaine = Get-ADDomain -ErrorAction Stop }
    catch {
        Write-Journal "Impossible de contacter l'AD : $_" "ERR"
        [System.Windows.Forms.MessageBox]::Show("Impossible de contacter l'AD :`n$_", "Erreur AD", "OK", "Error") | Out-Null
        return
    }

    $txtReset = if ($reset) { "OUI" } else { "non" }
    $recap    = "$($script:listeOUs.Count) OU(s), $($groupes.Count) groupe(s), $($utilisateurs.Count) utilisateur(s)"
    if ($ouDediee) { $recap += ", groupes dans l'OU dediee '$ouDediee'" }

    # 2. Recapitulatif et confirmation avant toute ecriture reelle ("Non" par defaut)
    if (-not $simulation) {
        $message = "ECRITURE REELLE dans le domaine $($domaine.DNSRoot)`n`n" +
                   "  - $($script:listeOUs.Count) OU(s)`n" +
                   "  - $($groupes.Count) groupe(s)`n" +
                   "  - $($utilisateurs.Count) utilisateur(s) (crees, ou mis a jour s'ils existent)`n"
        if ($ouDediee) { $message += "  - groupes places dans l'OU dediee '$ouDediee'`n" }
        $message += "`nReinitialisation des mots de passe des comptes existants : $txtReset`n`n" +
                    "Confirmer l'ecriture dans l'AD ?"
        $rep = [System.Windows.Forms.MessageBox]::Show($message, "Confirmation", "YesNo", "Warning", "Button2")
        if ($rep -ne "Yes") {
            Write-Journal "Provisioning annule a l'etape de confirmation." "INFO"
            return
        }
    }

    # 3. Execution : interface verrouillee, mots de passe gardes en memoire seulement
    $controles = @($tabControl, $grpBoxConfig, $grpBoxOption, $grpBoxExecution, $btnLancer)
    foreach ($ctrl in $controles) { $ctrl.Enabled = $false }
    $progressBar.Value = 0
    $motsDePasse = New-Object System.Collections.Generic.List[object]
    $nbErreurs   = 0
    try {
        if ($simulation) { Write-Journal "=== DEBUT SIMULATION (aucune ecriture dans l'AD) ===" "INFO" }
        else             { Write-Journal "=== DEBUT PROVISIONING ===" "INFO" }
        Write-Journal "Domaine : $($domaine.DNSRoot) | Operateur : $env:USERDOMAIN\$env:USERNAME" "INFO"
        Write-Journal "Recapitulatif : $recap | Reinitialisation des mots de passe existants : $txtReset" "INFO"

        $nbErreurs = Invoke-Provisioning -DomaineDN $domaine.DistinguishedName -DomaineDNS $domaine.DNSRoot `
            -OUs $script:listeOUs -Groupes $groupes -Utilisateurs $utilisateurs -OUDediee $ouDediee `
            -Simulation $simulation -ReinitialiserMdpExistants $reset -MotsDePasse $motsDePasse `
            -Progression {
                param($etape, $total)
                $progressBar.Value = [int][Math]::Min(100, $etape * 100 / $total)
                [System.Windows.Forms.Application]::DoEvents()
            }
        Write-Journal "=== FIN : $nbErreurs erreur(s) ===" "INFO"

        try {
            Save-Config
            Write-Journal "Config sauvegardee (sans mot de passe) : $($script:configPath)" "INFO"
        } catch { Write-Journal "Sauvegarde de la config impossible : $_" "WARN" }
    } catch {
        $nbErreurs++
        Write-Journal "Arret inattendu : $_" "ERR"
    } finally {
        $progressBar.Value = 100
        foreach ($ctrl in $controles) { $ctrl.Enabled = $true }
        # Affichage unique, puis oubli : la liste n'a jamais existe qu'en memoire.
        if ($motsDePasse.Count -gt 0) { Show-MotDePasseProvisoire -MotsDePasse $motsDePasse.ToArray() }
        $motsDePasse.Clear()
    }

    $bilan = if ($simulation) { "Simulation terminee : aucune modification dans l'AD." } else { "Provisioning termine." }
    [System.Windows.Forms.MessageBox]::Show(
        "$bilan`n$nbErreurs erreur(s).`n`nJournal : $($script:logPath)", "Termine", "OK", "Information") | Out-Null
})

# ============================================================
# CHARGEMENT AUTO AU DEMARRAGE
# ============================================================
$form.Add_Shown({
    Write-Journal "Journal : $($script:logPath)" "INFO"
    Write-Journal "Mode simulation actif : decoche-le pour ecrire dans l'AD." "INFO"
    if (Test-Path $script:configPath) {
        if (Import-Config -Chemin $script:configPath) {
            $lblConfigActif.Text      = "Config auto-chargee : $($script:configPath)"
            $lblConfigActif.ForeColor = [System.Drawing.Color]::DarkGreen
        }
    }
})

# ============================================================
$form.ShowDialog() | Out-Null
$form.Dispose()
