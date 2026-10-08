# ============================================================
#  PROVISIONING AD - Logique metier
#  Validation, generation des mots de passe, lecture/ecriture CSV.
#  Ne depend ni de l'Active Directory ni de l'interface : charge par
#  dot-sourcing depuis Provisioning_AD.ps1 et teste par
#  tests\ProvisioningCore.Tests.ps1 (Pester 5).
#  Auteur    : anotherj4ck
# ============================================================

# ============================================================
# VALIDATION DES SAISIES
# ============================================================

function Get-ErreurLogin {
    <#
    .SYNOPSIS
    Retourne la raison pour laquelle un login n'est pas un sAMAccountName valide, ou $null s'il est valide.
    .DESCRIPTION
    Regles Microsoft du sAMAccountName d'un utilisateur : 20 caracteres maximum,
    pas de " / \ [ ] : ; | = , + * ? < > ni de caractere de controle,
    pas uniquement des points ou des espaces.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Login)

    if ([string]::IsNullOrWhiteSpace($Login)) { return "Le login est obligatoire." }
    if ($Login.Length -gt 20) {
        return "Login '$Login' trop long ($($Login.Length) caracteres, maximum 20)."
    }
    if ($Login -match '[\x00-\x1F"/\\\[\]:;|=,+*?<>]') {
        return "Login '$Login' : caractere interdit. Interdits : `" / \ [ ] : ; | = , + * ? < >"
    }
    if ($Login -match '^[. ]+$') {
        return "Login '$Login' : ne peut pas contenir uniquement des points ou des espaces."
    }
    return $null
}

function Test-LoginValide {
    <#
    .SYNOPSIS
    Indique si un login respecte les regles du sAMAccountName (voir Get-ErreurLogin).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Login)

    return ($null -eq (Get-ErreurLogin -Login $Login))
}

function Get-ErreurNomAD {
    <#
    .SYNOPSIS
    Retourne la raison pour laquelle un nom d'OU ou de groupe est invalide, ou $null s'il est valide.
    .DESCRIPTION
    Le nom est insere tel quel dans un Distinguished Name ("OU=<nom>,DC=...") :
    les caracteres speciaux d'un DN (, = + < > # ; \ ") le casseraient.
    Le nom d'un groupe sert aussi de sAMAccountName : on interdit donc en plus
    / [ ] : | * ?. Une seule regle, la plus stricte, pour les OUs et les groupes.
    64 caracteres maximum (limite des attributs ou et cn).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Nom)

    if ([string]::IsNullOrWhiteSpace($Nom)) { return "Le nom est obligatoire." }
    if ($Nom.Length -gt 64) {
        return "Nom '$Nom' trop long ($($Nom.Length) caracteres, maximum 64)."
    }
    if ($Nom -match '[\x00-\x1F,=+<>#;\\"/\[\]:|*?]') {
        return "Nom '$Nom' : caractere interdit. Interdits : , = + < > # ; \ `" / [ ] : | * ?"
    }
    if ($Nom -match '^[. ]+$') {
        return "Nom '$Nom' : ne peut pas contenir uniquement des points ou des espaces."
    }
    return $null
}

function Test-NomADValide {
    <#
    .SYNOPSIS
    Indique si un nom d'OU ou de groupe est utilisable sans casser un DN (voir Get-ErreurNomAD).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][AllowEmptyString()][string]$Nom)

    return ($null -eq (Get-ErreurNomAD -Nom $Nom))
}

function Get-ErreurConfiguration {
    <#
    .SYNOPSIS
    Valide toute la configuration avant provisioning et retourne la liste des erreurs (vide si tout est correct).
    .DESCRIPTION
    Rien n'est ecrit dans l'AD tant qu'une erreur subsiste : mieux vaut tout
    corriger d'un coup que laisser l'annuaire a moitie provisionne.
    Les comparaisons ignorent la casse, comme l'AD.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string[]]$OUs = @(),
        [object[]]$Groupes = @(),
        [object[]]$Utilisateurs = @(),
        [string]$OUDediee = ''
    )

    $erreurs = New-Object System.Collections.Generic.List[string]

    $ousConnues = @{}
    foreach ($ou in $OUs) {
        $e = Get-ErreurNomAD -Nom $ou
        if ($e) { $erreurs.Add("OU : $e") }
        elseif ($ousConnues.ContainsKey($ou)) { $erreurs.Add("OU '$ou' en double.") }
        else { $ousConnues[$ou] = $true }
    }

    if ($OUDediee) {
        $e = Get-ErreurNomAD -Nom $OUDediee
        if ($e) { $erreurs.Add("OU dediee aux groupes : $e") }
    }

    $groupesConnus = @{}
    foreach ($g in $Groupes) {
        $e = Get-ErreurNomAD -Nom $g.Nom
        if ($e) { $erreurs.Add("Groupe : $e"); continue }
        if ($groupesConnus.ContainsKey($g.Nom)) {
            $erreurs.Add("Groupe '$($g.Nom)' en double (un nom de groupe est unique dans le domaine).")
        }
        $groupesConnus[$g.Nom] = $true
        if ([string]::IsNullOrWhiteSpace($g.OU)) {
            $erreurs.Add("Groupe '$($g.Nom)' : OU obligatoire.")
        } elseif (-not $ousConnues.ContainsKey($g.OU)) {
            $erreurs.Add("Groupe '$($g.Nom)' : OU '$($g.OU)' absente de la liste des OUs.")
        }
    }

    $loginsConnus = @{}
    foreach ($u in $Utilisateurs) {
        $e = Get-ErreurLogin -Login $u.Login
        if ($e) { $erreurs.Add("Utilisateur : $e"); continue }
        if ($loginsConnus.ContainsKey($u.Login)) { $erreurs.Add("Login '$($u.Login)' en double.") }
        $loginsConnus[$u.Login] = $true
        if ([string]::IsNullOrWhiteSpace($u.OU)) {
            $erreurs.Add("Utilisateur '$($u.Login)' : OU obligatoire.")
        } elseif (-not $ousConnues.ContainsKey($u.OU)) {
            $erreurs.Add("Utilisateur '$($u.Login)' : OU '$($u.OU)' absente de la liste des OUs.")
        }
        if (-not [string]::IsNullOrWhiteSpace($u.Groupe) -and -not $groupesConnus.ContainsKey($u.Groupe)) {
            $erreurs.Add("Utilisateur '$($u.Login)' : groupe '$($u.Groupe)' absent de la liste des groupes.")
        }
    }

    return $erreurs.ToArray()
}

function ConvertTo-ValeurFiltreLdap {
    <#
    .SYNOPSIS
    Echappe une valeur pour l'inserer dans un filtre LDAP (RFC 4515) : \ * ( ) et NUL.
    .DESCRIPTION
    Evite qu'un nom contenant des parentheses modifie le sens d'un filtre
    comme (sAMAccountName=<valeur>). Le \ est traite en premier pour ne pas
    re-echapper les sequences ajoutees ensuite.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Valeur)

    return ($Valeur -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00')
}

# ============================================================
# MOTS DE PASSE PROVISOIRES
# ============================================================

function Get-EntierAleatoire {
    <#
    .SYNOPSIS
    Retourne un entier uniforme dans [0, Max[ (Max <= 256) a partir d'un generateur cryptographique.
    .DESCRIPTION
    Un simple "octet % Max" favoriserait les premieres valeurs (biais modulo) :
    on rejette les octets au-dela du plus grand multiple de Max.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][System.Security.Cryptography.RandomNumberGenerator]$Generateur,
        [Parameter(Mandatory)][ValidateRange(1, 256)][int]$Max
    )

    $limite = 256 - (256 % $Max)
    $octet = New-Object byte[] 1
    do { $Generateur.GetBytes($octet) } while ($octet[0] -ge $limite)
    return [int]($octet[0] % $Max)
}

function New-MotDePasseProvisoire {
    <#
    .SYNOPSIS
    Genere un mot de passe provisoire aleatoire, conforme a la complexite AD.
    .DESCRIPTION
    - RandomNumberGenerator (cryptographique) et non Get-Random, dont la sortie
      est previsible si l'on connait la graine.
    - Au moins une majuscule, une minuscule, un chiffre et un caractere special,
      places a des positions aleatoires (melange de Fisher-Yates).
    - Sans caracteres ambigus (0 O 1 l I) pour eviter les erreurs de recopie,
      ni guillemets, espaces, ; ou \ qui posent probleme en ligne de commande.
    Le resultat n'existe qu'en memoire : il ne doit etre ni journalise ni ecrit sur disque.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Genere une chaine en memoire, ne modifie aucun etat du systeme.')]
    [CmdletBinding()]
    [OutputType([string])]
    param([ValidateRange(8, 128)][int]$Longueur = 16)

    $categories = @(
        'ABCDEFGHJKLMNPQRSTUVWXYZ'   # sans I ni O
        'abcdefghijkmnopqrstuvwxyz'  # sans l
        '23456789'                   # sans 0 ni 1
        '!#%*+-=?@_'
    )
    $alphabet = -join $categories

    $generateur = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $caracteres = New-Object System.Collections.Generic.List[char]
        foreach ($categorie in $categories) {
            $caracteres.Add($categorie[(Get-EntierAleatoire -Generateur $generateur -Max $categorie.Length)])
        }
        while ($caracteres.Count -lt $Longueur) {
            $caracteres.Add($alphabet[(Get-EntierAleatoire -Generateur $generateur -Max $alphabet.Length)])
        }
        for ($i = $caracteres.Count - 1; $i -gt 0; $i--) {
            $j = Get-EntierAleatoire -Generateur $generateur -Max ($i + 1)
            $tmp = $caracteres[$i]; $caracteres[$i] = $caracteres[$j]; $caracteres[$j] = $tmp
        }
        return (-join $caracteres)
    } finally {
        $generateur.Dispose()
    }
}

# ============================================================
# FICHIER DE CONFIGURATION CSV
# ============================================================
#  Format : separateur ;  encodage UTF-8 (avec BOM, pour Excel)
#  Type;Nom;OU;Login;Prenom;Nom2;Groupe
#  OU;<nom>;;;;;
#  GROUPE;<nom>;<ou>;;;;
#  USER;;<ou>;<login>;<prenom>;<nom>;<groupe>
#  Aucune colonne de mot de passe : aucun secret n'est ecrit sur disque.

function Get-ColonneConfig {
    <#
    .SYNOPSIS
    Retourne les colonnes du CSV de configuration, dans l'ordre.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()

    return [string[]]@('Type', 'Nom', 'OU', 'Login', 'Prenom', 'Nom2', 'Groupe')
}

function Read-FichierTexte {
    <#
    .SYNOPSIS
    Lit un fichier texte en UTF-8 (avec ou sans BOM), ou en ANSI s'il n'est pas de l'UTF-8 valide.
    .DESCRIPTION
    Excel enregistre "CSV (separateur : point-virgule)" en ANSI (Windows-1252) :
    lu comme de l'UTF-8, "Helene" avec accents deviendrait illisible.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Chemin)

    $octets = [System.IO.File]::ReadAllBytes($Chemin)
    $debut = 0
    if ($octets.Length -ge 3 -and $octets[0] -eq 0xEF -and $octets[1] -eq 0xBB -and $octets[2] -eq 0xBF) { $debut = 3 }
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8Strict.GetString($octets, $debut, $octets.Length - $debut)
    } catch [System.Text.DecoderFallbackException] {
        return [System.Text.Encoding]::GetEncoding(1252).GetString($octets)
    }
}

function Import-ConfigProvisioning {
    <#
    .SYNOPSIS
    Lit un CSV de configuration et retourne les OUs, groupes et utilisateurs qu'il decrit.
    .DESCRIPTION
    Retourne un objet :
      OUs                      : noms d'OU
      Groupes                  : objets { Nom ; OU }
      Utilisateurs             : objets { Prenom ; Nom ; Login ; OU ; Groupe }
      ColonneMotDePasseIgnoree : $true si le fichier vient d'une ancienne version
                                 (colonne MotDePasse) : cette colonne n'est jamais lue.
      Avertissements           : lignes ignorees et pourquoi
    L'import ne corrige rien en silence : doublons et references inconnues sont
    signales ensuite par Get-ErreurConfiguration.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Chemin)

    $cheminComplet = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Chemin)
    if (-not (Test-Path -LiteralPath $cheminComplet -PathType Leaf)) { throw "Fichier introuvable : $cheminComplet" }

    $lignes = @((Read-FichierTexte -Chemin $cheminComplet) -split '\r?\n' | Where-Object { $_.Trim() -ne '' })
    if ($lignes.Count -eq 0) { throw "Fichier vide : $cheminComplet" }

    $entetes = @($lignes[0] -split ';' | ForEach-Object { $_.Trim().Trim('"') })
    $manquantes = @(Get-ColonneConfig | Where-Object { $entetes -notcontains $_ })
    if ($manquantes.Count -gt 0) {
        throw "CSV invalide, colonne(s) manquante(s) : $($manquantes -join ', '). Attendu (separateur ;) : $((Get-ColonneConfig) -join ';')"
    }

    $resultat = [PSCustomObject]@{
        OUs                      = New-Object System.Collections.Generic.List[string]
        Groupes                  = New-Object System.Collections.Generic.List[object]
        Utilisateurs             = New-Object System.Collections.Generic.List[object]
        ColonneMotDePasseIgnoree = ($entetes -contains 'MotDePasse')
        Avertissements           = New-Object System.Collections.Generic.List[string]
    }

    # Seules les colonnes connues sont lues : l'eventuelle colonne MotDePasse
    # d'un ancien fichier n'est jamais copiee dans les objets retournes.
    foreach ($ligne in ($lignes | ConvertFrom-Csv -Delimiter ';')) {
        $champ = @{}
        foreach ($colonne in Get-ColonneConfig) { $champ[$colonne] = ([string]$ligne.$colonne).Trim() }
        if (-not ($champ.Values | Where-Object { $_ })) { continue }   # ligne de separateurs seuls

        switch ($champ.Type) {
            'OU' {
                if ($champ.Nom) { $resultat.OUs.Add($champ.Nom) }
                else { $resultat.Avertissements.Add("Ligne OU sans nom ignoree.") }
            }
            'GROUPE' {
                if ($champ.Nom) { $resultat.Groupes.Add([PSCustomObject]@{ Nom = $champ.Nom; OU = $champ.OU }) }
                else { $resultat.Avertissements.Add("Ligne GROUPE sans nom ignoree.") }
            }
            'USER' {
                if ($champ.Login) {
                    $resultat.Utilisateurs.Add([PSCustomObject]@{
                        Prenom = $champ.Prenom
                        Nom    = $champ.Nom2
                        Login  = $champ.Login
                        OU     = $champ.OU
                        Groupe = $champ.Groupe
                    })
                } else {
                    $resultat.Avertissements.Add("Ligne USER sans login ignoree ($($champ.Prenom) $($champ.Nom2)).")
                }
            }
            default {
                $resultat.Avertissements.Add("Ligne ignoree, type '$($champ.Type)' inconnu (attendu : OU, GROUPE ou USER).")
            }
        }
    }

    return $resultat
}

function ConvertTo-ChampCsv {
    <#
    .SYNOPSIS
    Met un champ entre guillemets seulement s'il contient ; " ou un saut de ligne (RFC 4180).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Valeur)

    if ($Valeur -match '[;"\r\n]') { return '"' + ($Valeur -replace '"', '""') + '"' }
    return $Valeur
}

function Export-ConfigProvisioning {
    <#
    .SYNOPSIS
    Ecrit la configuration dans un CSV (separateur ;, UTF-8 avec BOM), sans aucun mot de passe.
    .DESCRIPTION
    Les colonnes ecrites sont une liste fermee (Get-ColonneConfig) : meme si un
    objet utilisateur portait une propriete de mot de passe, elle ne serait pas ecrite.
    Le BOM permet a Excel d'afficher correctement les accents.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Chemin,
        [string[]]$OUs = @(),
        [object[]]$Groupes = @(),
        [object[]]$Utilisateurs = @()
    )

    $lignes = New-Object System.Collections.Generic.List[string]
    $lignes.Add((Get-ColonneConfig) -join ';')

    $ecrire = {
        param($Type, $Nom, $OU, $Login, $Prenom, $Nom2, $Groupe)
        $lignes.Add((@($Type, $Nom, $OU, $Login, $Prenom, $Nom2, $Groupe) | ForEach-Object { ConvertTo-ChampCsv $_ }) -join ';')
    }
    foreach ($ou in $OUs) { & $ecrire 'OU' $ou '' '' '' '' '' }
    foreach ($g in $Groupes) { & $ecrire 'GROUPE' $g.Nom $g.OU '' '' '' '' }
    foreach ($u in $Utilisateurs) { & $ecrire 'USER' '' $u.OU $u.Login $u.Prenom $u.Nom $u.Groupe }

    $cheminComplet = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Chemin)
    [System.IO.File]::WriteAllLines($cheminComplet, $lignes, (New-Object System.Text.UTF8Encoding($true)))
}
