#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# ============================================================
#  Tests Pester 5 de src\ProvisioningCore.ps1
#  Aucune dependance a l'AD ni a l'interface : s'executent partout,
#  y compris en CI (Windows PowerShell 5.1 et PowerShell 7).
#  Lancement : Invoke-Pester -Path .\tests
# ============================================================

BeforeAll {
    . (Join-Path $PSScriptRoot '..\src\ProvisioningCore.ps1')

    # Fichiers ASCII : les accents des donnees de test sont construits par code.
    $script:Helene = "H$([char]0x00E9)l$([char]0x00E8)ne"
}

Describe 'New-MotDePasseProvisoire' {

    It 'fait 16 caracteres par defaut' {
        (New-MotDePasseProvisoire).Length | Should -Be 16
    }

    It 'respecte la longueur demandee' {
        (New-MotDePasseProvisoire -Longueur 24).Length | Should -Be 24
    }

    It 'refuse une longueur inferieure a 8' {
        { New-MotDePasseProvisoire -Longueur 4 } | Should -Throw
    }

    It 'contient toujours majuscule, minuscule, chiffre et caractere special (complexite AD)' {
        foreach ($i in 1..300) {
            $mdp = New-MotDePasseProvisoire
            $mdp | Should -MatchExactly '[A-Z]'
            $mdp | Should -MatchExactly '[a-z]'
            $mdp | Should -Match '[0-9]'
            $mdp | Should -Match '[^A-Za-z0-9]'
        }
    }

    It 'ne contient aucun caractere ambigu (0 O 1 l I)' {
        $echantillon = -join (1..300 | ForEach-Object { New-MotDePasseProvisoire })
        $echantillon | Should -Not -MatchExactly '[0O1lI]'
    }

    It 'ne contient que des caracteres de son alphabet (ni espace, ni guillemet, ni ; ni \)' {
        $echantillon = -join (1..300 | ForEach-Object { New-MotDePasseProvisoire })
        $echantillon | Should -MatchExactly '^[A-HJ-NP-Za-km-z2-9!#%*+\-=?@_]+$'
    }

    It 'donne un resultat different a chaque appel' {
        New-MotDePasseProvisoire | Should -Not -BeExactly (New-MotDePasseProvisoire)
    }

    It 'ne produit aucun doublon sur 1000 tirages' {
        $tirages = 1..1000 | ForEach-Object { New-MotDePasseProvisoire }
        @($tirages | Select-Object -Unique).Count | Should -Be 1000
    }
}

Describe 'Test-LoginValide' {

    It 'accepte <Login>' -ForEach @(
        @{ Login = 'jdupont' }
        @{ Login = 'marie.curie' }
        @{ Login = 'a-b_c' }
        @{ Login = 'abcdefghij0123456789' }
    ) {
        Test-LoginValide -Login $Login | Should -BeTrue
    }

    It 'refuse un login de plus de 20 caracteres' {
        Test-LoginValide -Login ('a' * 21) | Should -BeFalse
        Get-ErreurLogin -Login ('a' * 21) | Should -Match 'maximum 20'
    }

    It 'refuse le caractere <Car>' -ForEach @(
        '"', '/', '\', '[', ']', ':', ';', '|', '=', ',', '+', '*', '?', '<', '>' | ForEach-Object { @{ Car = $_ } }
    ) {
        Test-LoginValide -Login "jean$Car" | Should -BeFalse
    }

    It 'refuse un login vide ou blanc' {
        Test-LoginValide -Login '' | Should -BeFalse
        Test-LoginValide -Login '   ' | Should -BeFalse
        Test-LoginValide -Login $null | Should -BeFalse
    }

    It 'refuse un login fait uniquement de points ou d espaces' {
        Test-LoginValide -Login '...' | Should -BeFalse
        Test-LoginValide -Login '. .' | Should -BeFalse
    }

    It 'refuse un caractere de controle' {
        Test-LoginValide -Login "jean$([char]9)dupont" | Should -BeFalse
    }
}

Describe 'Test-NomADValide' {

    It 'accepte <Nom>' -ForEach @(
        @{ Nom = 'Comptabilite' }
        @{ Nom = 'Ressources Humaines' }
        @{ Nom = 'GG_Informatique-2026' }
        @{ Nom = 'Service (siege)' }
    ) {
        Test-NomADValide -Nom $Nom | Should -BeTrue
    }

    It 'accepte les accents' {
        Test-NomADValide -Nom "Comptabilit$([char]0x00E9)" | Should -BeTrue
    }

    It 'refuse le caractere <Car>, qui casserait un Distinguished Name' -ForEach @(
        ',', '=', '+', '<', '>', '#', ';', '\', '"' | ForEach-Object { @{ Car = $_ } }
    ) {
        Test-NomADValide -Nom "Compta$($Car)Paie" | Should -BeFalse
    }

    It 'refuse le caractere <Car>, interdit dans le sAMAccountName d un groupe' -ForEach @(
        '/', '[', ']', ':', '|', '*', '?' | ForEach-Object { @{ Car = $_ } }
    ) {
        Test-NomADValide -Nom "Compta$($Car)Paie" | Should -BeFalse
    }

    It 'refuse un nom vide ou de plus de 64 caracteres' {
        Test-NomADValide -Nom '' | Should -BeFalse
        Test-NomADValide -Nom ('a' * 65) | Should -BeFalse
        Test-NomADValide -Nom ('a' * 64) | Should -BeTrue
    }
}

Describe 'ConvertTo-ValeurFiltreLdap' {

    It 'echappe \ * ( ) selon la RFC 4515' {
        ConvertTo-ValeurFiltreLdap -Valeur 'a(b)*\c' | Should -BeExactly 'a\28b\29\2a\5cc'
    }

    It 'laisse une valeur ordinaire intacte' {
        ConvertTo-ValeurFiltreLdap -Valeur 'jdupont' | Should -BeExactly 'jdupont'
    }
}

Describe 'Get-ErreurConfiguration' {

    BeforeAll {
        $script:ous     = @('Direction', 'Informatique')
        $script:groupes = @([PSCustomObject]@{ Nom = 'GG_Info'; OU = 'Informatique' })
    }

    It 'ne signale rien pour une configuration correcte' {
        $users = @([PSCustomObject]@{ Prenom = 'Jean'; Nom = 'Dupont'; Login = 'jdupont'; OU = 'Informatique'; Groupe = 'GG_Info' })
        @(Get-ErreurConfiguration -OUs $ous -Groupes $groupes -Utilisateurs $users).Count | Should -Be 0
    }

    It 'accepte un utilisateur sans groupe' {
        $users = @([PSCustomObject]@{ Login = 'jdupont'; OU = 'Direction'; Groupe = '' })
        @(Get-ErreurConfiguration -OUs $ous -Groupes $groupes -Utilisateurs $users).Count | Should -Be 0
    }

    It 'signale un login en double, sans tenir compte de la casse' {
        $users = @(
            [PSCustomObject]@{ Login = 'jdupont'; OU = 'Direction'; Groupe = '' }
            [PSCustomObject]@{ Login = 'JDupont'; OU = 'Direction'; Groupe = '' }
        )
        Get-ErreurConfiguration -OUs $ous -Utilisateurs $users | Should -Match 'en double'
    }

    It 'signale une OU ou un groupe inconnus' {
        $users = @([PSCustomObject]@{ Login = 'jdupont'; OU = 'Inconnue'; Groupe = 'GG_Inconnu' })
        $erreurs = @(Get-ErreurConfiguration -OUs $ous -Groupes $groupes -Utilisateurs $users)
        $erreurs.Count | Should -Be 2
        $erreurs -join "`n" | Should -Match "OU 'Inconnue'"
        $erreurs -join "`n" | Should -Match "groupe 'GG_Inconnu'"
    }

    It 'signale un utilisateur sans OU' {
        $users = @([PSCustomObject]@{ Login = 'jdupont'; OU = ''; Groupe = '' })
        Get-ErreurConfiguration -OUs $ous -Utilisateurs $users | Should -Match 'OU obligatoire'
    }

    It 'signale un groupe place dans une OU inconnue' {
        $g = @([PSCustomObject]@{ Nom = 'GG_X'; OU = 'Inconnue' })
        Get-ErreurConfiguration -OUs $ous -Groupes $g | Should -Match "OU 'Inconnue'"
    }

    It 'signale les noms et logins invalides' {
        $users = @([PSCustomObject]@{ Login = 'jean;dupont'; OU = 'Direction'; Groupe = '' })
        $erreurs = @(Get-ErreurConfiguration -OUs @('Compta,Paie') -Utilisateurs $users -OUDediee 'Grp=1')
        $erreurs.Count | Should -Be 3
    }
}

Describe 'Export-ConfigProvisioning' {

    BeforeAll {
        $script:chemin = Join-Path $TestDrive 'export.csv'
        # Un objet qui porte un mot de passe ne doit pas le faire fuiter dans le fichier.
        $user = [PSCustomObject]@{
            Prenom = $Helene; Nom = 'Durand'; Login = 'hdurand'; OU = 'Direction'; Groupe = 'GG_Dir'
            MotDePasse = 'ValeurSecreteDeTest'
        }
        Export-ConfigProvisioning -Chemin $chemin -OUs @('Direction') `
            -Groupes @([PSCustomObject]@{ Nom = 'GG_Dir'; OU = 'Direction' }) -Utilisateurs @($user)
    }

    It 'ecrit les colonnes Type;Nom;OU;Login;Prenom;Nom2;Groupe, sans colonne de mot de passe' {
        (Get-Content -LiteralPath $chemin -Encoding UTF8 -TotalCount 1) | Should -BeExactly 'Type;Nom;OU;Login;Prenom;Nom2;Groupe'
    }

    It 'n ecrit aucun mot de passe, meme present dans les objets' {
        $contenu = Get-Content -LiteralPath $chemin -Raw -Encoding UTF8
        $contenu | Should -Not -Match 'MotDePasse'
        $contenu | Should -Not -Match 'ValeurSecreteDeTest'
    }

    It 'ecrit en UTF-8 avec BOM (accents lisibles dans Excel)' {
        $octets = [System.IO.File]::ReadAllBytes($chemin)
        $octets[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
    }

    It 'met entre guillemets un champ qui contient un point-virgule' {
        $c = Join-Path $TestDrive 'guillemets.csv'
        $u = [PSCustomObject]@{ Prenom = 'Jean;Paul'; Nom = 'Dupont'; Login = 'jpdupont'; OU = 'Direction'; Groupe = '' }
        Export-ConfigProvisioning -Chemin $c -OUs @('Direction') -Utilisateurs @($u)
        (Import-ConfigProvisioning -Chemin $c).Utilisateurs[0].Prenom | Should -BeExactly 'Jean;Paul'
    }

    It 'est relu a l identique par Import-ConfigProvisioning (accents compris)' {
        $config = Import-ConfigProvisioning -Chemin $chemin
        $config.OUs | Should -Be @('Direction')
        $config.Groupes[0].Nom | Should -BeExactly 'GG_Dir'
        $config.Utilisateurs[0].Prenom | Should -BeExactly $Helene
        $config.Utilisateurs[0].Login | Should -BeExactly 'hdurand'
        $config.ColonneMotDePasseIgnoree | Should -BeFalse
    }
}

Describe 'Import-ConfigProvisioning' {

    It 'detecte et ignore la colonne MotDePasse d un ancien fichier' {
        # Format des versions precedentes : colonne MotDePasse et ; final en trop
        $chemin = Join-Path $TestDrive 'ancien.csv'
        @(
            'Type;Nom;OU;Login;Prenom;Nom2;MotDePasse;Groupe'
            'OU;Direction;;;;;;;'
            'GROUPE;GG_Dir;Direction;;;;;;'
            'USER;;Direction;jdupont;Jean;Dupont;AncienSecretEnClair;GG_Dir'
        ) | Set-Content -LiteralPath $chemin -Encoding UTF8

        $config = Import-ConfigProvisioning -Chemin $chemin
        $config.ColonneMotDePasseIgnoree | Should -BeTrue
        $config.Utilisateurs.Count | Should -Be 1
        $config.Utilisateurs[0].Groupe | Should -BeExactly 'GG_Dir'
        $config.Utilisateurs[0].PSObject.Properties.Name | Should -Not -Contain 'MotDePasse'
        ($config | ConvertTo-Json -Depth 5) | Should -Not -Match 'AncienSecretEnClair'
    }

    It 'lit un CSV enregistre en ANSI par Excel sans corrompre les accents' {
        $chemin = Join-Path $TestDrive 'ansi.csv'
        $texte = "Type;Nom;OU;Login;Prenom;Nom2;Groupe`r`nOU;Direction;;;;;`r`nUSER;;Direction;hdurand;$Helene;Durand;"
        [System.IO.File]::WriteAllBytes($chemin, [System.Text.Encoding]::GetEncoding(1252).GetBytes($texte))
        (Import-ConfigProvisioning -Chemin $chemin).Utilisateurs[0].Prenom | Should -BeExactly $Helene
    }

    It 'signale une ligne de type inconnu et ignore les lignes vides' {
        $chemin = Join-Path $TestDrive 'inconnu.csv'
        @('Type;Nom;OU;Login;Prenom;Nom2;Groupe', 'OU;Direction;;;;;', ';;;;;;', '', 'MACHINE;PC01;;;;;') |
            Set-Content -LiteralPath $chemin -Encoding UTF8
        $config = Import-ConfigProvisioning -Chemin $chemin
        $config.OUs.Count | Should -Be 1
        $config.Avertissements.Count | Should -Be 1
        $config.Avertissements[0] | Should -Match 'MACHINE'
    }

    It 'leve une erreur claire si une colonne obligatoire manque' {
        $chemin = Join-Path $TestDrive 'incomplet.csv'
        @('Type,Nom,OU', 'OU,Direction,') | Set-Content -LiteralPath $chemin -Encoding UTF8
        { Import-ConfigProvisioning -Chemin $chemin } | Should -Throw '*colonne(s) manquante(s)*'
    }

    It 'leve une erreur si le fichier n existe pas' {
        { Import-ConfigProvisioning -Chemin (Join-Path $TestDrive 'absent.csv') } | Should -Throw '*introuvable*'
    }
}

Describe 'Fichiers du depot' {

    It 'config_ad.example.csv est valide et ne contient pas de mot de passe' {
        $config = Import-ConfigProvisioning -Chemin (Join-Path $PSScriptRoot '..\config_ad.example.csv')
        $config.ColonneMotDePasseIgnoree | Should -BeFalse
        $config.Avertissements.Count | Should -Be 0
        $config.Utilisateurs.Count | Should -BeGreaterThan 0
        @(Get-ErreurConfiguration -OUs $config.OUs -Groupes $config.Groupes -Utilisateurs $config.Utilisateurs).Count |
            Should -Be 0
    }

    It 'Provisioning_AD.ps1 ne contient pas d erreur de syntaxe' {
        $erreurs = $null
        $tokens = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $PSScriptRoot '..\Provisioning_AD.ps1'), [ref]$tokens, [ref]$erreurs) | Out-Null
        $erreurs.Count | Should -Be 0
    }
}
