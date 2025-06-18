<#

.SYNOPSIS
    This script defines functions for managing encrypted passwords and retrieving secure credentials.

    .DESCRIPTION

    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    The script contains two functions:
    - New-MZZEncryptedPassword: Prompts the user to enter a password, hashes it, and stores the encrypted password in a file.
    - Get-MZZSecureCreds: Retrieves the stored encrypted password or credentials based on the provided parameters.

    .PARAMETER FileName
    Specifies the name of the file to store the encrypted password. If not provided, the stored password will be used to create a PS credential object.

    .PARAMETER AdminUser
    Specifies the username for which the credentials are stored. If not provided, $env:USERNAME will be used.

    .PARAMETER checkPassword
    Indicates whether to display the decrypted password after retrieving the credentials.

    .PARAMETER updatePassword
    Specifies whether to update the stored password or credentials.

    .NOTES
    - This script requires Git to be installed and accessible in the environment for retrieving the repository path.

    .EXAMPLE
    New-MZZEncryptedPassword -FileName "MyPassword"

    This example prompts the user to enter a password and stores the encrypted password in the "MyPassword.txt" file.

    .EXAMPLE
    Get-MZZSecureCreds -FileName "MyPassword" -CheckPassword

    This example retrieves the encrypted password from the "MyPassword.txt" file, decrypts it, and displays the decrypted password.

    .EXAMPLE
    New-MZZEncryptedPassword -AdminUser "admin@domain.com"

    This example prompts the user to enter a password and stores the encrypted password in the "admin@domain.com" file.

    .EXAMPLE
    Get-MZZSecureCreds -AdminUser "admin@domain.com"

    This example retrieves the encrypted password from the "admin@domain.com" file, as part of a PS credential object.


#>

function New-MZZEncryptedPassword {
    param (
        [Parameter(Mandatory=$false)][string]$FileName
    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }
    
    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    $SecureStringPassword = Read-Host "Please enter the password you would like to hash" -AsSecureString
    
    $PasswordHash = $SecureStringPassword | ConvertFrom-SecureString

    if ($FileName) {

        Set-Content -Path "$secureCredsFolder\$($FileName).txt" -Value $PasswordHash -Force

        . Get-MZZSecureCreds -fileName $FileName

    }

    else {

        Set-Content -Path "$secureCredsFolder\$($adminUser).txt" -Value $PasswordHash -Force

        . Get-MZZSecureCreds -AdminUser $adminUser

    }

}

function Get-MZZSecureCreds {
    param (
        [Parameter(Mandatory=$false)][switch]$CheckPassword,
        [Parameter(Mandatory=$false)][string]$FileName
    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }

    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    if (!(Test-Path -Path "$localRepoPath\.local\SecureCreds\$FileName.txt")) {

        Write-Host "No password found for filename: $FileName..." -ForegroundColor Yellow

        . New-MZZEncryptedPassword -fileName $FileName

    }

    else {

        $passwordEncrypted = Get-Content -Path "$localRepoPath\.local\SecureCreds\$($FileName).txt" | ConvertTo-SecureString

        if (!$passwordEncrypted) {

            . New-MZZEncryptedPassword -fileName $FileName

        }
        
        if ($CheckPassword) {

            $global:passwordDecrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordEncrypted))

            Write-Host "Decrypted password: $passwordDecrypted" -ForegroundColor Cyan

        }

    }

}

