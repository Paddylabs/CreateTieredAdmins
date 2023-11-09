<#
  .SYNOPSIS
  Creates Tier Admin Accounts for AD Tiering Model.

  .DESCRIPTION
  Creates the required Admin Accounts for users in the AD Tiering Model.

  .PARAMETER
  None

  .EXAMPLE
  None

  .INPUTS
  A csv file with FirstName', 'LastName' and 'Password' columns if creating multiple accounts in bulk

  .OUTPUTS
  None

  .NOTES
  Author:        Patrick Horne
  Creation Date: 31/10/23
  Requires:      Active Directory Module

  Change Log:
  V1.0:         Initial Development
#>

function Get-OpenFileDialog {
    [CmdletBinding()]
    param (
        [string]
        $Directory = [Environment]::GetFolderPath('Desktop'),
        
        [string]
        $Filter = 'CSV (*.csv)| *.csv'
    )

    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openFileDialog.InitialDirectory = $Directory
    $openFileDialog.Filter = $Filter
    $openFileDialog.ShowDialog()
    $openFileDialog
}
function Import-ValidCSV {
    param (
        [parameter(Mandatory)]
        [ValidateScript({Test-Path $_ -type leaf})]
        [string]
        $inputFile,

        [string[]]
        $requiredColumns
    )

    $csvImport = Import-Csv -LiteralPath $inputFile
    $requiredColumns | ForEach-Object {
        if ($_ -notin $csvImport[0].psobject.properties.name) {
            Write-Error "$inputFile is missing the $_ column"
            exit 10
        }
    }

    $csvImport
}

#requires -module ActiveDirectory

$UPNSuffix = "@paddylab.net" #Remember to set this to match your environment

# Prompt for Tier  
$t0 = New-Object System.Management.Automation.Host.ChoiceDescription '&0', 'The accounts will be created in the Tier0 OU'
$t1 = New-Object System.Management.Automation.Host.ChoiceDescription '&1', 'The accounts will be created in the Tier1 OU'
$options = [System.Management.Automation.Host.ChoiceDescription[]]($t0, $t1)
$title = 'Which Tier'
$message = 'Which Tier are you creating accounts for?'
$result = $host.ui.PromptForChoice($title, $message, $options, 1)

# Set group name suffixes and target OU based on menu result
switch ($result)
{
    0 { $Path            = "OU=Tier0 Accounts,OU=_Tier0,DC=corp,DC=paddylab,DC=net" #Remember to set these to match your environment
        $UserNameStub    = ".tier0"
        $DisplayNameStub = " (Admin Tier 0)"        }
    1 { $Path            = "OU=Tier1 Accounts,OU=_Tier1,DC=corp,DC=paddylab,DC=net" #Remember to set these to match your environment
        $UserNameStub    = ".tier1"
        $DisplayNameStub = " (Admin Tier 1)"
}
                           
}

# Prompt for Account enabled / disabled
$t0 = New-Object System.Management.Automation.Host.ChoiceDescription '&Enabled', 'The accounts will be enabled when they are created'
$t1 = New-Object System.Management.Automation.Host.ChoiceDescription '&Disabled', 'The accounts will be disabled when created and will need to be enabled before use'
$options = [System.Management.Automation.Host.ChoiceDescription[]]($t0, $t1)
$title = 'Which method'
$message = 'Do you want to enable the accounts when creating them?'
$result1 = $host.ui.PromptForChoice($title, $message, $options, 0)

switch ($result1)
{
    0 { $AccountStatus = $true }
    1 { $AccountStatus = $false }                      
}

# Prompt for Multiple or Single Users  
$t0 = New-Object System.Management.Automation.Host.ChoiceDescription '&Single', 'I am creating a single user account'
$t1 = New-Object System.Management.Automation.Host.ChoiceDescription '&Multiple', 'I have a csv to create multiple user accounts'
$options = [System.Management.Automation.Host.ChoiceDescription[]]($t0, $t1)
$title = 'Which method'
$message = 'Do you want to create accounts for Multiple Users from a CSV File or a single User?'
$result2 = $host.ui.PromptForChoice($title, $message, $options, 0)

switch ($result2)
{
    0 {
    $first          = (Read-Host "Enter the users First name").ToLower()
    $Last           = (Read-Host "Enter the users Last name").ToLower()
        $UserDetails = @{
            GivenName         = (Get-Culture).TextInfo.ToTitleCase($first)
            Surname           = (Get-Culture).TextInfo.ToTitleCase($Last)
            Name              = (Get-Culture).TextInfo.ToTitleCase($first)+ " " +(Get-Culture).TextInfo.ToTitleCase($Last)+ $displayNameStub
            sAMAccountName    = $first.Substring(0,1) + $Last + $UserNameStub
            UserPrincipalName = ($first.Substring(0,1) + $Last + $UserNameStub) + $UPNSuffix
            DisplayName       = (Get-Culture).TextInfo.ToTitleCase($first)+ " " +(Get-Culture).TextInfo.ToTitleCase($Last)+ $displayNameStub
            Path              = $Path
            AccountPassword   = Read-Host "Password" -AsSecureString
            Enabled           = $AccountStatus
        }
        New-ADUser @UserDetails
    }

    1 { 
    $csvPath = Get-OpenFileDialog
    $Users = Import-ValidCSV -inputFile $csvpath.FileName -requiredColumns 'FirstName','LastName','Password'
    foreach ($User in $Users) {
        $first          = ($user.FirstName).ToLower()
        $Last           = ($User.LastName).ToLower()
        $SecPassword    = convertto-securestring $user.Password -asplaintext -force
            $UserDetails = @{
            GivenName         = (Get-Culture).TextInfo.ToTitleCase($first)
            Surname           = (Get-Culture).TextInfo.ToTitleCase($Last)
            Name              = (Get-Culture).TextInfo.ToTitleCase($first)+ " " +(Get-Culture).TextInfo.ToTitleCase($Last)+ $displayNameStub
            sAMAccountName    = $first.Substring(0,1) + $Last + $UserNameStub
            UserPrincipalName = ($first.Substring(0,1) + $Last + $UserNameStub) + $UPNSuffix
            DisplayName       = (Get-Culture).TextInfo.ToTitleCase($first)+ " " +(Get-Culture).TextInfo.ToTitleCase($Last)+ $displayNameStub
            Path              = $Path
            AccountPassword   = $SecPassword
            Enabled           = $AccountStatus
        }
            New-ADUser @UserDetails
    }

}
                       
}