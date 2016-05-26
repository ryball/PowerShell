#Example of use
<#

You have to know the username used in the set-mycredential function.
Connect-MsolService -Credential (get-mycredential -User "Username" -File ".\Filename.cred")

#>

function set-mycredential ($File)
{
    $Credential = Get-Credential
    $Credential.Password | ConvertFrom-SecureString | Set-Content $File
}

function get-mycredential ($User, $File)
{
    $Password = Get-Content $File | ConvertTo-SecureString
    $Credential = New-Object System.Management.Automation.PSCredential($User,$Password)

    $Credential
}

function LogWrite
{
    param (
        [String]$Logfile,
        [String]$LogString
        )

    Add-Content -Path $Logfile -Value $LogString
}




