#Repport

Param (
    [string]$Path
)

#region Setup
#check if the module is installed, if not download and install
if (!(Get-Module ImportExcel))
{
    if($PSVersionTable.PSVersion.Major -gt 4)
    {
        try 
        {
            Install-Module -Name ImportExcel
        } catch {
            try
            {
                Invoke-Expression -Command (new-object System.Net.WebClient).DownloadString('https://raw.github.com/dfinke/ImportExcel/master/Install.ps1')
            } catch {
                Write-Error "Error installing Excel Module."
            }
        }
    }
}

$ReportFile = $Path + "Report-"+(Get-Date -Format "ddMMyyyy-HHmm")+".html"
Write-Verbose "Report will be generated and stored in : $ReportFile"
#endregion Setup


$html = "<h1>Office 365 licensing report</h1>";
$MultiLicense = $LicensedUserDetails | Where-Object {$_.License.count -gt 2}

$html += "<h3>Multilicensed users</h3>`n";
$html += "<table><tr><th align=""left"" width=""300""><b>Displayname</b></th><th align=""left"" width=""300""><b>UPN</b></th></tr>`n";

$MultiLicense | %{$html += "<tr><td>"+$_.UserPrincipalName+"</td><td>"+$_.License+"</td></tr>`n"}
$html += "</table>`n";

$html += "<h3>Licensing Reports</h3>`n";
$html += "<table><tr><th align=""left"" width=""300""><b>Description</b></th><th align=""left"" width=""300""><b>Value</b></th></tr>`n";

$html += "<tr><td>Added licencse in this run </td><td>"+$addedCount+"</td></tr>`n"
$html += "<tr><td>Updated licencse in this run </td><td>"+$ChangeCount+"</td></tr>`n"
$html += "<tr><td>Removed licencse in this run </td><td>"+$RemoveCount+"</td></tr>`n"
$html += "</table>`n";

if(($dirsyncIssue | measure).Count -gt 0) {
	$html += "<h3>Active AD users not in Office 365 (DirSync issue)</h3>`n";
	$html += "<table><tr><th align=""left"" width=""300""><b>Displayname</b></th><th align=""left"" width=""300""><b>UPN</b></th></tr>`n";
	$dirsyncIssue | %{$html += "<tr><td>"+$_.DisplayName+"</td><td>"+$_.UserPrincipalName+"</td></tr>`n"}
	$html += "</table>`n";
}

if(($couldnotlicense | measure).Count -gt 0) {
	$html += "<h3>Exceptions</h3>`n";
	$html += "<table><tr><th align=""left"" width=""300""><b>UserPrincipalName</b></th><th align=""left"" width=""500""><b>Exception</b></th></tr>`n";
	for($i = 0; $i -lt $couldnotlicense.Count; $i+=2) {
        $html += "<tr><td>{0}</td><td>{1}</td></tr>`n" -f $couldnotlicense[$i].UserPrincipalName,$couldnotlicense[$i+1]
	}
    $html += "</table>`n";
}

Get-MsolAccountSku | ForEach-Object {
        $html += "<tr><td>"+$_.AccountSkuId+" available</td><td>"+$_.ActiveUnits+"</td></tr>`n"
		if($_.ConsumedUnits -ge $_.ActiveUnits -and $_.ConsumedUnits -ne 0) {
			$html += "<tr><td style=""color: #ff0000;"">"+$_.AccountSkuId+" assigned</td><td style=""color: #ff0000;"">"+$_.ConsumedUnits+"</td></tr>`n"
		} else {
			$html += "<tr><td>"+$_.AccountSkuId+" assigned</td><td>"+$_.ConsumedUnits+"</td></tr>`n"
		}
}
$html += "</table>`n";

$html | Out-File $ReportFile

if($SendMail) 
{
    Write-Verbose "Sending reports"
	
    $encryptedpassword = Get-Content $EmailCredfile | ConvertTo-SecureString
	$smtpcredential = New-Object System.Management.Automation.PsCredential($EmailAddress, $encryptedpassword)

	$msg = new-object Net.Mail.MailMessage
	$msg.IsBodyHtml = $true;
	$msg.Body = "<html><body>$html</body></html>";
	$msg.Subject = "Office 365 licensing report"
	$msg.From = ""
    
	$msg.To.Add("")
	
	$smtp = new-object Net.Mail.SmtpClient($emailSmtpServer , $emailSmtpServerPort)
    $smtp.EnableSsl = $true
	$smtp.Credentials = $smtpcredential
	$smtp.Send($msg)
}
