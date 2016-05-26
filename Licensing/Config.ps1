#Config File

#TargetAddress
$TargetAddress = ''
$UsageLocation = 'NO'

#How to run the script (Test = $false / Prod = $True) 
$ProdRun = $false

<#Verbose Settings 
     Stop : Displays the verbose message and an error message and then stops executing.
     Inquire : Displays the verbose message and then displays a prompt that asks you whether you want to continue.
     Continue : Displays the verbose message and then continues with execution.
     SilentlyContinue: Does not display the verbose message. Continues executing. (Default)
    #>
$VerbosePreference = "Continue"

#Msol Credentials (O365 Global admin):
$MsolUserName = ""
$CredFile = ""

#LogFile
$LogFilePath = ".\log\"

#Exception File ($null if you dont whant to use and exceptionfile) Example:  $ExceptionFile = ".\Exceptions.csv"
$ExceptionFile = ".\Exceptions.csv"

#UserDefined Sku CSV File
$UserSkuFile = ".\AccountSku.csv"

#Report
$usereport = $true   # $true / $false
$format = "HTML"     #PDF, EXCEL, HTML
$ReportPath = ".\Report\"


#Email Setup
$SendMail = $true

$EmailAddress = ""
$EmailCredfile = ""

$emailSmtpServer = "smtp.office365.com"
$emailSmtpServerPort = "587"
