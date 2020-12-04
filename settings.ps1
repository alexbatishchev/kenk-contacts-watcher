#####################################################
# settings.ps1
# edit this variables to customize main script
#####################################################

# Количество изменений, больше которого не фиксировать отчёт (возможно произошёл сбой при выгрузке объектов из Exchange)
$iSuspiciousChangesCount = 25

$sLogFileNameTemplate = "yyyy-MM-dd-HH-mm-ss" #"yyyy-MM-dd-HH-mm-ss"
$sLogFilePathTemplate = "yyyy-MM-dd"

$sDoSendEmails = $true # $false 

#mail to send alerts from
$sAlerterAddress = "account@domain.com"

#mail to send alert to
$sRegularReportAddress = "monitoring@domain.com"
#mail to send developer alerts to
$sDevReportAddress = "admin@domain.com"

$sSMTPServer = "smtpserver.domain.com"

$sEWSHost = "ewsserver.domain.com"

$sGroupToWatchDN = "CN=MailboxesToWatch,DC=domain,DC=com" 
