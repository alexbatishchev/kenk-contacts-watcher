# ��������� ��������� � �������� ������ ��������� Exchange � �������� ������
. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1' Connect-ExchangeServer -auto

. .\settings.ps1
. .\uszfunctions.ps1

# ������ �� ���� Notes ������-�� �� �����������
# ������ ��� http://stackoverflow.com/questions/4286835/reading-contact-notes-field-from-exchange

#####################################################
Function sendReportToUser($sCaption, $sText, $sTo) {
	Wlog ("������� � �������� ����� sendReportToUser")			
	$sHeader = generateHtmlHeader
	$sText= $sHeader + $sText + "</body></html>"

	$sThisSMTPServer = $sSMTPServer
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($sThisSMTPServer)
	$msg.From = $sAlerterAddress
	$msg.ReplyTo = $sAlerterAddress
	$msg.To.Add($sTo)
		
	$msg.subject = $sCaption 
	$msg.body =   $sText
	$msg.IsBodyHTML = $true
	if ($sDoSendEmails) { $smtp.Send($msg)}

}


########################################################
function Compare-String {
  param(
    [String] $string1,
    [String] $string2
  )
  if ( $string1 -ceq $string2 ) {
    return -1
  }
  for ( $i = 0; $i -lt $string1.Length; $i++ ) {
    if ( $string1[$i] -cne $string2[$i] ) {
      return $i
    }
  }
  return $string1.Length
}

########################################################
function ExportContactsDataToCSV($thisMailbox,$sFileName)
{
	# ��������� "��������"
	$ContactsFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts
    $ContactsFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($ContactsFolderName, $thisMailbox)
    try {
		$ContactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $ContactsFolderId)
	}
	catch {
		$ContactFolder = $null
	}
	if ($ContactFolder -eq $null)
    {
		logred ("Error. Could not open Contacts folder for mailbox: " + $emailAddress) ([ref]$strReport)
        return $false
    }
	$aAllContacts = @()
	# �����������  
	$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100)
	$itemView | fl
	while (($folderItems = $ContactFolder.FindItems($itemView)).Items.Count -gt 0)
	{
		#����������� ������� ��������� ����� ����� ������ �� ���� �����
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		[Void]$exchService.LoadPropertiesForItems($folderItems,$psPropset)  

		foreach ($aItem in $folderItems)
		{
			$thisData = "" | Select DisplayName,CompanyName,EmailAddresses,PhoneNumbers,Birthday,JobTitle,DateTimeReceived,DateTimeCreated,LastModifiedTime, Id ,Notes # Size,DateTimeSent, LastModifiedName,
			$sTempNote = $aItem.Body.text
			if ($sTempNote -ne $null) {
				$sTempNote = $sTempNote.replace("`n"," ").replace("`r"," ")
				$sTempNote = $sTempNote -replace "<.*?>"," "
				$sTempNote = $sTempNote.replace("&nbsp;"," ")
				$sTempNote = $sTempNote.replace("&#43;","+")
				$sTempNote = $sTempNote -replace '\s+', ' '
				$thisData.Notes	= $sTempNote
			}	
			else {
				$thisData.Notes	= ""
			}

			$thisData.Id 		= $aItem.Id 		
			$thisData.DisplayName 		= $aItem.DisplayName 		
			$thisData.CompanyName       = $aItem.CompanyName
			$thisData.EmailAddresses      = $aItem.EmailAddresses
			$thisData.PhoneNumbers        = $aItem.PhoneNumbers
			$thisData.Birthday            = $aItem.Birthday
			$thisData.JobTitle            = $aItem.JobTitle
			$thisData.DateTimeReceived    = $aItem.DateTimeReceived
			$thisData.DateTimeCreated     = $aItem.DateTimeCreated
			$thisData.LastModifiedTime    = $aItem.LastModifiedTime
			
			#################################
			# �������� ������ �� ���� �������� ���� � ������
			if ($aItem.PhoneNumbers -ne $null) {
				$sCollectedString = ""
				$eEnumNames = [enum]::getvalues([Microsoft.Exchange.WebServices.Data.PhoneNumberKey])
				$sPh = ""
				foreach ($enumName in $eEnumNames) {
					$bRet = $aItem.PhoneNumbers.TryGetValue($enumName,[ref] $sPh)
					if ($bRet) {
						if ($sPh -ne "" -and $sPh -ne $null) {
							$sPh = "[" + $sPh + "]"
							if ($sPh -ne "[]") {
								$sCollectedString = $sCollectedString + $sPh
							}	
						}	
					}
				}
				$thisData.PhoneNumbers = $sCollectedString
			}	
			#################################
			# �������� ������ �� ���� �������� ���� � ������
			if ($aItem.EmailAddresses -ne $null) {
				$sCollectedString = ""
				$eEnumNames = [enum]::getvalues([Microsoft.Exchange.WebServices.Data.EmailAddressKey])
				$sPh = ""
				foreach ($enumName in $eEnumNames) {
					$bRet = $aItem.EmailAddresses.TryGetValue($enumName,[ref] $sPh)
					if ($bRet) {
						if ($sPh -ne "" -and $sPh -ne $null) {
							$sPh = "[" + $sPh + "]"
							if ($sPh -ne "[]") {
								$sCollectedString = $sCollectedString + $sPh
							}	
						}	
					}
				}
				$thisData.EmailAddresses = $sCollectedString
			}	
			
			$aAllContacts = $aAllContacts + $thisData
		}
		$offset += $folderItems.Items.Count
		$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(100, $offset)
	}
	#��������� ������ � CSV ��� ������������ �������
	
	$aAllContacts | sort-object Id | Export-Csv $sFileName  -Encoding:UTF8 -notype -Delimiter ";" 
	return $true
} # function ExportContactsDataToCSV

########################################################
function IfThereWereChanges ($sAddress, $sReferencePath) {
	Wlog ("��������� ���� $sAddress")
	$dir = $PSScriptRoot + "\out\" + $sAddress
	#������� ��� ����� ������ ����� ��� ���������
	Wlog ("���� ����� �� ���� $dir")
	$latest = Get-ChildItem -Path $dir | Sort-Object Name  -Descending | Select-Object -First 5

	if ($latest.Count -lt 2) {
		logred("������� ������ ��� ��� ����� ��������� ��� ����� $sAddress ") ([ref]$strReport)
		return $true
	}
	if ($latest[0].FullName -ne $sReferencePath) {
		logred("�� ��������� ���������� ���� �������� " + $latest[0].FullName + "   " + $sReferencePath) ([ref]$strReport)
		return $true
	}
	$bRetVal = $false
	
	
	$CSVData = Import-Csv $latest[1].FullName -Encoding "UTF8" -Delimiter:";"
	$tCsvCounter = 0
	$tCsvCounter  = $CSVData.count
	
	if (($tCsvCounter -eq $null) -or ($CSVData -isnot [system.array])) {
		logred("������ �������� ����������� ����� �������� " + $latest[1].FullName + " ��� ����� $sAddress") ([ref]$strReportDev)
		return $false
	}
	#��������� ���� �� CSV ��� ���������� ������� 

	
	$sNewHeader = (Get-Content $latest[1].FullName)[0] + ';"CompareResult"'
	$aResultArr = @()
	$aResultArr = $aResultArr + $sNewHeader
	# ���������� 
	$aCompared = Compare-Object -ReferenceObject (Get-Content $latest[0].FullName) -DifferenceObject (Get-Content $latest[1].FullName)
	
	Wlog ( "���������� " + $latest[0].FullName + " � " + $latest[1].FullName)
	if ($aCompared) { # ���� �����-�� ���������
		$bRetVal = $true
		#$aCompared | fl
		if ($aCompared.count -gt $iSuspiciousChangesCount ) {
			logred ("������������� ����� (������ $iSuspiciousChangesCount) ��������� ���������� ��� ����� $sAddress ("+ $aCompared.count + "), �������� ��������� ���� ����������� ��� �������� �������� �� Exchange, ������������� ��������� ���������� ������") ([ref]$strReportDev)
			$sNewName0 = ($env:temp) +  $latest[0].Name
			$sNewName1 = ($env:temp) +  $latest[1].Name
			Copy-Item -Path $latest[0].FullName -Destination $sNewName0
			Copy-Item -Path $latest[1].FullName -Destination $sNewName1
			log ("������������ ���� ���������� � $sNewName0")
			log ("������������ ���� ���������� � $sNewName1")
			$bRetVal = $false
			return 	$bRetVal
		}
		else {
			wlog ("����� ��������� ����������: "+ $aCompared.count) ([ref]$strReportDev)
		}
		foreach ($item in $aCompared) {
			$sOut = $item.InputObject + ';"' + $item.SideIndicator + '"'
			$aResultArr = $aResultArr + $sOut
		}
		$aResultArr = $aResultArr | ConvertFrom-CSV -Delimiter:";"

		# �������� ��������� ��������� � ������
		$aResultArr = $aResultArr | group-object Id

		#�������� �� ���� ������������ ����������
		foreach ($oResult in $aResultArr) {
			if ($oResult.Count -eq 2) {
				if ($oResult.Group[0].LastModifiedTime -eq $oResult.Group[1].LastModifiedTime) {
					if ($oResult.Group[0].DisplayName -ne "��������� ������ ������������") { # ���� ����, ����� ���� � ���� � ������ ����� ������� ����-���� � . ������� ��� ����
						logH2 ("���������� ��������� ������ ��� ��������� LastModifiedTime � �������� ����� ����� $sAddress") ([ref]$strReportDev)
						loggreen ("�������� ������") ([ref]$strReportDev)
						foreach ($oVal in $oResult.Group) {
							if ($oVal.CompareResult  -eq "=>") {
								$sVerd = "�������� ������:"
							} 	else {
								$sVerd = "���������� ������:"
							}
							logH3 ($sVerd) ([ref]$strReportDev)
							$showVAL = $oVal | select @{N='�����������'; E={$_.CompanyName}},@{N='���� ��������'; E={$_.Birthday}},@{N='e-mail(�)'; E={$_.EmailAddresses}},@{N='���'; E={$_.DisplayName}},@{N='�������(�)'; E={$_.PhoneNumbers}}, @{N='���� ��������'; E={$_.DateTimeCreated}}, @{N='���� ���������� ���������'; E={$_.LastModifiedTime}}, @{N='���������'; E={$_.JobTitle}}, @{N='�������'; E={$_.Notes}}
							$tRet = logtable $showVAL ([ref]$strReportDev) 
						}
					}	
				}
				else {
					logH2 ("���������� ��������� � �������� ����� ����� $sAddress") ([ref]$strReport)
					log ("�������� ������, ���� � ����������� ���������� �������") ([ref]$strReport)
					
					$aPrintableValues = @()
					foreach ($oVal in $oResult.Group) {
						$showVAL = $oVal | select @{N='�����������'; E={$_.CompanyName}},@{N='���� ��������'; E={$_.Birthday}},@{N='e-mail(�)'; E={$_.EmailAddresses}},@{N='���'; E={$_.DisplayName}},@{N='�������(�)'; E={$_.PhoneNumbers}}, @{N='���� ��������'; E={$_.DateTimeCreated}}, @{N='���� ���������� ���������'; E={$_.LastModifiedTime}}, @{N='���������'; E={$_.JobTitle}}, @{N='�������'; E={$_.Notes}}
						$aPrintableValues = $aPrintableValues + $showVAL
					}

					$varList = $aPrintableValues[0] | Get-Member -membertype properties | select -expand Name
					foreach ($var in $varList) {
						$tRes = Compare-String $aPrintableValues[0].$var $aPrintableValues[1].$var 
						if ($tRes -ne -1) {
							#there is difference in strings
							$aPrintableValues[0].$var = '<span style="color:red">' + $aPrintableValues[0].$var + '</span>'
							$aPrintableValues[1].$var = '<span style="color:red">' + $aPrintableValues[1].$var + '</span>'
						}
					}
					$tIndex = 0
					foreach ($oVal in $oResult.Group) {
						
						if ($oVal.CompareResult  -eq "=>") {
							$sVerd = "�������� ������:"
						} 	else {
							$sVerd = "���������� ������:"
						}
				 
						logH3 ($sVerd) ([ref]$strReport)
					
						$tRet = logtable $aPrintableValues[$tIndex] ([ref]$strReport) 
						$tIndex = $tIndex + 1
					}
					$sText = "���������� ����� �����: <br>" + $latest[0].FullName + "<br>" + $latest[1].FullName 
					loggray ($sText) ([ref]$strReport)
				}
			}
			elseif ($oResult.Count -eq 1)
			{
				$oVal = $oResult.Group[0]
				if ($oVal.CompareResult  -eq "=>") {
					$sVerd = "������� ������:"
				} 	else {
					$sVerd = "��������� ������:"
				}
				logH2 ("���������� ��������� � �������� ����� ����� $sAddress") ([ref]$strReport)
				loggreen ($sVerd) ([ref]$strReport)
				$showVAL = $oVal | select @{N='�����������'; E={$_.CompanyName}},@{N='���� ��������'; E={$_.Birthday}},@{N='e-mail(�)'; E={$_.EmailAddresses}},@{N='���'; E={$_.DisplayName}},@{N='�������(�)'; E={$_.PhoneNumbers}}, @{N='���� ��������'; E={$_.DateTimeCreated}}, @{N='���� ���������� ���������'; E={$_.LastModifiedTime}}, @{N='���������'; E={$_.JobTitle}}, @{N='�������'; E={$_.Notes}}
				$tRet = logtable $showVAL ([ref]$strReport) 
				$sText = "���������� ����� �����: <br>" + $latest[0].FullName + "<br>" + $latest[1].FullName 
				loggray ($sText) ([ref]$strReport)
			}
			else {
				logred ("������ ���������") ([ref]$strReportDev)
				foreach ($oVal in $oResult.Group) {
					$showVAL = $oVal | select @{N='�����������'; E={$_.CompanyName}},@{N='���� ��������'; E={$_.Birthday}},@{N='e-mail(�)'; E={$_.EmailAddresses}},@{N='���'; E={$_.DisplayName}},@{N='�������(�)'; E={$_.PhoneNumbers}}, @{N='���� ��������'; E={$_.DateTimeCreated}}, @{N='���� ���������� ���������'; E={$_.LastModifiedTime}}, @{N='���������'; E={$_.JobTitle}}, @{N='�������'; E={$_.Notes}}, CompareResult 
					$tRet = logtable $showVAL ([ref]$strReportDev) 
				}
			}
		} # foreach #�������� �� ���� ������������ ����������
	} # ���� �����-�� ���������
	return 	$bRetVal 
}

#####################################################
# E N T R Y
#####################################################

$StartDate=(GET-DATE)

$strReport = ""
$strReportAll = ""
$strReportDev = ""


#########################################
# ����������� � EWS � ������ ������ ����������
# Update the path below to match the actual path to the EWS managed API DLL.
Import-Module -Name ".\Microsoft.Exchange.WebServices.dll"
$mailboxList = new-object 'System.Collections.Generic.List[string]'
# If a URL was specified we'll use that; otherwise we'll use Autodiscover 
$exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010) 

# no creds because we runing script as apropriate user
#     $exchService.Credentials = new-object System.Net.NetworkCredential($UserName, $password, "") 

$HostName = $sEWSHost
Wlog ("##################################################################") 
if ($HostName -ne "") 
{ 
    Wlog ("Using EWS URL " + "https://" + $HostName + "/EWS/Exchange.asmx") 
    $exchService.Url = new-object System.Uri(("https://" + $HostName + "/EWS/Exchange.asmx")) 
} 
else
{ 
    ("Autodiscovering " + $mailboxList[0] + "...")
    $exchService.AutoDiscoverUrl($mailboxList[0], {$true}) 
}

if ($exchService.Url -eq $null) 
{ 
	logred("exchService.Url -eq $null") ([ref]$strReport)
    return 
}


#########################################################
# �������� ������

#########################################################
# ������� ������ ������ �� ������� ��� ��������

$sGroupDN = $sGroupToWatchDN
$aAllMailboxes = Get-DistributionGroupMember -Identity $sGroupDN | Get-Mailbox
# for only 1 mailbox
#    $aAllMailboxes = Get-Mailbox "user@domain.com" 

$aMailboxesReviwed = @()

foreach ($tThisMailbox in $aAllMailboxes) {

	$aMailboxesReviwed = $aMailboxesReviwed + ([system.String] $tThisMailbox.PrimarySmtpAddress)

	$mbx = new-object Microsoft.Exchange.WebServices.Data.Mailbox($tThisMailbox.PrimarySmtpAddress)

	# ��������
	$sPath = $PSScriptRoot + "\out\" + $tThisMailbox.PrimarySmtpAddress + "\"
	$tRet = New-Item -ItemType Directory -Force -Path $sPath
	$sOutFileName = $sPath + (Get-Date).ToString("yyyy-MM-dd-HH-mm-ss")+ ".log.csv"

	$strReport = ""

	Wlog ("%%%%%%%%%%%%%%%%%%%%%%")			
	Wlog ("��������� ������ �� " + $tThisMailbox.PrimarySmtpAddress + " � ���� $sOutFileName")			
	$bRet = ExportContactsDataToCSV $mbx $sOutFileName
	
	if (-not ($bRet)) {
		Wlog ("������ ��������, ���������� ����")			
		continue
	}
	Wlog ("������� ���������")			
			
	#checking if there is more than 1 file to compare
	$dir = $PSScriptRoot + "\out\" + $sAddress
	Wlog ("checking if there is more than 1 file to compare at  $dir")
	$latest = Get-ChildItem -Path $dir | Sort-Object Name  -Descending | Select-Object -First 5
	if ($latest.Count -lt 2) {
		# first dump, no need to report
		logred("������� ������ ��� ��� ����� ��������� ��� ����� $sAddress. ��������� ����, �������� ����� ��������� ��� ������� ����� ��� ������ �����.") ([ref]$strReport)
		continue
	}
			
	#now checking if there were differences between this and previous dump
	$bRet = IfThereWereChanges $tThisMailbox.PrimarySmtpAddress $sOutFileName

	if (-not $bRet) {
		# no differences, we can delete this dump, no need to report
		Wlog ("��� ���������, �������� �������� ���� $sOutFileName")			
		Remove-Item $sOutFileName
		continue
	} 
	
	#we found differences
	Wlog ("���������� ���������, ��������� ���� $sOutFileName � ��������� ���������� � �����")
	$strReportAll = $strReportAll + $strReport

	# checking if we must send personal report for current mailbox
	if ($tThisMailbox.CustomAttribute7 -ne $null) {
		$sTempTo = $tThisMailbox.CustomAttribute7
		Wlog ("����� ����� ��� ������������� ������ � ���� CustomAttribute7 �������� $sTempTo, �������� ������������ �����")
		if ($strReport -eq "") {
			Wlog ("����� ������, ������ �������")
		}
		else {
			if ($sDoSendEmails) { 
				sendReportToUser "���������� ��������� � �������� �����" $strReport $sTempTo	
			}
			Wlog ("�������")
		}	
	}	
}

$sMailobxesReviwed = [system.String]::Join(", ", $aMailboxesReviwed)
Wlog ("�������� ������ ������ $sMailobxesReviwed")			

if ($strReportAll -eq "") {
	Wlog ("�������� ������, ����������� ������")			
}

if ($strReportAll -ne "") {
	Wlog ("������� � �������� ������")			
	$sHeader = generateHtmlHeader
	$strReport = $sHeader+ $strReportAll
	$EndDate=(GET-DATE)
	$sTimediff = NEW-TIMESPAN �Start $StartDate �End $EndDate

	log ("#####################################################") ([ref]$strReport)
	log ("����� �� ��������: $sMailobxesReviwed") ([ref]$strReport)
	log ("����� �������� ������: " + $EndDate ) ([ref]$strReport)
	log ("������������ ���������: " + $sTimediff ) ([ref]$strReport)
	log ("����� ����������� ��������  " + $MyInvocation.MyCommand.Definition + " �� " + "$env:computername.$env:userdnsdomain" ) ([ref]$strReport)

	$strReport =  $strReport + "</body></html>"

	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($sSMTPServer)
	$msg.From = $sAlerterAddress
	$msg.ReplyTo = $sAlerterAddress
	$msg.IsBodyHTML = $true
	
	$msg.To.Add($sRegularReportAddress )
	$msg.subject = "���������� ��������� � �������� �����"
	$msg.body =   $strReport

	$sResPath = $PSScriptRoot + "\res\"
	$tRet = New-Item -ItemType Directory -Force -Path $sResPath
	$sResOutFileName = $sResPath + (Get-Date).ToString("yyyy-MM-dd-HH-mm-ss")+ ".html"
	$strReport | Out-File $sResOutFileName

	#Sending email 
	if ($sDoSendEmails) { $smtp.Send($msg) }
}

if ($strReportDev -ne "") {
#if ($false) {
	Wlog ("������� � �������� ������ Dev")			
	$sHeader = generateHtmlHeader
	$strReportDev = $sHeader+ $strReportDev
	$EndDate=(GET-DATE)
	$sTimediff = NEW-TIMESPAN �Start $StartDate �End $EndDate

	log ("#####################################################") ([ref]$strReportDev)
	log ("����� �� ��������: $sMailobxesReviwed") ([ref]$strReportDev)
	log ("����� �������� ������: " + $EndDate ) ([ref]$strReportDev)
	log ("������������ ���������: " + $sTimediff ) ([ref]$strReportDev)
	log ("����� ����������� ��������  " + $MyInvocation.MyCommand.Definition + " �� " + "$env:computername.$env:userdnsdomain" ) ([ref]$strReportDev)

	$strReportDev =  $strReportDev + "</body></html>"

	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($sSMTPServer)
	$msg.From = $sAlerterAddress
	$msg.ReplyTo = $sAlerterAddress
	$msg.IsBodyHTML = $true


	$msg.To.Add($sDevReportAddress)
	$msg.subject = "���������� ��������� � �������� ����� - Dev �����"
	$msg.body =   $strReportDev
	
	if ($sDoSendEmails) { $smtp.Send($msg) }
}