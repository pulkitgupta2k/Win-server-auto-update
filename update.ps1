$ErrorActionPreference = "SilentlyContinue"
If ($Error) {
	$Error.Clear()
}
$Today = Get-Date

$Searcher = get-WindowsUpdate

Write-Host
Write-Host "`t Initialising and Checking for Applicable Updates. Please wait ..." -ForeGroundColor "Yellow"

If($Searcher.count -EQ 0){
	Write-Host "`t There are no applicable updates for this computer."
}

Else{
	$ReportFile = $Env:ComputerName + "_Report.txt"
	If (Test-Path $ReportFile) {
		Remove-Item $ReportFile
	}
	New-Item $ReportFile -Type File -Force -Value "Windows Update Report For Computer: $Env:ComputerName`r`n" | Out-Null
	Add-Content $ReportFile "Report Created On: $Today`r"
	Add-Content $ReportFile "==============================================================================`r`n"
	Write-Host "`t Preparing List of Applicable Updates For This Computer ..." -ForeGroundColor "Yellow"
	Add-Content $ReportFile "List of Applicable Updates For This Computer`r"
	Add-Content $ReportFile "------------------------------------------------`r"
	For ($Counter = 0; $Counter -LT $Searcher.count; $Counter++){
		Add-Content $ReportFile $Searcher[$Counter].Title
		Add-Content $ReportFile $Searcher[$Counter].KB
		Add-Content $ReportFile ".........................`r"
	}
	$Counter = 0
	Add-Content $ReportFile "==============================================================================`r`n"
	##$Session = Install-WindowsUpdate -acceptall
	Add-Content $ReportFile "`r`n"
	Write-Host "`t Initialising Download of Applicable Updates ..." -ForegroundColor "Yellow"
	Add-Content $ReportFile "Downloading and Installing"
	Add-Content $ReportFile "------------------------------------------------`r"
	$FlagError = 0
	For ($Counter = 0; $Counter -LT $Searcher.count; $Counter++){
		$DowIns = (Get-WindowsUpdate -KBArticleID $Searcher[$Counter].KB -Install -Confirm:$false)
		Write-Output $DowIns
		Add-Content $ReportFile $DowIns.Title
		Add-Content $ReportFile $DowIns.Status
		Add-Content $ReportFile $DowIns.DownloadResult
		Add-Content $ReportFile $DowIns.InstallResult
		Add-Content $ReportFile $DowIns.Description
		If($DowIns.InstallResult -eq "Failed"){
			##email
			$FlagError = 1
		}
	}
	If($FlagError -eq 1){
		#EMAIL
		$emailSmtpServer = "smtp.gmail.com"
                $emailSmtpServerPort = "587"
               	$emailSmtpUser = "username@gmail.com"
              	$emailSmtpPass = "password"
		        $attachment = "C:\Users\pulkit_gupta_2k\Desktop\WIN-SERVER_Report.txt"	

                $emailFrom = "from_user@gmail.com"
               	$emailTo = "to_user@gmail.com"
               	#$emailcc="myboss@gmail.com"

                $emailMessage = New-Object System.Net.Mail.MailMessage( $emailFrom , $emailTo )
                #$emailMessage.cc.add($emailcc)
                $emailMessage.Subject = "My test mail" 
                $emailMessage.Body = "Hii Hello"
		        $attach = new-object Net.Mail.Attachment($attachment) 
		        $emailMessage.Attachments.Add($attach)

                $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
                $SMTPClient.EnableSsl = $True
                $SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
                $SMTPClient.Send( $emailMessage )
		#REBOOT
		Restart-Computer
	}
	If($Searcher.count -ge 1){
		#REBOOT
		Restart-Computer
	}
}