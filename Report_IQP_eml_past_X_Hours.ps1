Start-Transcript -Path "X:\patch\TranscriptDailyRun_IQP.txt" -Append
function Get-TimeStamp {return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)}

Write-Host "$(Get-TimeStamp) Start of script by user $env:UserName on the server $env:computername."
$timestampStartFull = get-date
$dt = Get-Date -format s
$DateFormat=$dt.replace(":",".")
Write-Host "$(Get-TimeStamp) Setting the parameters"
$StartDate = (Get-date).AddHours(-25)
$EndDate = (Get-date)


Write-Host "$(Get-TimeStamp) Defined Start date $StartDate and EndDate $EndDate"
$outFile = "X:\Scripts\Logs\LogForEEL\LogOfDay" + $DateFormat + ".csv"
Write-Host "$(Get-TimeStamp) Setting log file to $outfile"
Write-Host "$(Get-TimeStamp) Reset the index"
$indexNumber = 0
$indexNumberIQPcases = 0
Write-Host "$(Get-TimeStamp) The $indexNumber should be 0"

$LogEELfolder = "X:\path"


Write-Host "$(Get-TimeStamp) Setting eml file folder to $logeelfolder"
Write-Host "$(Get-TimeStamp) Setting the category specific folder to $CategoryFolder "
Write-Host "$(Get-TimeStamp) Loading array with files from $logeelfolder based on filter"

$AllFilesToInspect = Get-ChildItem -Path $logeelfolder | where-object {$_.LastWriteTime -gt $StartDate -AND $_.lastwritetime -lt $EndDate}



Write-Host "$StartDate and $EndDate"
Write-Host "$(Get-TimeStamp) Array loaded with $($AllFilesToInspect.Count) files"

Write-Host "$(Get-TimeStamp) Clearing the array for the results"

Write-Host "$(Get-TimeStamp) Logging the report header"

#logging the report header
$outputArray = "Usage report of old IQP email addresses.<br>Report period: $StartDate to $EndDate<br>Investigated a total of $($AllFilesToInspect.Count) files<br><br>CSV file is also provided attached."
[PSObject[]]$OutputIQPCases = @()
Clear-Variable -Name OutputIQPCases
$allSenders = @()

<#
$Text =get-content -LiteralPath $AllFilesToInspect[0].fullname -TotalCount 20
$text | Where-Object {$_.Contains("x-receiver: ")}

#>



Write-Host "$(Get-TimeStamp) File inspection starting"
foreach ($filetoInspect in $AllFilesToInspect){
	$indexNumber += 1
    Write-Host "$(Get-TimeStamp) Loop $indexNumber - Start of the loop"

    #the Contains switch is case senstive
    Write-Host "$(Get-TimeStamp) Loop $indexNumber - Reading file $($filetoInspect.FullName) with size $($filetoInspect.Length)"
    $text = Get-Content -LiteralPath $filetoInspect.FullName -TotalCount 200
	$ReceiverContainer = $text | Where-Object {$_.Contains("x-receiver: ")}
    $SenderContainer = $text | Where-Object {$_.Contains("x-sender: ")}
    $CategoryContainer = "EMPTY"
	Write-Host "$(Get-TimeStamp) Loop $indexNumber - starting category for $($filetoInspect.FullName) "
    if ($ReceiverContainer -like "*alias0@internal.domain.com*"){ $CategoryContainer = "ALIAS_0"}
    elseif ($ReceiverContainer -like "*alias1@internal.domain.com"){ $CategoryContainer = "ALIAS_1"}
    elseif ($ReceiverContainer -like "*alias2@internal.domain.com"){ $CategoryContainer = "ALIAS_2"}
    elseif ($ReceiverContainer -like "*alias3@internal.domain.com){ $CategoryContainer = "ALIAS_3"}
   else { $CategoryContainer = "Any_Other"}
    
    Write-Host "$(Get-TimeStamp) Loop $indexNumber - Running the IQP check"
    if ($CategoryContainer -like "IQP*"){
        Write-Host "$(Get-TimeStamp) Loop $indexNumber - Positive IQP loop" -ForegroundColor Blue -BackgroundColor White
        $ReceiverEmail = $ReceiverContainer.ToLower().Replace("x-receiver: ","").Replace("@infinet.infineum.com","")
        $SenderEmail = $SenderContainer.ToLower().Replace("x-sender: ","")
        Write-Host "$(Get-TimeStamp) Loop $indexNumber - Writing result to array" -ForegroundColor Blue -BackgroundColor White
        
        
        write-host "$SenderEmail"
        
               
        $allSenders += $SenderEmail + ";"
        
        $Email = @{Name = $($filetoInspect.Name)
        Receiver = $ReceiverEmail
        Sender = $SenderEmail
        Category = $CategoryContainer
        Date = $text | Where-Object {$_.Contains("Date: ")} 
        }
        
        $emailobject = New-Object -TypeName PsObject -Property $email
        
        $OutputIQPCases += $emailobject
        
        
        $indexNumberIQPcases += 1
        Write-Host "$(Get-TimeStamp) Loop $indexNumber - End of positive IQP check" -ForegroundColor Blue -BackgroundColor White
        }
    else {
        Write-Host "$(Get-TimeStamp) Loop $indexNumber - Not an IQP item. $ReceiverContainer"
        }
    
    Write-Host "$(Get-TimeStamp) Loop $indexNumber - End of loop, count of IQP cases: $indexNumberIQPcases" -ForegroundColor DarkBlue -BackgroundColor White
    #This is troubleshooting break
    #if ($indexNumber -gt 200){break}

}

Write-Host "$(Get-TimeStamp) Ended loop through all $indexNumber files"
$outputArray += "Number of IQP Cases: $indexNumberIQPcases out of $indexNumber EEL files<br><br>"
$OutputIQPCases | Export-Csv -Path $outFile
Write-Host "The type of the OutputIQPCases $($OutputIQPCases.GetType())"
$OutputIQPCases.gettype()



#Checking the time elapsed
$timestampEndFull = Get-Date
Write-Host "$(Get-TimeStamp) Calculating the elapsed $timestampEndFull MINUS $timestampStartFull"
$elapsedTime = $timestampEndFull - $timestampStartFull
Write-Host "$(Get-TimeStamp) Calculated the elapsed time: $elapsedTime"

#variables for the emailing function 
Write-Host "$(Get-TimeStamp) Setting up email variables including body"
$SmtpServer = "smtp.domain.com"
$MailFrom = "email"
$MailtToITAdmin	= "email admin" 
$MailtToIQPAdmin = "email other admin"


$BottomoutputArray += "These are all the senders in order (for easy copy/paste):<br>"
$BottomoutputArray += $allsenders | Sort-Object | Get-Unique
$BottomoutputArray += "<br><br>Elapsed time: $elapsedTime<br>Time stamp: $dt by $env:UserName on the server $env:computername."

$PostContent = $BottomoutputArray | Out-String
$preContent = $outputArray | Out-String

$FinaloutputArray += $OutputIQPCases | ConvertTo-Html -postContent $PostContent -PreContent $preContent -Title "IQP 1.11 Old EEL Report"

$MailBody += $FinaloutputArray | Out-String

$Mailbody = $MailBody.Replace("<table>",'<table border="1" cellpadding="1" cellspacing="1">')



if ($indexNumberIQPcases -gt 0){
    Write-Host "$(Get-TimeStamp) There are $indexNumberIQPcases IQP cases today."
    Write-Host "$(Get-TimeStamp) Sending email on $smtpServer"
    $MailSubject = "EEL IQP Report - $indexNumberIQPcases cases of wrong use of IQP emails today - $($timestampStartFull.ToShortDateString())"
    Send-MailMessage -To $MailtToIQPAdmin -Cc $MailtToITAdmin -from $MailFrom -Subject $MailSubject -SmtpServer $SmtpServer -Body $MailBody -Priority Low -BodyAsHtml -Attachments $outFile
}
else {
    Write-Host "$(Get-TimeStamp) There are 0 cases today. Index of IQP: $indexNumberIQPcases."
    $MailSubject = "EEL IQP Report - $indexNumberIQPcases cases of wrong use of IQP emails today - $($timestampStartFull.ToShortDateString())"
    $MailBody = "Investigated a total of $($AllFilesToInspect.Count) files and there were no IQP cases to report on the period: $StartDate to $EndDate<br><br>Elapsed time: $elapsedTime<br>Time stamp: $dt by $env:UserName on the server $env:computername."
    Send-MailMessage -To $MailtToIQPAdmin -Cc $MailtToITAdmin -from $MailFrom -Subject $MailSubject  -SmtpServer $SmtpServer -Body $MailBody -Priority Low # -Attachments $outFile
}

Write-Host "$(Get-TimeStamp) End of Script"


Stop-transcript
