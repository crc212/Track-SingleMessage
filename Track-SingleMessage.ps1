cls

#Ask the questions#############################################################################
Write-Host "I'll search through the last x number of hours for the message, starting from now"
[string]$hours = Read-Host "How many hours back do you want to search?"
Write-Host " "
[string]$recipient = Read-Host "Who is the recipient? (full email address)"
Write-Host " "
[string]$sender = Read-Host "Who is the sender? (full email address)"
Write-Host " "
###############################################################################################

Write-Host Please wait - this can take a few minutes depending on your selected timeframe

[string]$negativeHours = "-" + $hours

$track = Get-TransportService | Get-MessageTrackingLog -Start (get-date).AddHours($negativeHours) -End (get-date) -Recipients $recipient -Sender $sender

if (($track | select -Unique messageid | Group-Object).Count -gt "0"){
    Write-Host Found ($track | select -Unique messageid | Group-Object).Count messages -ForegroundColor Green
    $resultTable = @()
    foreach ($result in $track){
        $resultEntry = New-Object PSObject
        $resultEntry | Add-Member NoteProperty -Name "ClientHostname" -Value $result.ClientHostname
        $resultEntry | Add-Member NoteProperty -Name "ClientIP" -Value $result.IP
        $resultEntry | Add-Member NoteProperty -Name "ServerHostname" -Value $result.ServerHostname
        $resultEntry | Add-Member NoteProperty -Name "ConnectorId" -Value $result.ConnectorId
        $resultEntry | Add-Member NoteProperty -Name "EventId" -Value $result.EventId
        $resultEntry | Add-Member NoteProperty -Name "EventData" -Value $result.EventData
        $resultEntry | Add-Member NoteProperty -Name "MessageId" -Value $result.MessageId
        $resultEntry | Add-Member NoteProperty -Name "Subject" -Value $result.MessageSubject
        $resultEntry | Add-Member NoteProperty -Name "Sender" -Value $result.Sender
        $resultEntry | Add-Member NoteProperty -Name "Recipients" -Value $result.Recipients
        $resultEntry | Add-Member NoteProperty -Name "Timestamp" -Value $result.Timestamp
        $resultEntry | Add-Member NoteProperty -Name "MessageLatency" -Value $result.MessageLatency
        
        $resultTable += $resultEntry 
    }
    
    
}

else {
    Write-Host "No results found" -ForegroundColor Yellow
}

if (($track | select -Unique messageid | Group-Object).Count -gt "0"){
foreach ($r in $resultTable){
    ### Edit the below line to look like your internal servers ###
    if ($r.ServerHostname -like "CONTOSO*"){
        if ($r.EventId -eq "SEND"){
            $latency = [math]::Round($r.MessageLatency.TotalSeconds,2)
            Write-Host "Delivered message internally to " -NoNewline
            Write-Host $r.Recipients -ForegroundColor Green -NoNewline
            Write-Host " with subject " -NoNewline
            Write-Host $r.Subject -ForegroundColor Green -NoNewline
            Write-Host " in " -NoNewline
            Write-Host $latency -ForegroundColor Green -NoNewline
            Write-Host " seconds"
        }
        
    }
    
    else{
        if ($r.EventId -eq "SEND"){
            $latency = [math]::Round($r.MessageLatency.TotalSeconds,2)
            Write-Host "Delivered message externally to " -NoNewline
            Write-Host $r.Recipients -ForegroundColor Green -NoNewline
            Write-Host " on server " -NoNewline
            Write-Host $r.ServerHostname -ForegroundColor Green -NoNewline
            Write-Host " with the subject " -NoNewline
            Write-Host $r.Subject -ForegroundColor Green -NoNewline
            Write-Host " in " -NoNewline
            Write-Host $latency -ForegroundColor Green -NoNewline
            Write-Host " seconds"
        }
    }
}
}

Start-Sleep -Seconds 1

if (($track | select -Unique messageid | Group-Object).Count -gt "0"){

$title = "Data Review"
$message = "Do you want to view detailed results in a table?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"

$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

switch ($result)
    {
        0 {$resultTable | Where-Object {$_.eventId -notlike "HA*"} | select timestamp,sender,recipients,subject,eventid,clienthostname,clientip,serverhostname,messagelatency,messageid | Sort-Object Timestamp | Out-GridView -Title "Message Tracking Search Results"}
        1 {Break}
    }



}