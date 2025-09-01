
<#
    Script: SharePoint Server Monitor with Email Alert
    Author: Ifeanyi Nwodo
    Description: Checks SharePoint server events(Database), Pools, Windows Services, Drive Space,
                 Connectivity  to SQL Server via TCP connection on port 1433 and Sends email alert if connection fails.
#>


$thresholdGB = 10
$thresholdBytes = $thresholdGB * 1GB
$drives = @("C:", "E:")
$sqlServer = "Sqlserver" 
$sqlPort = 1433 

$smtpServer = "ip address"
$smtpFrom = "info@from.com"
$smtpTo = "FirestReceiver@email.com","SecoundReceiver@email.com""
$smtpCc = "CcReceiver@email.com"
$smtpStorageCc = "CcReceiverStorage1@email.com","CcReceiverStorage2@email.com"
$smtpDBCc = "CcReceiverDB1@email.com","CcReceiverDB2@email.com"
$subjectEvent = "Critical! SharePoint Database Error Detected."
$subjectAppPool = "Critical! IIS App Pool Stopped: SharePoint - 80."
$subjectAppPool2 = "Critical! IIS App Pool Stopped: SecurityTokenServiceApplicationPool."
$subjectAppPool3 = "Critical! SharePoint Central Administration v4."
$subjectDrives = "Critical! SharePoint Server Drive Below ThreshHold."
$subjectSharePointSQl = "Critical! SharePoint  to SQL Server Connectivity Failed"

$serverName = $env:COMPUTERNAME
$ipAddress = (Test-Connection -ComputerName $serverName -Count 1).IPV4Address.IPAddressToString
$logPath = "C:\Program Files\SharePointMonitor\Logs\SharePointMonitor.log"
$logFile = "C:\\Logs\\SharePointServerCheck.log"


try{
New-Item -Path $logPath -ItemType File -Force | Out-Null
}catch{

}

function Save-Message($msg) {
    Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
}


function Send-NotificationEmail($subject, $body) {
    #Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
    Send-MailMessage -From $smtpFrom -To $smtpTo -Cc $smtpCc -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
}

function Send-DBNotificationEmail($subject, $body) {
    #Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
    Send-MailMessage -From $smtpFrom -To $smtpTo -Cc $smtpDBCc -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
   
}

function Send-StorageNotificationEmail($subject, $body) {
    #Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
    Send-MailMessage -From $smtpFrom -To $smtpTo -Cc $smtpStorageCc -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer
   
}

function Test-TcpPort {
    param (
        [string]$ComputerName,
        [int]$Port,
        [int]$Timeout = 3000  # Timeout in milliseconds
    )

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $waitHandle = $asyncResult.AsyncWaitHandle
        if ($waitHandle.WaitOne($Timeout, $false)) {
            $client.EndConnect($asyncResult)
            $client.Close()
            return $true
        } else {
            $client.Close()
            return $false
        }
    } catch {
        return $false
    }
}


$global:processedEvents = @{}
$global:lastAppPoolState = "Started"
$global:lastAppPoolState2 = "Started"
$global:lastAppPoolState3 = "Started"
$global:processedServices = @{}

$alertSent = @{}
foreach ($drive in $drives) {
    $alertSent[$drive] = $false
}


try{
   

while ($true) {


#Check Drives
foreach ($drive in $drives) {
 
    try {


        $psDrive = Get-PSDrive -Name $drive.TrimEnd(':')
        $freeSpace = $psDrive.Free

        if ($freeSpace -lt $thresholdBytes) {
            $freeGB = [math]::Round($freeSpace / 1GB, 2)
            $message = "Warning: Drive $drive has only $freeGB GB free space left."
             $body =@"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: Critical! SharePoint Server Drive Below ThreshHold.</h3>
  <p>
Server: <b>$serverName</b> <br/><br/>
IP Address: <b>$ipAddress</b> <br/><br/>
 $message
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>
</body>
</html>
"@
            Send-StorageNotificationEmail -subject $subjectDrives -body $body
            Save-Message "Drive threshhold alert sent for $drive of $serverName($ipAddress)"
            $alertSent[$drive] = $true
        }else{
        $alertSent[$drive] = $false
        }
    } catch {
        try{
             
             $eventSource = "DriveSpaceMonitor"
             $logName = "Application"

            # Check if the event source exists
            if (-not [System.Diagnostics.EventLog]::SourceExists($eventSource)) {
                try {
                    New-EventLog -LogName $logName -Source $eventSource
                } catch {
                    Write-Host "Failed to create event source: $eventSource"
                }
            }
                $errorMessage = "Error checking drive $drive"
                Write-Host $errorMessage




                # Log to Event Viewer
                try{
                Write-EventLog -LogName Application -Source "DriveSpaceMonitor" -EntryType Error -EventId 1001 -Message $errorMessage
                }catch{

                }


                } catch {

                }
            }
    }


#Monitor SharePoint server  to  SQL Server connectivity via Test-NetConnection on port 1433.
 try{
#$tcpTest = Test-NetConnection -ComputerName $sqlServer -Port $sqlPort
if (Test-TcpPort -ComputerName $sqlServer -Port $sqlPort -Timeout 3000) {
   # Write-Host "Connection successful"
} else {
#if (-not $tcpTest.TcpTestSucceeded) { 
     $body =@"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: SharePoint  to SQL Server Connectivity Failed.</h3>
  <p>
Server: <b>$serverName</b> <br/><br/>
IP Address: <b>$ipAddress</b> <br/><br/>
Time: <b>$(Get-Date)</b> <br/><br/>
SQL Server: <b>$sqlServer</b> <br/><br/>
SQL Port: <b>$sqlPort</b> <br/><br/>
Message: <b>$msg</b> <br/><br/>
Please investigate the connectivity issue
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>
</body>
</html>
"@

    Send-NotificationEmail -subject $subjectSharePointSQl -body $body
    Save-Message "SharePoint Server to SQL Server connectivity faillure alert sent for $drive of $serverName($ipAddress)"
    
} 


 }catch{

 }





    # Monitor Event Viewer
    $events = Get-WinEvent -LogName Application -MaxEvents 10 | Where-Object {
        ($_.LevelDisplayName -in @("Critical", "Error")) -and
        ($_.ProviderName -like "*SharePoint*") -and
        ($_.TaskDisplayName -like "*Database*")
    }

    foreach ($event in $events) {
        $eventKey = "$($event.TimeCreated.Ticks)-$($event.Id)"
        if (-not $global:processedEvents.ContainsKey($eventKey)) {
            $body =@"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: A critical or error event related to SharePoint Database has been detected.</h3>
  <p>
Server: <b>$serverName</b> <br/><br/>
IP Address: <b>$ipAddress</b> <br/><br/>
Time: <b>$($event.TimeCreated)</b> <br/><br/>
Source: <b>$($event.ProviderName)</b> <br/><br/>
Event ID: <b>$($event.Id)</b> <br/><br/>
Level: <b>$($event.LevelDisplayName)</b> <br/><br/>
Task Category: <b>$($event.TaskDisplayName)</b> <br/><br/>
Message: <b>$($event.Message)</b> <br/><br/>
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>
</body>
</html>
"@
            Send-DBNotificationEmail -subject $subjectEvent -body $body
            Save-Message "Event alert sent for Event ID $($event.Id)"
            $global:processedEvents[$eventKey] = $true
        }
    }

    # Monitor IIS App Pool
    Import-Module WebAdministration
    $appPoolName = "SharePoint - sharepoint80"
    $appPool = Get-Item "IIS:\AppPools\$appPoolName" -ErrorAction SilentlyContinue

    if ($appPool -and $appPool.state -ne $global:lastAppPoolState) {
        if ($appPool.state -eq "Stopped") {
            $body =  @"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: The IIS Application Pool '$appPoolName' has stopped.</h3>
  <p>
Server: </b>$serverName</b> <br/><br/>
IP Address: </b>$ipAddress</b> <br/><br/>
Time: </b>$(Get-Date)</b> <br/><br/>
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>

</body>
</html>
"@
            Send-NotificationEmail -subject $subjectAppPool -body $body
            Save-Message "App Pool alert sent for $appPoolName"
        }
        $global:lastAppPoolState = $appPool.state
    }

     #Import-Module WebAdministration
    $appPoolName2 = "SecurityTokenServiceApplicationPool"
    $appPool2 = Get-Item "IIS:\AppPools\$appPoolName2" -ErrorAction SilentlyContinue

    if ($appPool2 -and $appPool2.state -ne $global:lastAppPoolState2) {
        if ($appPool2.state -eq "Stopped") {
            $body2 = @"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: The IIS Application Pool '$appPoolName2' has stopped.</h3>
  <p>
Server: </b>$serverName</b> <br/><br/>
IP Address: </b>$ipAddress</b> <br/><br/>
Time: </b>$(Get-Date)</b> <br/><br/>
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>

</body>
</html>
"@

            Send-NotificationEmail -subject $subjectAppPool2 -body $body2
            Save-Message "App Pool alert sent for $appPoolName2"
        }
        $global:lastAppPoolState2 = $appPool2.state
    }


      #Import-Module WebAdministration
    $appPoolName3 = "SharePoint Central Administration v4"
    $appPool3 = Get-Item "IIS:\AppPools\$appPoolName3" -ErrorAction SilentlyContinue

    if ($appPool3 -and $appPool3.state -ne $global:lastAppPoolState3) {
        if ($appPool3.state -eq "Stopped") {
            $body3 = @"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: The IIS Application Pool '$appPoolName3' has stopped.</h3>
  <p>
Server: </b>$serverName</b> <br/><br/>
IP Address: </b>$ipAddress</b> <br/><br/>
Time: </b>$(Get-Date)</b> <br/><br/>
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>

</body>
</html>
"@

            Send-NotificationEmail -subject $subjectAppPool3 -body $body3
            Save-Message "App Pool alert sent for $appPoolName3"
        }
        $global:lastAppPoolState3 = $appPool3.state
    }




    # Monitor Windows Services
    $servicesToMonitor = @{
        "SPTimerV4"    = "SharePoint Timer Service"
        "SPAdminV4"  = "SharePoint Administration"
    }

    foreach ($serviceName in $servicesToMonitor.Keys) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service -and $service.Status -ne 'Running') {
            if (-not $global:processedServices.ContainsKey($serviceName)) {
                $body = @"
<html>
<head>
<style>
  body { font-family: Arial; }
  h3 { color: navy; }
</style>
</head>
<body>
  <h3>ALERT: $($servicesToMonitor[$serviceName]) is not running.</h3>
  <p>
Server: </b>$serverName</b> <br/><br/>
IP Address: </b>$ipAddress</b> <br/><br/>
Time: </b>$(Get-Date)</b> <br/><br/>
Status: </b>$($service.Status)</b> <br/><br/>
</p>
<br/>
<br/>
<p>
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AA'>
<br/><br/><br/>
Server | Monitor and Notification.<br/> 
<img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAm8AAA'>
</p>

</body>
</html>
"@

                Send-NotificationEmail -subject "Critical! Service Alert: $($servicesToMonitor[$serviceName])." -body $body
                Save-Message "Service alert sent for $($servicesToMonitor[$serviceName])"
                $global:processedServices[$serviceName] = $true
            }
        } elseif ($service -and $service.Status -eq 'Running') {
            if ($global:processedServices.ContainsKey($serviceName)) {
                $global:processedServices.Remove($serviceName)
            }
        }
    }

    Start-Sleep -Seconds 10
}
}catch{
    
     # Write the error message to a log file
    $errorMessage = "Error occurred: " + $_.Exception.Message
     #Log-MessageErr $errorMessage
    Add-Content -Path "error_log.log" -Value $errorMessage
# Display the full error message
    #Write-Host "Error message: $($_.Exception.Message)"
    
    # Optionally display more details
   # Write-Host "Full error details:"
   # Write-Host $_
}
