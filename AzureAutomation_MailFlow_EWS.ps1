
#EWS
$MaxWaitTime = 120
$Recipient = "...@gmail.com"
$cred = Get-AutomationPSCredential -Name "<credential>"
$mailbox = "....@outlook.com"
$serviceUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
$location = "Lisbon"

#OMS
$customerId = Get-AutomationVariable -Name '<WorkspaceID>'
$SharedKey = Get-AutomationVariable -Name '<PrimaryKey>'
$dataType = "MailRoundTrip"
$TimeStampField = "DateTime"
#$TimeStampField = [DATETIME]::Now

# Function to create the authorization signature - TECHNET example
Function New-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource)
{
  $xHeaders = 'x-ms-date:' + $date
  $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

  $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
  $keyBytes = [Convert]::FromBase64String($sharedKey)

  $sha256 = New-Object -TypeName System.Security.Cryptography.HMACSHA256
  $sha256.Key = $keyBytes
  $calculatedHash = $sha256.ComputeHash($bytesToHash)
  $encodedHash = [Convert]::ToBase64String($calculatedHash)
  $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
  return $authorization
}

#Send data to OMS
Function Send-OMSData($customerId, $sharedKey, $body, $logType)
{
  $method = 'POST'
  $contentType = 'application/json'
  $resource = '/api/logs'
  $rfc1123date = [DateTime]::UtcNow.ToString('r')
  $contentLength = $body.Length
  $signature = New-Signature `
  -customerId $customerId `
  -sharedKey $sharedKey `
  -date $rfc1123date `
  -contentLength $contentLength `
  -fileName $fileName `
  -method $method `
  -contentType $contentType `
  -resource $resource
  $uri = 'https://' + $customerId + '.ods.opinsights.azure.com' + $resource + '?api-version=2016-04-01'

  $headers = @{
    'Authorization'      = $signature
    'Log-Type'           = $logType
    'x-ms-date'          = $rfc1123date
    'time-generated-field' =  $TimeStampField
  }

  $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
  return $response.StatusCode
}


Try
{

    import-module EWS
    $service = Connect-EWSService -Mailbox $mailbox -ServiceUrl $serviceUrl -Credential $cred 
	
    #create message unique id
    $mID = [Guid]::NewGuid().ToString("D")
    $body = 'This message is being sent through EWS with Azure Automation. TimeStamp = ' + ([datetime]::Now).ToString("dd.MM.yyyy HH:mm:ss")

	#Create StopWatch to measure the time and send test email
	$sw = New-Object Diagnostics.Stopwatch
	$sw.Start()
    $m = New-EWSMessage -To $Recipient -Subject $mID -Body $body 
    

}
Catch [Exception]
{
        $RoundTripData = @{
			'TestExecution' = 'Azure'
            'Location' = $location
			'From' = $mailbox
			'To' = $Recipient
			'Result' = $("Failed:" + $_.Exception.Message)
			'RoundTripTime' = 0
	    }
		
}



# Wait for Mail to arrive - If it takes longer than $MaxWaitTime, stop waiting and declare a fail
if($m)
{
    $roundTrip= $False
    while($roundTrip -eq $False)
    {
        $message = Get-EWSFolder Inbox | Get-EWSItem -Filter subject:$mID
        If($message)
        {
            $roundTrip = $True
            $sw.Stop()
            $message.Delete("HardDelete")
            $RoundTripData = @{
                'TestExecution' = 'Azure'
                'Location' = $location
                'From' = $mailbox
                'To' = $Recipient
                'Result' = 'Success'
                'RoundTripTime' = $($sw.Elapsed.TotalSeconds)
            }
        }

        If($sw.Elapsed.TotalSeconds -ge $MaxWaitTime)
        {
            $sw.Stop()
            $RoundTripData = @{
                'TestExecution' = 'Azure'
                'Location' = $location
                'From' = $mailbox
                'To' = $Recipient
                'Result' = 'Failed:TimeOut'
                'RoundTripTime' = $MaxWaitTime
            }
            break
        }

        Start-Sleep -Milliseconds 1000
    }
  
}



$payload =  $RoundTripData | ConvertTo-Json

#send Data to OMS
Send-OMSData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($payload)) -logType $dataType


