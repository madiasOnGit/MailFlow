# Specify the full path to EWS Installation
# Example: C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll
#
# Specify the Exchange Version
# This can be one of the following: Exchange2007_SP1 - Exchange2010 - Exchange2010_SP1 - Exchange2013

# EXO Example
#.\test_MailFlow_EWS.ps1 <user> <pass> Lisbon 120 <testAccount>@gmail.com Exchange2013 "https://outlook.office365.com/EWS/Exchange.asmx"

Param($User, $Pass, $location, $MaxWaitTime, $Recipients, $ExchangeVersion, $ServiceURL)


If ($User.Length -eq 0)
{    
	"MailFlow Script requires an Account"
	break	
}

Function GetItems($service)
{
	try
	{
		#Opens the Inbox and define to get 50 Items
		$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $User)
		$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderId)
		$ItemView = New-Object Microsoft.Exchange.WebServices.Data.itemView(50)

		#Finally really get the Data
		$GetItems = $null
		$GetItems = $service.FindItems($Inbox.Id, $ItemView)
    }
	catch
	{
		$_.Exception.Message
		break
	}

    Return $GetItems
}

Try
{
	Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
	$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $ExchangeVersion
	$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($User, $Pass)
	
	$service.Url = $ServiceURL
	#$service.AutodiscoverUrl("$O365User", {$true})

	$RandId = $null
	$RandId = [Guid]::NewGuid().ToString("D")
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
	$message.Subject = $env:computername + " " + $RandId
	$message.Body = 'This message is being sent through EWS with PowerShell. TimeStamp = ' + ([datetime]::Now).ToString("dd.MM.yyyy HH:mm:ss")
	$message.ToRecipients.Add($Recipients)

	#Create StopWatch to measure the time
	$sw = New-Object Diagnostics.Stopwatch
	$sw.Start()

	$message.Send()
}
Catch [Exception]
{
		$_.Exception.Message
		break
}

$MailFlow = $False

# Wait for Mail Echo to arrive
# If it takes longer than $MaxWaitTime, stop waiting and declare a fail

while($MailFlow -eq $False)
{
    $Data = GetItems($service)
    ForEach($Item in $Data)
    {
        If($Item.Subject -like "*$RandId*")
        {
            $MailFlow = $True
            $sw.Stop()
			$Item.Delete("HardDelete")
        }
    }

    If($sw.Elapsed.TotalSeconds -ge $MaxWaitTime)
    {
        $sw.Stop()
        break
    }
    Start-Sleep -Milliseconds 1000
}

If($MailFlow)
{
	"round trip from:" + $User + " to:  " + $Recipients + " toke: " +$sw.Elapsed.TotalSeconds +" sec"
	#$location + " " + $env:computername
}
Else
{
	"No Mail received back from " + $Recipients + " within " + $MaxWaitTime + " seconds. "
}
