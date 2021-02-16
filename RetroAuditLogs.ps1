#Modify these values to configure the audit log search.  Lower $intervalMinutes if results are exceeding 5000.
$resultSize = 5000
$intervalMinutes = 1440
$textSpeed = 50

#Windows Terminal Settings
#Font Provider: https://www.kreativekorp.com/software/fonts/apple2.shtml

#"initialCols": 160,
#"initialRows": 40,
#"profiles":
#[
#	{
#		// Retro Audit Logs
#		"guid": "{enter-your-own-guid-here}",
#		"name": "Retro Audit Logs",
#		"fontFace": "Print Char 21",
#		"experimental.retroTerminalEffect": true,
#		"cursorShape": "filledBox",
#		"useAcrylic": true,
#		"acrylicOpacity": 0.8,
#		"commandline": "powershell.exe",
#		"hidden": false
#	}
#]

Function Get-Logs {param($record,$days)
	clear
	if (Test-Path -Path $CurrPath\AuditLogTruncated.csv -PathType Leaf) { remove-item AuditLogTruncated.csv }
	if (Test-Path -Path $CurrPath\AuditLogRecords.csv -PathType Leaf) { remove-item AuditLogRecords.csv }
	write-host
	[DateTime]$start = [DateTime]::UtcNow.AddDays(-$days)
	[DateTime]$end = [DateTime]::UtcNow
	[DateTime]$currentStart = $start
	[DateTime]$currentEnd = $start
	"Retrieving audit records for the date range between $($start) and $($end)" | SlowText -Milliseconds $textSpeed
	"Days retrieved=$($days)" | SlowText -Milliseconds $textSpeed
	"Interval Minutes=$($intervalMinutes)" | SlowText -Milliseconds $textSpeed
	"RecordType=$($record)" | SlowText -Milliseconds $textSpeed
	"ResultsSize=$($resultSize)" | SlowText -Milliseconds $textSpeed
	write-host
	$totalCount = 0
	while ($true)
	{
		$currentEnd = $currentStart.AddMinutes($intervalMinutes)
		if ($currentEnd -gt $end)
		{
			$currentEnd = $end
		}

		if ($currentStart -eq $currentEnd)
		{
			break
		}

		$sessionID = [Guid]::NewGuid().ToString() + "_" +  "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
		"Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)" | SlowText -Milliseconds $textSpeed
		$currentCount = 0

		$sw = [Diagnostics.StopWatch]::StartNew()
		do
		{
			$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize

			if (($results | Measure-Object).Count -ne 0)
			{
				$results | export-csv -Path $outputFile -Append -NoTypeInformation
				$currentTotal = $results[0].ResultCount
				$totalCount += $results.Count
				$currentCount += $results.Count

				if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
				{
					"Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." | SlowText -Milliseconds $textSpeed
					""
					break
				} 
			}
		}
		while (($results | Measure-Object).Count -ne 0)
		$currentStart = $currentEnd
	}
	Import-Csv AuditLogRecords.csv | select RecordType,CreationDate,UserIds,Operations | ForEach-Object { 
		"{0},{1},{2},{3}" -f ($_.RecordType).Substring(0,[Math]::Min(30,($_.RecordType).Length)),$_.CreationDate,($_.UserIds).Substring(0,[Math]::Min(25,($_.UserIds).Length)),$_.Operations >> AuditLogTruncated.csv
	}
	write-host "Audit record retrieval complete, displaying results now:"  | SlowText -Milliseconds $textSpeed
	write-host
	Import-Csv AuditLogTruncated.csv | SlowText -Milliseconds $textSpeed
}

function SlowText{
	param([int]$Milliseconds= $textSpeed)
    $text = $input | Out-String 
    [char[]]$text | ForEach-Object{
        Write-Host -NoNewline $_
        if($_ -notmatch "\s"){Sleep -Milliseconds $Milliseconds}
    }
}

$host.ui.RawUI.WindowTitle = "Retro Audit Logs"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
$CurrPath = Get-Location
$outputFile = "$($CurrPath)\AuditLogRecords.csv"

Do {
# schema reference: https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-schema
	Get-Logs AzureActiveDirectory 2
	Get-Logs ExchangeAdmin 2
	Get-Logs SharePoint 2
	Get-Logs MicrosoftTeams 2
} Until (1 -le 0)






