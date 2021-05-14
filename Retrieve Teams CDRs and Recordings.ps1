########################################################################
# This script will retrieve Vorco Calling for Teams CDRs and recordings
# for the last x days, and zip them for archival.
########################################################################


#####################################################
# Update these variables as required
#
$apiKey = 'Your API Key'                                     # Vorco Extranet API Key
$teamsServiceId = 0                                          # Vorco Extranet MSCallService id
$outputDirectory = 'C:\Temp'                                 # Directory to leave .zip in once script completes
$outputFilenamePrefix = 'CDRs and Recordings'                # Prefix for .zip filename
$days = 32                                                   # Total days to retrieve CDRs/recordings for

#####################################################
# No changes are required below here                #
#####################################################

function Get-Token {
    $tokenHeaders = @{
        Authorization = "Token $apiKey"
    }

    try {
      Write-Host 'Logging in to Vorco Extranet'
      $tokenRequest = Invoke-WebRequest -Uri "$apiBase/login" -Method 'POST' -Headers $tokenHeaders
      if ($tokenRequest -and $tokenRequest.StatusCode -eq 200) {
        $tokenObject = $tokenRequest.Content | ConvertFrom-Json
        return $tokenObject.access_token
      }
    } catch {
      Write-Error "Authentication failed"
      throw
    }
}

function Get-CDRs {
    param(
        [Parameter(Mandatory)] $Token,
        [Parameter(Mandatory)] $StartDate,
        [Parameter(Mandatory)] $EndDate
    )
    $headers = @{
      Authorization = "Bearer $token"
    }

    try {
      Write-Host "Retrieving CDRs between $startDate and $endDate"
      $cdrsUri = "$apiBase/ms_call_services/$teamsServiceId/cdrs?start_date=$StartDate&end_date=$EndDate"
	  $ProgressPreference = 'SilentlyContinue'
      $cdrRequest = Invoke-WebRequest -UseBasicParsing -Uri $cdrsUri -Headers $headers
      if ($cdrRequest -and $cdrRequest.StatusCode -eq 200) {
        return $($cdrRequest.Content | ConvertFrom-Json)
      }
    } catch {
      Write-Error "Failed to retrieve CDRs"
      throw
    }
}

function Get-Recordings {
    param(
        [Parameter(Mandatory)] $Token,
        [Parameter(Mandatory)] $CDRs,
        [Parameter(Mandatory)] $OutputDir
    )
    $headers = @{
      Authorization = "Bearer $token"
    }
    try {
      Write-Host "Retrieving recordings for CDRs"
      foreach($cdr in $cdrs) {
        if ($cdr.recording_url) {
			Get-Recording -URL $cdr.recording_url -Headers $headers
        }
      }
    } catch {
      Write-Error "Failed to retrieve recordings"
      throw
    }
}

function Get-Recording {
	param(
		[Parameter(Mandatory)] $URL,
		[Parameter(Mandatory)] $Headers,
		$retries = 0
	)
	$maxRetries = 5
	try {
		Write-Host "Retrieving Recording: $URL"
		$ProgressPreference = 'SilentlyContinue'
		$recordingResponse = Invoke-WebRequest -UseBasicParsing -Uri $URL -Headers $headers
		if ($recordingResponse.Content -and $recordingResponse.Content.length -gt 0) {
			$contentDisposition = $recordingResponse.Headers.'Content-Disposition'
			$fileName = $contentDisposition.Split("=")[1].Replace("`"","")
			$path = Join-Path $OutputDir $fileName
			[io.file]::WriteAllBytes($path, $recordingResponse.Content)
			Write-Host "Saved $path"
		} else {
			Write-Error "Recording empty: $cdr.recording_url"
		}
	} catch {
		if ($retries -lt $maxRetries) {
			$retries++
			$sleepSeconds = [Math]::Pow(2, $retries + 2)
			Write-Host "Failed to retrieve recording $URL on attempt $retries of $maxRetries. Waiting $sleepSeconds seconds then trying again."
			Start-Sleep -Seconds $sleepSeconds
			Get-Recording -URL $URL -Headers $Headers -retries $retries
		} else {
			Write-Error "Max retries exceeded for recording $URL. Not retrieved, stopping processing."
			throw
		}
	}
}

if (!$(Test-Path $outputDirectory)) {
    Write-Error "`$outputDirectory ($outputDirectory) must exist."
    Exit
}

try {
    $apiBase = 'https://extranet.vorco.net/api/v1'
    $startTime = [int]$(Get-Date -UFormat %s)
    $tempDirPath = Join-Path $outputDirectory $('VorcoTeamsData-'+$startTime)
    New-Item -Path $tempDirPath -ItemType directory
    $endDate = Get-Date
    $startDate = $endDate.AddDays(-$days)
    $startDateStr = $startDate.ToString('yyyy-MM-dd')
    $endDateStr = $endDate.ToString('yyyy-MM-dd')
    $outputZip = Join-Path $outputDirectory $("$outputFilenamePrefix $startDateStr - $endDateStr ($startTime).zip")

    $token = Get-Token
    $cdrs = Get-CDRs -Token $token -StartDate $startDateStr -EndDate $endDateStr
    Get-Recordings -Token $token -CDRs $cdrs -OutputDir $tempDirPath
    $csvPath = $(Join-Path $tempDirPath 'CDRs.csv')
    $cdrs | Select-Object -Property id,a_party,b_party,timestamp,duration,price | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Compressing CDRs and recordings to $outputZip"
    Compress-Archive -Path $tempDirPath -DestinationPath $outputZip
} finally {
    # Delete temp dir contents and folder
    Write-Host 'Cleanup Temp Files'
    Get-ChildItem -Path $tempDirPath -Recurse | Remove-Item -force -recurse
    Remove-Item $tempDirPath -Force
}
