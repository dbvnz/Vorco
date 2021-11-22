#############################################################################################################################
# Set the $ForwardToDestinations list to the numbers to cycle between for the call forward.                                 #
# Every time this script is run, it will automatically swap the on-call destination to the next number in the list.         #
# When it reaches the end of the list, it will then loop around to the beginning.                                           #
#                                                                                                                           #
# The punctuation is essential in this script, each list item should be on its own line, be surrounded by double quotes,    #
# and all items in the list except the last one must end with a comma. e.g.                                                 #
#                                                                                                                           #
# $ForwardToDestinations = @(                                                                                               #
#   "021987654",                                                                                                            #
#   "021234567",                                                                                                            #
#   "021765432"                                                                                                             #
# )                                                                                                                         #
#                                                                                                                           #
# In case the script fails, there is a catch block at the end where you can insert your own error reporting as appropriate. #
#############################################################################################################################

$ForwardToDestinations = @(

)
$apiKey = "" # Your Vorco API Key
$CentrexUserId = 0 # The User ID of the line to update the call forward destination for

##########################
# Do not edit below here #
##########################

$SecondsDelay = 300
$GetCompleted = $false
$CurrentForwardsResponse = $null
$GetRetryCount = 0
$GetRetries = 5
$apiBase = 'https://extranet.vorco.net/api/v1'

function Get-Token {
    $tokenHeaders = @{
        Authorization = "Token $apiKey"
    }

    try {
      Write-Host 'Logging in to Vorco Extranet'
      $tokenRequest = Invoke-WebRequest `
        -Uri "$apiBase/login" `
        -Method 'POST' `
        -Headers $tokenHeaders `
		-UseBasicParsing
      if ($tokenRequest -and $tokenRequest.StatusCode -eq 200) {
        $tokenObject = $tokenRequest.Content | ConvertFrom-Json
        return $tokenObject.access_token
      }
    } catch {
      Write-Error "Authentication failed"
      throw
    }
}

try {
  Write-Host "Retrieving call forward settings"
  while (-not $GetCompleted) {
    try {
      $token = Get-Token
      $headers = @{
        Authorization = "Bearer $token"
      }
      $CurrentForwardsResponse = Invoke-WebRequest `
        -Uri "$apiBase/centrex_users/$CentrexUserId/features/call_forwards" `
        -Method "GET" `
        -Headers $headers `
		-UseBasicParsing
      if ($CurrentForwardsResponse.StatusCode -ge 400) {
        throw "Expecting response code under 400"
      }
      $GetCompleted = $true
    } catch {
      if ($GetRetryCount -ge $GetRetries) {
        throw
      } else {
        Start-Sleep $SecondsDelay
        $GetRetryCount++
      }
    }
  }
  $CurrentForwards = ConvertFrom-Json $CurrentForwardsResponse.Content
  $CurrentForwardIndex = [array]::indexof($ForwardToDestinations,$CurrentForwards.call_forward_always.number)
  $NewForward = $ForwardToDestinations[0]
  if($CurrentForwardIndex -ne ($ForwardToDestinations.length - 1)) {
    $NewForward = $ForwardToDestinations[($CurrentForwardIndex+1)]
  }
  $SetRetryCount = 0
  $SetRetries = 180
  $SetCompleted = $false
  Write-Host "Updating call forward settings"
  while (-not $SetCompleted) {
    try {
      $SetResponse = Invoke-WebRequest `
        -Uri "$apiBase/centrex_users/$CentrexUserId/features/call_forwards" `
        -Method "PUT" `
        -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
        -Body "call_forward_always%5Bnumber%5D=$NewForward" `
        -Headers $headers `
		-UseBasicParsing
      if ($SetResponse.StatusCode -ge 400) {
        throw "Expecting response code under 400"
      }
      $SetCompleted = $true
      Write-Host "Successfully updated Call Forward for Centrex User $CentrexUserId to $NewForward"
    } catch {
      if ($SetRetryCount -ge $SetRetries) {
        throw
      } else {
        Start-Sleep $SecondsDelay
        $SetRetryCount++
      }
    }
  }
} catch {
  ########################################################################################
  # Set your own error reporting here, e.g. Email alert, syslog entry, log to disk, etc. #
  ########################################################################################
  Write-Error "Failed to update Call Forward for Centrex User $CentrexUserId"
  throw
}
