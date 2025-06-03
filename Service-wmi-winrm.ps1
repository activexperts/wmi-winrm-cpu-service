#################################################################################
# ActiveXperts Network Monitor PowerShell script, © ActiveXperts Software B.V.
# Script is based on WMI/WINRM.
# For more information about ActiveXperts, visit https://www.activexperts.com
#################################################################################
# Script
#     Service-wmi-winrm.ps1
# Description:
#     Checks if a service, specified by strCheckService, is running on the machine specified by strWinHost.
# Declare Parameters:
#     1) strWinHost (string) - Hostname or IP address of the Windows machine you want to check
#     2) bWinrmHttps (boolean) - $true to use WinRM secure (https), $false to use Winrm no secure (http)
#     3) nWinrmPort (int) - WinRM http(s) port number
#     4) strCheckService (string) - Name of the service
#     5) strConnectAs (string) - Specify an empty string to use Network Monitor service credentials.
#        To use alternate credentials, enter a server that is defined in Windows Machines credentials table.
#        (To define a Windows Machine entry, choose Tools->Options->Windows Machines)
# Usage:
#      .\Service-wmi-winrm.ps1 '<Hostname | IP>' <$true|$false> <Port> '<Service name>' '<Empty string | Windows Machine entry>' 
# Sample:
#      .\Service-wmi-winrm.ps1 'localhost' $false 5985 'spooler'
#################################################################################

### Declare Parameters
param( [string]$strWinHost = '', [boolean]$bWinrmHttps = $false, [int]$nWinrmPort = 5985, [string]$strCheckService = '', [string]$strConnectAs = '' )
  
### Use activexperts.ps1 with common functions 
. 'Include (ps1)\activexperts.ps1' 
. 'Include (ps1)\activexperts-wmi-winrm.ps1' 


#################################################################################
# // --- Main script ---
#################################################################################

### Clear error
$Error.Clear()

### Validate parameters, return on parameter mismatch
if( $strWinHost -eq '' -or $strCheckService -eq '' )
{
  echo 'UNCERTAIN: Invalid number of parameters - Usage: .\Service-wmi-winrm.ps1 ''<Hostname | IP>'' <$true|$false> <Port> ''<Service Name>'' ''[Empty String | Connect As]'''
  exit
}

### Declare local variables by assigning an initial value to it
$strAltLogin = ''
$strAltPassword = ''
$strExplanation = ''

# If alternate credentials are specified, retrieve the alternate login and password from the ActiveXperts global settings
if( $strConnectAs -ne '' )
{
  # Get the Alternate Credentials object. Function "AxGetCredentialInfo" is implemented in "activexperts.ps1"
  if( ( AxGetCredentialInfo $strWinHost $strConnectAs ([ref]$strAltLogin) ([ref]$strAltPassword) ([ref]$strExplanation) ) -ne $true )
  {
    echo $strExplanation
    exit
  }
}

$objWinrmSession = $null
if( ( AxWinrmCreateSession $strWinHost $bWinrmHttps $nWinrmPort $strAltLogin $strAltPassword ([ref]$objWinrmSession) ([ref]$strExplanation) ) -ne $true )
{ 
  echo $strExplanation
  exit  
}

### Check service
$strServiceState = ''
if( ( AxWinrmGetServiceInfo $objWinrmSession $strWinHost $strCheckService ([ref]$strServiceState) ([ref]$strExplanation) ) -ne $true )
{
  echo $strExplanation
  exit
}

if( $strServiceState -eq 'Running' )
{
  $strExplanation = 'SUCCESS: Service [' + $strCheckService + '] is running on [' + $strWinHost + ']'
}
else
{
  $strExplanation = 'ERROR: Service [' + $strCheckService + '] is not running on [' + $strWinHost + ']'
}


echo $strExplanation
exit




#################################################################################
# // --- Catch script exceptions ---
#################################################################################

trap [Exception]
{
  $strSourceFile = Split-Path $_.InvocationInfo.ScriptName -leaf
  $res = 'UNCERTAIN: Exception occured in ' + $strSourceFile + ' line #' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.Exception.Message
  echo $res
  exit
}