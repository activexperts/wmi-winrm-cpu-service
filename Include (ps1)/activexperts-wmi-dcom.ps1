###############################################################################
# ActiveXperts Network Monitor PowerShell script, © ActiveXperts Software B.V.
# For more information about ActiveXperts Network Monitor, visit the ActiveXperts 
# Network Monitor web site at https://www.activexperts.com/
###############################################################################

$IDXSERVICESHORTNAME = 0
$IDXSERVICELONGNAME = 1
$IDXPERFOBJECT = 0
$IDXPERFCONTEXT = 1
$IDXPERFITEM = 2
$IDXPERFCONDITION = 3

### Use activexperts.ps1 with common defines and functions
. 'C:\ProgramData\ActiveXperts\Network Monitor\Scripts\Include (ps1)\activexperts.ps1' 

##############################################################################################################################################################
# Function AxWmiCheckService
#  Parameter strCheckServiceName: can be either the service registry shortname (e.g.: AxNmSvc), or the service displayname (.e.g.: ActiveXperts Network Monitor')
##############################################################################################################################################################
function AxWmiCheckService( $strHost, $strCheckServiceName, $objWinCredentials, [ref]$strExplanation )
{
  if( $objWinCredentials -eq $Null )
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  else
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -Credential $objWinCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
  }

  if( $Error -ne '' )
  {
    $strExplanation.value = 'UNCERTAIN: ' + $Error
    return $false
  }
  
 
  $bServiceExists = $false
        

  # Check if service is running
  foreach( $objWmiSvc in $colWmiServices )
  {    
  
    if( ( $objWmiSvc.name -eq $strCheckServiceName ) -or ( $objWmiSvc.displayname -eq $strCheckServiceName ) ) 
    {
      $bServiceExists = $true
    
      if( $objWmiSvc.State.tolower() -ne "running" )
      {
        $strExplanation.value = 'ERROR: Service [' + $strCheckServiceName + '] is not running on [' + $strHost + ']'
        return $false
      }
    }
  }
  
  if( -not $bServiceExists )
  {
    $strExplanation.value = 'ERROR: Service [' + $strCheckServiceName + '] does not exist on [' + $strHost + ']'
    return $false        
  }
  
  return $true    
}

##############################################################################################################################################################
# Function AxWmiStopService
#  Parameter strServiceName: can be either the service registry shortname (e.g.: AxNmSvc), or the service displayname (.e.g.: ActiveXperts Network Monitor')
##############################################################################################################################################################
function AxWmiStopService( $strHost, $strServiceName, $nSecondsTimeout, $objWinCredentials, [ref]$strExplanation )
{
  $objWmiSvc = AxWmiGetService $strHost $strServiceName $objWinCredentials
  if( $objWmiSvc -eq $Null )
  {
    $strExplanation.value = 'ERROR: Service [' + $strServiceName + '] does not exist on [' + $strHost + ']'  
    return $false        
  }

  # Store result of StopService, otherwise these results will be returned too (together with $true $false
  $stopSvc = $objWmiSvc.StopService()
  
  $startDate = Get-Date

  do 
  { 
    $objWmiSvc = AxWmiGetService $strHost $strServiceName $objWinCredentials
    Start-Sleep -Seconds 1      
  } while( $objWmiSvc.State.tolower() -ne "stopped" -and $startDate.AddSeconds( $nSecondsTimeout ) -gt (Get-Date) )    
  
  if( $objWmiSvc.State.tolower() -ne "stopped" )
  {
    $strExplanation.value = 'ERROR: Service [' + $strServiceName + '] could not be stopped on [' + $strHost + ']'  
    return $false        
  }  
  
  $strExplanation.value = 'Success: Service [' + $strServiceName + '] successfully stopped on [' + $strHost + ']'    
  return $true    
}

##############################################################################################################################################################
# Function AxWmiStartService
#  Parameter strServiceName: can be either the service registry shortname (e.g.: AxNmSvc), or the service displayname (.e.g.: ActiveXperts Network Monitor')
##############################################################################################################################################################
function AxWmiStartService( $strHost, $strServiceName, $nSecondsTimeout, $objWinCredentials, [ref]$strExplanation )
{
  $objWmiSvc = AxWmiGetService $strHost $strServiceName $objWinCredentials
  if( $objWmiSvc -eq $Null )
  {
    $strExplanation.value = 'ERROR: Service [' + $strServiceName + '] does not exist on [' + $strHost + ']'  
    return $false        
  }

  $startSvc = $objWmiSvc.StartService()
  $startDate = Get-Date
  do 
  { 
    $objWmiSvc = AxWmiGetService $strHost $strServiceName $objWinCredentials
    Start-Sleep -Seconds 1      
  } while( $objWmiSvc.State.tolower() -ne "running" -and $startDate.AddSeconds( $nSecondsTimeout ) -gt (Get-Date) )    
  
  if( $objWmiSvc.State.tolower() -ne "running" )
  {
    $strExplanation.value = 'ERROR: Service [' + $strServiceName + '] could not be started on [' + $strHost + ']'  
    return $false        
  }  
  
  $strExplanation.value = 'Success: Service [' + $strServiceName + '] successfully started on [' + $strHost + ']'    
  return $true    
}

##############################################################################################################################################################
# Function AxWmiRebootComputer
##############################################################################################################################################################
function AxWmiRebootComputer( $strHost, $objWinCredentials, [ref]$strExplanation )
{
  $objWmiOS = AxWmiGetOS $strHost $objWinCredentials
  if( $objWmiOS -eq $Null )
  {
    $strExplanation.value = 'ERROR: Failed to retrieve OS information for computer [' + $strHost + ']'  
    return $false        
  }
  
  $rebootComputer = $objWmiOS.Reboot() 
  return $true
}    
  

##############################################################################################################################################################

function AxWmiGetService( $strHost, $strServiceName, $objWinCredentials)
{
 
  if( $objWinCredentials -eq $Null )
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  else
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -Credential $objWinCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  
           
  # Check if service is running
  foreach( $obj in $colWmiServices )
  {      
    if( ( $obj.name -eq $strServiceName ) -or ( $obj.displayname -eq $strServiceName ) ) 
    {
      return $obj
    }
  }
  
  return $Null   
}

##############################################################################################################################################################

function AxWmiGetOS( $strHost, $objWinCredentials)
{
 
  if( $objWinCredentials -eq $Null )
  {
    $colWmiComputers = Get-WmiObject -ComputerName $strHost -Class Win32_OperatingSystem -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  else
  {
    $colWmiComputers = Get-WmiObject -ComputerName $strHost -Class Win32_OperatingSystem -Credential $objWinCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  
           
  # Check if service is running
  foreach( $obj in $colWmiComputers )
  {     
    if( $obj.primary -eq $true ) 
    {
      return $obj
    }

  }
  
  return $Null   
}


##############################################################################################################################################################
# Function AxWmiCheckMultipleServices
# Check a list of services, to see if services are running. Services preceeded with a '!' are ignored (i.e. checking is disabled)
##############################################################################################################################################################
function AxWmiCheckMultipleServices( $strHost, [ref]$lstCheckServices, $objWinCredentials, [ref]$strExplanation )
{
  if( $objWinCredentials -eq $Null )
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  else
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Service -Credential $objWinCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
  }

  if( $Error -ne '' )
  {
    $strExplanation.value = 'UNCERTAIN: ' + $Error
    return $false
  }

  foreach( $checkService in $lstCheckServices.value )
  {
  
    $bCheck = $True
    if( $checkService[$IDXSERVICESHORTNAME] -eq '' -or $checkService[$IDXSERVICESHORTNAME].StartsWith('!') )
    {
      $bCheck = $false
    }
   
    if( -not $bCheck ) { continue; }
            
    $bServiceExists = $false
        
    # Check if service is running
    foreach( $objWmiSvc in $colWmiServices )
    {    
      if( ( $objWmiSvc.name -eq $checkService[$IDXSERVICESHORTNAME] ) -or ( $objWmiSvc.displayname -eq $checkService[$IDXSERVICELONGNAME] ) ) 
      {
        $bServiceExists = $true
      
        if( $objWmiSvc.State.tolower() -ne "running" )
        {
          $strExplanation.value = 'ERROR: Service [' + $checkService[$IDXSERVICESHORTNAME] + '] is not running on [' + $strHost + ']'
          return $false
        }
      }
    }
    
    if( -not $bServiceExists )
    {
      $strExplanation.value = 'ERROR: Service [' + $checkService[$IDXSERVICESHORTNAME] + '] does not exist on [' + $strHost + ']'
      return $false
    }
        
  }
  
  return $true    
}


##############################################################################################################################################################
# Function AxWmiCheckMultipleProcesses
# Check a list of processes, to see if processes are running. Processes preceeded with a '!' are ignored (i.e. checking is disabled)
##############################################################################################################################################################
function AxWmiCheckMultipleProcesses( $strHost, [ref]$lstCheckProcesses, $objWinCredentials, [ref]$strExplanation )
{
  if( $objWinCredentials -eq $null )
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Process -ErrorVariable Error -ErrorAction SilentlyContinue
  }
  else
  {
    $colWmiServices = Get-WmiObject -ComputerName $strHost -Class Win32_Process -Credential $objWinCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
  }

  if( $Error -ne '' )
  {
    $strExplanation.value = 'UNCERTAIN: ' + $Error
    return $false
  }
  
  foreach( $checkProcess in $lstCheckProcesses.value )
  {
    $bCheck = $true
    if( $checkProcess.StartsWith('!') )
    {
      $bCheck = $false
    }
    
    if( -not $bCheck ) { continue; }

    # Check if process exists
    $objWmiNameList = ( $colWmiServices | select name )
    
    $bFound = $false
    foreach( $processName in $objWmiNameList )
    {
      if( $processName -match $checkProcess )
      {
        $bFound = $true
      }
    }
    
    if( -not $bFound )
    {
      $strExplanation.value = 'ERROR: Process [' + $checkProcess + '] is not running on [' + $strHost + ']'
      return $false
    }
  }
  return $true    
}



##############################################################################################################################################################
# Function AxTcprecv
# Read a stream of strings from a TCP socket until there is no more data
##############################################################################################################################################################
function AxTcprecv( [ref]$strAllRespons )
{
  $strResponse = $null
  $bHasData = $null
      
  $bHasData = $objTcp.HasData()
  while( $bHasData )
  {
    $strResponse = $objTcp.ReceiveString()
    $bHasData = $objTcp.HasData()

    if( $strAllRespons.value -ne '' )
    {
      $strAllRespons.Value += '`r`n'
    }
    $strAllRespons.value += $strResponse
  }
}
  
##############################################################################################################################################################
# Function AxTcprecv
# Read a stream of strings from a TCP socket until there is no more data
##############################################################################################################################################################
function AxTcpsend( $strCommand )
{
  if( $strCommand -ne '' )
  {
    $objTcp.SendString( $strCommand.value )
  }
}

##############################################################################################################################################################
# Function AxCheckTelnet
# Read a stream of strings from a TCP socket until there is no more data
##############################################################################################################################################################
function AxCheckTelnet( $strServer, $numPort, $strCommand1, $strCommand2, $strCommand3, $strReceive, [ref]$strExplanation )
{
  # Define variables
  $strAllResponses = ''
   
  # Create object 
  $objTcp = new-object -comobject AxNetwork.Tcp
  
  # 1 means: raw, 2 means: telnet (see also ActiveXperts Network Component manual)
  $objTcp.Protocol = 2 

  $objTcp.Connect( $strServer, $numPort )
  if( $objTcp.LastError -ne 0 -or $objTcp.ConnectionState -ne 3 )
  {
    $objTcp.Disconnect() 
    $strExplanation.value = 'UNCERTAIN: Unable to connect to [' + $strServer + ']'
    return $false
  }

  #echo $objTcp.HasData()
  $objTcp.Sleep( 2000 ) # Allow some time

  AxTcprecv( [ref]$strAllResponses )

  AxTcpsend( [ref]$strCommand1 )
  AxTcprecv( [ref]$strAllResponses )

  AxTcpsend( [ref]$strCommand2 )
  AxTcprecv( [ref]$strAllResponses )

  AxTcpsend( [ref]$strCommand3 )
  AxTcprecv( [ref]$strAllResponses )

  #echo $strAllResponses
                  
  $objTcp.Disconnect()

  if( $strAllResponses.ToUpper() -match $strReceive.ToUpper() )
  {
    $strExplanation.value = 'SUCCESS: Response=[' + $strAllResponses + ']'
    return $true
  }  
  else
  {
    $strExplanation.value = 'ERROR: No response'
    return $false
  }  
}


##############################################################################################################################################################
# Function compareValue
# Compares 2 values with given operator
##############################################################################################################################################################

function compareValue( $strOperator, $nVMValue, $nUserValue )
{
  switch( $strOperator )
  {
    '-eq' { return ( $nVMValue -eq $nUserValue ) }
    '-lt' { return ( $nVMValue -lt $nUserValue ) }
    '-gt' { return ( $nVMValue -gt $nUserValue ) }
    '-ge' { return ( $nVMValue -ge $nUserValue ) }
    '-le' { return ( $nVMValue -le $nUserValue ) }
    '-ne' { return ( $nVMValue -ne $nUserValue ) }
  } 
}

##############################################################################################################################################################
# Function AxTranslateOperatorToText
# Return verbose operator
##############################################################################################################################################################

function AxTranslateOperatorToText( $strOperator )
{
  switch( $strOperator )
  {
    '-eq' { return 'EQUALS' }
    '-lt' { return 'LESSER THAN' }
    '-gt' { return 'GREATER THAN' }
    '-ge' { return 'GREATER OR EQUALS THAN' }
    '-le' { return 'LESSER OR EQUALS THAN' }
    '-ne' { return 'NOT EQUALS' }
  }
}

##############################################################################################################################################################
# Function AxIsCounterIDValid
# 
##############################################################################################################################################################
function AxIsCounterIDValid( $nCounterID )
{
  if( ( $nCounterID -eq 100 ) -or
      ( $nCounterID -eq 200 ) -or
      ( $nCounterID -eq 201 ) -or
      ( $nCounterID -eq 201 ) -or
      ( $nCounterID -eq 300 ) -or
      ( $nCounterID -eq 301 ) -or      
      ( $nCounterID -eq 350 ) -or      
      ( $nCounterID -eq 351 ) -or
      ( $nCounterID -eq 400 ) -or
      ( $nCounterID -eq 401 ) -or
      ( $nCounterID -eq 402 ) -or
      ( $nCounterID -eq 501 ) -or
      ( $nCounterID -eq 502 ) )
    {
      return $true
    }
  return $false
}

##############################################################################################################################################################
# Function AxIsOperatorValid
# 
##############################################################################################################################################################
function AxIsOperatorValid( $strOperator )
{
  if( ( $strOperator -eq '-eq' ) -or
      ( $strOperator -eq '-lt' ) -or
      ( $strOperator -eq '-gt' ) -or
      ( $strOperator -eq '-ge' ) -or
      ( $strOperator -eq '-le' ) -or
      ( $strOperator -eq '-ne' ) )
    {
      return 1
    }
  return 0
}
