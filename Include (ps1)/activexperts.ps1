###############################################################################
# ActiveXperts Network Monitor PowerShell script, © ActiveXperts Software B.V.
# For more information about ActiveXperts Network Monitor, visit the ActiveXperts 
# Network Monitor web site at https://www.activexperts.com/
###############################################################################



##############################################################################################################################################################
# Function AxGetCredentialObject
# Retrieve alternate credentials, as defined in Manager's  Tools->Options->Server Credentials-tab
##############################################################################################################################################################
function AxGetCredentialObject( $strHost, $strConnectAs, [ref]$objWinCredentials, [ref]$strExplanation )
{  
  $objWinCredentials.value = $null
  if( $strConnectAs -eq '' )
  {
    # No alternate credentials specified, so login and password are empty and service credentials will be used
    return $true
  }
  
  $objNmCredentials = new-object -comobject ActiveXperts.NMServerCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
 
  if( $error -ne '' ) 
  {
    $strExplanation.value = $error
    return $false
  }
  $strAltLogin = $objNmCredentials.GetLogin( $strConnectAs )
  $strAltPassword = $objNmCredentials.GetPassword( $strConnectAs )
  
  if( $strAltLogin -eq '' )
  {
    $strExplanation.value = 'UNCERTAIN: No alternate credentials defined for [' + $strConnectAs + ']. In the Manager application, select <Options> from the <Tools> menu and select the <Server Credentials> tab to enter alternate credentials'
    return $false
  }
  
  $strAltPasswordSecure = ConvertTo-SecureString -string $strAltPassword -AsPlainText -Force
  $objWinCredentials.Value = new-object -typename System.Management.Automation.PSCredential $strAltLogin, $strAltPasswordSecure  
  
  return $true
}


##############################################################################################################################################################
# Function AxGetCredentialInfo
# Retrieve alternate credentials, as defined in Manager's  Tools->Options->Server Credentials-tab
##############################################################################################################################################################
function AxGetCredentialInfo( $strHost, $strConnectAs, [ref]$strAltLogin, [ref]$strAltPassword, [ref]$strExplanation )
{  
  $strAltLogin.value = ''
  $strAltPassword.value = ''
  
  if( $strConnectAs -eq '' )
  {
    # No alternate credentials specified, so login and password are empty and service credentials will be used
    return $true
  }
  
  $objNmCredentials = new-object -comobject ActiveXperts.NMServerCredentials -ErrorVariable Error -ErrorAction SilentlyContinue
 
  if( $error -ne '' ) 
  {
    $strExplanation.value = $error
    return $false
  }
  $strAltLogin.value = $objNmCredentials.GetLogin( $strConnectAs )
  $strAltPassword.value = $objNmCredentials.GetPassword( $strConnectAs )
    
  return $true
}



##############################################################################################################################################################
# Function AxCheckMultiplePerfCounters
# Check a list of performance counters, to see if performance counters are OK. Performance counters preceeded with a '!' are ignored (i.e. checking is disabled)
##############################################################################################################################################################
function AxCheckMultiplePerfCounters( $strHost, $lstCheckPerfCounters, $objWinCredentials, [ref]$strExplanation )
{
  if( $objWinCredentials -ne $null )
  {
    $strAltUserName = $objWinCredentials.Username
    $strAltPassword = $objWinCredentials.Password
  }
  else
  {
    $strAltUserName = ''
    $strAltPassword = ''
  }

  $objPerf = new-object -comobject ActiveXperts.NMPerf
  $objPerf.Initialize('')
    
  $objPerf.Connect( $strHost, $strAltUserName, $strAltPassword )  
  
  foreach( $arrCheckPerfCounter in $lstCheckPerfCounters.value )
  {
    if( ( $arrCheckPerfCounter[$IDXPERFOBJECT] -ne '' ) -and ( -not $arrCheckPerfCounter[$IDXPERFOBJECT].StartsWith('!') ) )
    {
      $strPerfPath = $objPerf.BuildPath( $strHost, $arrCheckPerfCounter[$IDXPERFOBJECT], $arrCheckPerfCounter[$IDXPERFCONTEXT], $arrCheckPerfCounter[$IDXPERFITEM] )
      $numPerfValue = $objPerf.GetIntegerValue( $strPerfPath )

      if( $objPerf.LastError -ne 0 )
      {
        $strExplanation.value  = 'ERROR: Failed to retrieve value for counter [' + $strPerfPath + ']'
        return $false
      }
      
      # [string] is necessary otherwise Powershell tries to convert the variables to an integer
      $strEval = [string]$numPerfValue + ' ' + [string]$arrCheckPerfCounter[$IDXPERFCONDITION]

      $bCompareResult = invoke-expression -Command $strEval 
      
      if( -not $bCompareResult )
      {
        $strExplanation.value  = 'ERROR: Path [' + $strPerfPath + '], Condition[' + $arrCheckPerfCounter[$IDXPERFITEM] + $arrCheckPerfCounter[$IDXPERFCONDITION] + '] failed, Current Value=[' + $numPerfValue + ']'
        return $false
      }
    }
  }
  
  $objPerf.Shutdown()

  $strExplanation.value = "SUCCESS: Performance counters checked"
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
