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
# Function AxWinrmCreateSession
# Retrieve alternate credentials, as defined in Manager's  Tools->Options->Server Credentials-tab
##############################################################################################################################################################
function AxWinrmCreateSession( $strHost, $bWinrmHttps, $nWinrmPort, $strAltLogin, $strAltPassword, [ref]$objWinrmSession, [ref]$strExplanation )
{  
  $objWsMan= new-object -comobject Wsman.Automation -ErrorVariable Error -ErrorAction SilentlyContinue
 
  if( $error -ne '' ) 
  {
    $strExplanation.value = $error
    return $false
  }
  
  # Create strUrl like: https://myserver01:5986/wsman  
  if( $bWinrmHttps -eq $true )
  {
    $strUrl = 'https://'
  }
  else
  {
    $strUrl = 'http://'
  }
  $strUrl = $strUrl + $strHost + ':' + $nWinrmPort + '/wsman'

  if( $strAltLogin -ne '' )
  {
    $objConnOptions = $null
    $objConnOptions = $objWsman.CreateConnectionOptions()
    if( $error -ne '' ) 
    {
      $strExplanation.value = $error
      return $false
    }
  
    $objConnOptions.UserName = $strAltLogin
    $objConnOptions.Password = $strAltPassword
       
    $objWinRMSession.value = $objWsman.CreateSession( $strUrl, 4096, $objConnOptions )
  }
  else
  {
    $objWinRMSession.value = $objWsman.CreateSession( $strUrl )  
  }
  
  if( $error -ne '' ) 
  {
    $strExplanation.value = $error
    return $false
  }
  if( $objWinRMSession.value -eq $null ) 
  {
    $strExplanation.value = $error
    echo 'Null in de functie...'
    return $false
  }  
 
  $strExplanation.value = ''
  return $true
}

##############################################################################################################################################################
# Function AxWinrmCheckCpu
##############################################################################################################################################################
function AxWinrmGetCpu( $objWinrmSession, $strHost, $strFindCpu, [ref]$nRefCpuProcentProcTime, [ref]$strRefExplanation )
{
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colCPUs = $null
  $colCPUs = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_PerfFormattedData_PerfOS_Processor', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colCPUs.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No CPUs on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colCPUs.AtEndOfStream )
  {
    $strResponse = $colCPUs.ReadItem()
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strCpuName             = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    if( $strCpuName -eq $strFindCpu )
    {
      $strCpuProcentProcTime  = $objNMUtilities.WinRMGetAttributeValue( "PercentProcessorTime" )
      $nRefCpuProcentProcTime.value    = [int]$strCpuProcentProcTime
      
      return $true
    }       
  }
 
  $strRefExplanation.value = 'UNCERTAIN: CPU [' + $strFindCpu + '] does not exist on [' + $strHost + ']'
  return $false
}

##############################################################################################################################################################
# Function AxWinrmGetDiskDrives
##############################################################################################################################################################
function AxWinrmGetDiskDrives( $objWinrmSession, $strHost, [ref]$strRefGoodDisks, [ref]$strRefBadDisks, [ref]$strRefExplanation )
{
  $strRefGoodDisks.value    = ''
  $strRefBadDisks.value     = ''
  $strRefExplanation.value  = ''
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colDiskDrives = $null
  $colDiskDrives = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_DiskDrive', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colDiskDrives.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No diskdrives on windows machine [' + $strHost + ']'
    return $false
  }  

  $strDisks = ''
  while( ! $colDiskDrives.AtEndOfStream )
  {
    $strResponse = $colDiskDrives.ReadItem()
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strDiskCaption = $objNMUtilities.WinRMGetAttributeValue( 'Caption' )
    $strDiskStatus = $objNMUtilities.WinRMGetAttributeValue( 'Status' )
      
    if( $strDiskStatus -eq 'OK' ) 
    {
      if( $strRefGoodDisks.value -ne '' ) 
      {
        $strRefGoodDisks.value = $strRefGoodDisks.value + ','
      }
      $strRefGoodDisks.value  = $strRefGoodDisks.value + $strDiskCaption    
    }
    else
    {
      if( $strRefBadDisks -ne '' ) 
      {
        $strRefBadDisks.value = $strRefBadDisks.value + ','
      }
      $strRefBadDisks.value  = $strRefBadDisks.value 	+ $strDiskCaption

    }
  }
  
  return $true    
}


##############################################################################################################################################################
# Function AxWinrmGetDiskSpace
##############################################################################################################################################################
function AxWinrmGetDiskSpace( $objWinrmSession, $strHost, $strCheckDriveLetter, [ref]$nRefDiskTotalSizeGB, [ref]$nRefDiskFreeSpaceGB, [ref]$nRefDiskUsedSpaceGB, [ref]$strRefExplanation )
{
  $nRefDiskTotalSizeGB.value = 0
  $nRefDiskFreeSpaceGB.value = 0
  $nRefDiskUsedSpaceGB.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colDisks = $null
  $colDisks = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_LogicalDisk', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colDisks.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No logical disks on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colDisks.AtEndOfStream )
  {
    $strResponse = $colDisks.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strDiskName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strDiskFreeSpace = $objNMUtilities.WinRMGetAttributeValue( 'FreeSpace' )
    $strDiskSize = $objNMUtilities.WinRMGetAttributeValue( 'Size' )
    
    if( $strDiskName -eq $strCheckDriveLetter )
    {
      $nRefDiskFreeSpaceGB.value  = [math]::Round( ([long]$strDiskFreeSpace) / ( 1024 * 1024 * 1024 ) )  
      $nRefDiskTotalSizeGB.value  = [math]::Round( ([long]$strDiskSize) / ( 1024 * 1024 * 1024 ) )
      $nRefDiskUsedSpaceGB.value  = $nRefDiskTotalSizeGB.value - $nRefDiskFreeSpaceGB.value 
            
      return $true
    }
  }
       
  $strRefExplanation.value = 'UNCERTAIN: Disk [' + $strCheckDriveLetter + '] was not found on [' + $strHost + ']'       
  return $false    
}

##############################################################################################################################################################
# Function AxWinrmGetMountPointDiskSpace
##############################################################################################################################################################
function AxWinrmGetMountPointDiskSpace( $objWinrmSession, $strHost, $strCheckMountPoint, [ref]$nRefDiskFreeSpaceGB, [ref]$strRefExplanation )
{
 
  $nRefDiskFreeSpaceGB.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colDisks = $null
  $colDisks = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_PerfFormattedData_PerfDisk_LogicalDisk', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colDisks.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No logical disks on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colDisks.AtEndOfStream )
  {
    $strResponse = $colDisks.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strDiskName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strDiskFreeSpaceMB = $objNMUtilities.WinRMGetAttributeValue( 'FreeMegabytes' )   
    
    if( $strDiskName -eq $strCheckMountPoint )
    {
      $nRefDiskFreeSpaceGB.value  = [math]::Round( ([long]$strDiskFreeSpaceMB) / ( 1024 ) )  
            
      return $true
    }
  }
       
  $strRefExplanation.value = 'UNCERTAIN: Disk [' + $strCheckMountPoint + '] was not found on [' + $strHost + ']'       
  return $false    
}

##############################################################################################################################################################
# Function AxWinrmGetMemory
##############################################################################################################################################################
function AxWinrmGetMemory( $objWinrmSession, $strHost, [ref]$nRefAvailableMemoryMB, [ref]$nRefCommittedMemoryMB, [ref]$nRefPagesPerSecond, [ref]$strRefExplanation )
{
  $nRefAvailableMemoryMB.value = 0
  $nRefCommittedMemoryMB.value = 0
  $nRefPagesPerSecond.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colItems = $null
  $colItems = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_PerfFormattedData_PerfOS_Memory', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colItems.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No logical disks on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colItems.AtEndOfStream )
  {
    $strResponse = $colItems.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strAvailableBytes  = $objNMUtilities.WinRMGetAttributeValue( 'AvailableBytes' )
    $strCommittedBytes  = $objNMUtilities.WinRMGetAttributeValue( 'CommittedBytes' )
    $strPagesPerSec     = $objNMUtilities.WinRMGetAttributeValue( 'PagesPerSec' )
    
    $nRefAvailableMemoryMB.value       = [math]::Round( ([long]$strAvailableBytes) / ( 1024 * 1024 ) ) 
    $nRefCommittedMemoryMB.value       = [math]::Round( ([long]$strCommittedBytes) / ( 1024 * 1024 ) )
    $nRefPagesPerSecond.value          = [long]$strPagesPerSec
    
    return $true
  }
       
  $strRefExplanation.value = 'UNCERTAIN: No memory information available'       
  return $false    
}


##############################################################################################################################################################
# Function AxWinrmNetworkInterfaceInfo
##############################################################################################################################################################
function AxWinrmNetworkInterfaceInfo( $objWinrmSession, $strHost, $strCheckNicName, [ref]$nRefTotalKBytesPerSec, [ref]$nRefBandwidthMbitSec, [ref]$strRefExplanation )
{
  $nRefTotalKBytesPerSec.value = 0
  $nRefBandwidthMbitSec.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colNICs = $null
  $colNICs = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_PerfFormattedData_Tcpip_NetworkInterface', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colNICs.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No NICs on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colNICs.AtEndOfStream )
  {
    $strResponse = $colNICs.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strNicName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strNicBytesTotalPerSec = $objNMUtilities.WinRMGetAttributeValue( 'BytesTotalPerSec' )
    $strNicBandwidthMbitSec = $objNMUtilities.WinRMGetAttributeValue( 'CurrentBandwidth' )
    if( $strNicName -eq $strCheckNicName )
    {
      $nRefTotalKBytesPerSec.value  = [math]::Round( ([long]$strNicBytesTotalPerSec) / ( 1024 ) )  
      $nRefBandwidthMbitSec.value  = [math]::Round( ([long]$strNicBandwidthMbitSec) / ( 1000 * 1000 ) )
            
      return $true
    }
  }
       
  $strRefExplanation.value = 'UNCERTAIN: NIC [' + $strCheckNicName + '] was not found on [' + $strHost + ']'       
  return $false    
}

##############################################################################################################################################################
# Function AxWinrmGetPrinterInfo
##############################################################################################################################################################

function getPrinterStatusString($numPS)
{
  switch( $numPS )
  {
    0 { $strPS= 'Undefined' }
    1 { $strPS= 'Other' }
    2 { $strPS= 'Unknown' }
    3 { $strPS= 'Idle' }
    4 { $strPS= 'Printing' }
    5 { $strPS= 'Warmup' }
    6 { $strPS= 'Stopped Printing' }
    7 { $strPS= 'Offline' }
    default { $strPS = 'Unknown (' + $numDES + ')' }
  }
  
  return $strPS
}

function getPrinterErrorStateString($numDES)
{
  switch( $numDES )
  {
    0  { $numDES= 'Unknown' }
    1  { $numDES= 'Other' }
    2  { $numDES= 'No Error' }
    3  { $numDES= 'Low Paper' }
    4  { $numDES= 'No Paper' }
    5  { $numDES= 'Low Toner' }
    6  { $numDES= 'No Toner' }
    7  { $numDES= 'Door Open' }
    8  { $numDES= 'Jammed' }
    9  { $numDES= 'Offline' }
    10 { $numDES= 'Service Requested' }
    11 { $numDES= 'Output Bin Full' }
    default { $numDES = 'Unknown (' + $numDES + ')' }
  }
  
  return $numDES
}
function AxWinrmGetPrinterInfo( $objWinrmSession, $strHost, $strCheckPrinterName, [ref]$bRefIsPrinterUp, [ref]$strRefPrinterStatus, [ref]$strRefPrinterErrorState, [ref]$strRefExplanation )
{
  $bRefIsPrinterUp.value = $false
  $strRefPrinterStatus.value = ''
  $strRefPrinterErrorState.value = ''
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colPrinters = $null
  $colPrinters = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_Printer', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colPrinters.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No NICs on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colPrinters.AtEndOfStream )
  {
    $strResponse = $colPrinters.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strPrinterName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strPrinterStatus = $objNMUtilities.WinRMGetAttributeValue( 'PrinterStatus' )
    $strPrinterDetectedErrorState = $objNMUtilities.WinRMGetAttributeValue( 'DetectedErrorState' )
    if( $strPrinterName -eq $strCheckPrinterName )
    {
      $nPrinterStatus  = ([int]$strPrinterStatus)
      $nPrinterDetectedErrorState  = ([int]$strPrinterDetectedErrorState)
              
      if( $nPrinterStatus -eq 3 -or $nPrinterStatus -eq 4 -or $nPrinterStatus -eq 5 )
      {
        $bRefIsPrinterUp.value = $true
      }
      else 
      {
        $bRefIsPrinterUp.value = $false
      }
      
      $strRefPrinterStatus.value  = getPrinterStatusString( $nPrinterStatus )
      $strRefPrinterErrorState.value  = getPrinterErrorStateString( $nPrinterDetectedErrorState )
      
      return $true
    }
  }
       
  $strRefExplanation.value = 'UNCERTAIN: Printer [' + $strCheckPrinterName + '] was not found on [' + $strHost + ']'       
  return $false    
}

##############################################################################################################################################################
# Function AxWinrmGetProcessCount
##############################################################################################################################################################

function AxWinrmGetProcessCount( $objWinrmSession, $strHost, $strCheckProcessName, [ref]$nRefProcessCount, [ref]$strRefExplanation )
{
  $nRefProcessCount.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colProcesses = $null
  $colProcesses = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_Process', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colProcesses.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No services on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colProcesses.AtEndOfStream )
  {
    $strResponse = $colProcesses.ReadItem()
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strProcessName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
      
    if( $strProcessName -eq $strCheckProcessName )
    {
      $nRefProcessCount.value = $nRefProcessCount.value + 1
    }
  }
   
  return $true    
}


##############################################################################################################################################################
# Function AxWinrmGetQueueInfo
##############################################################################################################################################################
function AxWinrmGetQueueInfo( $objWinrmSession, $strHost, $strCheckQueueName, [ref]$nRefQueueLength, [ref]$nRefKBytesInQueue, [ref]$strRefExplanation )
{
  $nRefQueueLength.value = 0
  $nRefKBytesInQueue.value = 0
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colQueues = $null
  $colQueues = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_PerfRawData_MSMQ_MSMQQueue', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colQueues.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No NICs on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colQueues.AtEndOfStream )
  {
    $strResponse = $colQueues.ReadItem()
    
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    # NOTE: $objQueue.Name is formatted like this: Server01\Private$\AdminQueue$
    # However, we cannot use Server01 because the user may have specified localhost, or 192.168.1.1 as computername
    # So now, filter the queuename by stripping the 
    $strFullQueueName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $nPos = $strFullQueueName.IndexOf("\")
    if( $nPos -ge 0 )
    {
      $strQueueName = $strFullQueueName.substring( $nPos + 1 )
    }
    else
    {
      $strQueueName = $strFullQueueName   
    }

    $strMessagesInQueue = $objNMUtilities.WinRMGetAttributeValue( 'MessagesInQueue' )
    $strBytesInQueue = $objNMUtilities.WinRMGetAttributeValue( 'BytesInQueue' )
    
    if( $strQueueName -eq $strCheckQueueName )
    {
      $nRefQueueLength.value  = ( [int]$MessagesInQueue )  
      $nRefKBytesInQueue.value  = [math]::Round( ([long]$strBytesInQueue) / ( 1024) )
            
      return $true
    }
  }
       
  $strRefExplanation.value = 'UNCERTAIN: Queue [' + $strCheckQueueName + '] was not found on [' + $strHost + ']'       
  return $false    
}


##############################################################################################################################################################
# Function AxWinrmGetServiceInfo
#  Parameter strCheckServiceName: can be either the service registry shortname (e.g.: AxNmSvc), or the service displayname (.e.g.: ActiveXperts Network Monitor')
##############################################################################################################################################################

function AxWinrmGetServiceInfo( $objWinrmSession, $strHost, $strCheckServiceName, [ref]$strRefServiceState, [ref]$strRefExplanation )
{
  $strRefServiceState.value = ''
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colServices = $null
  $colServices = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_Service', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colServices.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No services on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colServices.AtEndOfStream )
  {
    $strResponse = $colServices.ReadItem()
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strServiceName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strServiceDisplayName = $objNMUtilities.WinRMGetAttributeValue( 'DisplayName' )
    $strServiceState = $objNMUtilities.WinRMGetAttributeValue( 'State' )
      
    if( ( $strServiceName -eq $strCheckServiceName ) -or ( $strServiceDisplayName -eq $strCheckServiceName ) ) 
    {
      $strRefServiceState.value = $strServiceState
      return $true
    }
  }
 
  $strRefExplanation.value = 'UNCERTAIN: Service [' + $strCheckServiceName + '] not found on [' + $strHost + ']'  
  
  return $false    
}

##############################################################################################################################################################
# Function AxWinrmGetShareStatus
##############################################################################################################################################################

function AxWinrmGetShareStatus( $objWinrmSession, $strHost, $strCheckShareName, [ref]$strRefShareStatus, [ref]$strRefExplanation )
{
  $strRefShareStatus.value = ''  
  $strRefExplanation.value = ''
  
  $objNMUtilities = $null
  $objNMUtilities = new-object -comobject ActiveXperts.NMUtilities

  $strResource = 'http://schemas.microsoft.com/wbem/wsman/1/wmi/root/cimv2/*'

  $colShares = $null
  $colShares = $objWinRMSession.Enumerate( $strResource, 'Select * from Win32_Share', 'http://schemas.microsoft.com/wbem/wsman/1/WQL' )
  if( $error -ne '' ) 
  {
    $strRefExplanation.value = 'UNCERTAIN: ' + $error
    return $false
  }

  if( $colShares.AtEndOfStream  )
  {
    $strRefExplanation.value = 'UNCERTAIN: No shares on windows machine [' + $strHost + ']'
    return $false
  }  

  while( ! $colShares.AtEndOfStream )
  {
    $strResponse = $colShares.ReadItem()
    if( $objNMUtilities.WinRMSetNode( $strResponse ) -ne $true )
    {
      $strRefExplanation.value = 'UNCERTAIN: WinRMSetNode failed'
      return $false
    }  
    
    $strShareName = $objNMUtilities.WinRMGetAttributeValue( 'Name' )
    $strShareStatus = $objNMUtilities.WinRMGetAttributeValue( 'Status' )
      
    if( $strShareName -eq $strCheckShareName )
    {
      $strRefShareStatus.value = $strShareStatus
      return $true
    }
  }
  
  $strRefExplanation.value = 'ERROR: Share does not exist on windows machine [' + $strHost + ']'
  return $false    
}





