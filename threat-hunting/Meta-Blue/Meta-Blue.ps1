<#
.SYNOPSIS  
    

.EXAMPLE
    .\Meta-Blue.ps1

.NOTES  
    File Name      : Meta-Blue.ps1
    Version        : v.0.1
    Author         : newhandle
    Prerequisite   : PowerShell
    Created        : 1 Oct 18
    Change Date    : April 9th 2021
    
#>

$timestamp = (get-date).Tostring("yyyy_MM_dd_hh_mm_ss")

<#
    Define the root directory for results. CHANGE THIS TO BE WHEREVER YOU WANT.
#>
$metaBlueFolder = "C:\Meta-Blue\"
$outFolder = "$metaBlueFolder\$timestamp"
#$outFolder = "$home\desktop\collection\$timestamp"
$rawFolder = "$outFolder\raw"
$anomaliesFolder = "$outFolder\Anomalies"
#$jsonFolder = "C:\MetaBlue Results"
$excludedHostsFile = "C:\Meta-Blue\ExcludedHosts.csv"

function Build-Directories{
    if(!(test-path $outFolder)){
        new-item -itemtype directory -path $outFolder -Force
    }
    if(!(test-path $rawFolder)){
        new-item -itemtype directory -path $rawFolder -Force
    }if(!(test-path $anomaliesFolder)){
        new-item -itemtype directory -path $anomaliesFolder -Force
    }
    <#if(!(test-path $jsonFolder)){
        new-item -itemtype directory -path $jsonFolder -Force
    }#>
}

$adEnumeration = $false
$winrm = $true
$localBox = $false
$waitForJobs = ""
$runningJobThreshold = 5
$jobTimeOutThreshold = 20
$isRanAsSchedTask = $false

$nodeList = [System.Collections.ArrayList]@()


<#
    This function will convert a folder of csvs to json.
#>
function Make-Json{

    do{
        $json = Read-Host "Do you want to convert to json for forwarding?(y/n)"

        if($json -ieq 'n'){
            return $false
        }elseif($json -ieq 'y'){
            foreach ($file in Get-ChildItem $rawFolder){
                $name = $file.basename
                Import-Csv $file.fullname | ConvertTo-Json | Out-File "$jsonFolder\$name.json"
            }
            return $true
        }else{
            Write-Host "Not a valid option"
        }
    }while($true)
    
}

function Get-Exports {
<#
.SYNOPSIS
Get-Exports, fetches DLL exports and optionally provides
C++ wrapper output (idential to ExportsToC++ but without
needing VS and a compiled binary). To do this it reads DLL
bytes into memory and then parses them (no LoadLibraryEx).
Because of this you can parse x32/x64 DLL's regardless of
the bitness of PowerShell.

.DESCRIPTION
Author: Ruben Boonen (@FuzzySec)
License: BSD 3-Clause
Required Dependencies: None
Optional Dependencies: None

.PARAMETER DllPath

Absolute path to DLL.

.PARAMETER CustomDll

Absolute path to output file.

.EXAMPLE
C:\PS> Get-Exports -DllPath C:\Some\Path\here.dll

.EXAMPLE
C:\PS> Get-Exports -DllPath C:\Some\Path\here.dll -ExportsToCpp C:\Some\Out\File.txt
#>
	param (
        [Parameter(Mandatory = $True)]
		[string]$DllPath,
		[Parameter(Mandatory = $False)]
		[string]$ExportsToCpp
	)

	Add-Type -TypeDefinition @"
	using System;
	using System.Diagnostics;
	using System.Runtime.InteropServices;
	using System.Security.Principal;
	
	[StructLayout(LayoutKind.Sequential)]
	public struct IMAGE_EXPORT_DIRECTORY
	{
		public UInt32 Characteristics;
		public UInt32 TimeDateStamp;
		public UInt16 MajorVersion;
		public UInt16 MinorVersion;
		public UInt32 Name;
		public UInt32 Base;
		public UInt32 NumberOfFunctions;
		public UInt32 NumberOfNames;
		public UInt32 AddressOfFunctions;
		public UInt32 AddressOfNames;
		public UInt32 AddressOfNameOrdinals;
	}
	
	[StructLayout(LayoutKind.Sequential)]
	public struct IMAGE_SECTION_HEADER
	{
		public String Name;
		public UInt32 VirtualSize;
		public UInt32 VirtualAddress;
		public UInt32 SizeOfRawData;
		public UInt32 PtrToRawData;
		public UInt32 PtrToRelocations;
		public UInt32 PtrToLineNumbers;
		public UInt16 NumOfRelocations;
		public UInt16 NumOfLines;
		public UInt32 Characteristics;
	}
	
	public static class Kernel32
	{
		[DllImport("kernel32.dll")]
		public static extern IntPtr LoadLibraryEx(
			String lpFileName,
			IntPtr hReservedNull,
			UInt32 dwFlags);
	}
"@

	# Load the DLL into memory so we can refference it like LoadLibrary
	$FileBytes = [System.IO.File]::ReadAllBytes($DllPath)
	[IntPtr]$HModule = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($FileBytes.Length)
	[System.Runtime.InteropServices.Marshal]::Copy($FileBytes, 0, $HModule, $FileBytes.Length)

	# Some Offsets..
	$PE_Header = [Runtime.InteropServices.Marshal]::ReadInt32($HModule.ToInt64() + 0x3C)
	$Section_Count = [Runtime.InteropServices.Marshal]::ReadInt16($HModule.ToInt64() + $PE_Header + 0x6)
	$Optional_Header_Size = [Runtime.InteropServices.Marshal]::ReadInt16($HModule.ToInt64() + $PE_Header + 0x14)
	$Optional_Header = $HModule.ToInt64() + $PE_Header + 0x18

	# We need some values from the Section table to calculate RVA's
	$Section_Table = $Optional_Header + $Optional_Header_Size
	$SectionArray = @()
	for ($i; $i -lt $Section_Count; $i++) {
		$HashTable = @{
			VirtualSize = [Runtime.InteropServices.Marshal]::ReadInt32($Section_Table + 0x8)
			VirtualAddress = [Runtime.InteropServices.Marshal]::ReadInt32($Section_Table + 0xC)
			PtrToRawData = [Runtime.InteropServices.Marshal]::ReadInt32($Section_Table + 0x14)
		}
		$Object = New-Object PSObject -Property $HashTable
		$SectionArray += $Object
		
		# Increment $Section_Table offset by Section size
		$Section_Table = $Section_Table + 0x28
	}

	# Helper function for dealing with on-disk PE offsets.
	# Adapted from @mattifestation:
	# https://github.com/mattifestation/PowerShellArsenal/blob/master/Parsers/Get-PE.ps1#L218
	function Convert-RVAToFileOffset($Rva, $SectionHeaders) {
		foreach ($Section in $SectionHeaders) {
			if (($Rva -ge $Section.VirtualAddress) -and
				($Rva-lt ($Section.VirtualAddress + $Section.VirtualSize))) {
				return [IntPtr] ($Rva - ($Section.VirtualAddress - $Section.PtrToRawData))
			}
		}
		# Pointer did not fall in the address ranges of the section headers
		echo "Mmm, pointer did not fall in the PE range.."
	}

	# Read Magic UShort to determin x32/x64
	if ([Runtime.InteropServices.Marshal]::ReadInt16($Optional_Header) -eq 0x010B) {
		echo "`n[?] 32-bit Image!"
		# IMAGE_DATA_DIRECTORY[0] -> Export
		$Export = $Optional_Header + 0x60
	} else {
		echo "`n[?] 64-bit Image!"
		# IMAGE_DATA_DIRECTORY[0] -> Export
		$Export = $Optional_Header + 0x70
	}

	# Convert IMAGE_EXPORT_DIRECTORY[0].VirtualAddress to file offset!
	$ExportRVA = Convert-RVAToFileOffset $([Runtime.InteropServices.Marshal]::ReadInt32($Export)) $SectionArray

	# Cast offset as IMAGE_EXPORT_DIRECTORY
	$OffsetPtr = New-Object System.Intptr -ArgumentList $($HModule.ToInt64() + $ExportRVA)
	$IMAGE_EXPORT_DIRECTORY = New-Object IMAGE_EXPORT_DIRECTORY
	$IMAGE_EXPORT_DIRECTORY = $IMAGE_EXPORT_DIRECTORY.GetType()
	$EXPORT_DIRECTORY_FLAGS = [system.runtime.interopservices.marshal]::PtrToStructure($OffsetPtr, [type]$IMAGE_EXPORT_DIRECTORY)

	# Print the in-memory offsets!
	echo "`n[>] Time Stamp: $([timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($EXPORT_DIRECTORY_FLAGS.TimeDateStamp)))"
	echo "[>] Function Count: $($EXPORT_DIRECTORY_FLAGS.NumberOfFunctions)"
	echo "[>] Named Functions: $($EXPORT_DIRECTORY_FLAGS.NumberOfNames)"
	echo "[>] Ordinal Base: $($EXPORT_DIRECTORY_FLAGS.Base)"
	echo "[>] Function Array RVA: 0x$('{0:X}' -f $EXPORT_DIRECTORY_FLAGS.AddressOfFunctions)"
	echo "[>] Name Array RVA: 0x$('{0:X}' -f $EXPORT_DIRECTORY_FLAGS.AddressOfNames)"
	echo "[>] Ordinal Array RVA: 0x$('{0:X}' -f $EXPORT_DIRECTORY_FLAGS.AddressOfNameOrdinals)"

	# Get equivalent file offsets!
	$ExportFunctionsRVA = Convert-RVAToFileOffset $EXPORT_DIRECTORY_FLAGS.AddressOfFunctions $SectionArray
	$ExportNamesRVA = Convert-RVAToFileOffset $EXPORT_DIRECTORY_FLAGS.AddressOfNames $SectionArray
	$ExportOrdinalsRVA = Convert-RVAToFileOffset $EXPORT_DIRECTORY_FLAGS.AddressOfNameOrdinals $SectionArray

	# Loop exports
	$ExportArray = @()
	for ($i=0; $i -lt $EXPORT_DIRECTORY_FLAGS.NumberOfNames; $i++){
		# Calculate function name RVA
		$FunctionNameRVA = Convert-RVAToFileOffset $([Runtime.InteropServices.Marshal]::ReadInt32($HModule.ToInt64() + $ExportNamesRVA + ($i*4))) $SectionArray
		$HashTable = @{
			FunctionName = [System.Runtime.InteropServices.Marshal]::PtrToStringAnsi($HModule.ToInt64() + $FunctionNameRVA)
			ImageRVA = echo "0x$("{0:X8}" -f $([Runtime.InteropServices.Marshal]::ReadInt32($HModule.ToInt64() + $ExportFunctionsRVA + ($i*4))))"
			Ordinal = [Runtime.InteropServices.Marshal]::ReadInt16($HModule.ToInt64() + $ExportOrdinalsRVA + ($i*2)) + $EXPORT_DIRECTORY_FLAGS.Base
		}
		$Object = New-Object PSObject -Property $HashTable
		$ExportArray += $Object
	}

	# Print export object
	$ExportArray |Sort-Object Ordinal

	# Optionally write ExportToC++ wrapper output
	if ($ExportsToCpp) {
		foreach ($Entry in $ExportArray) {
			Add-Content $ExportsToCpp "#pragma comment (linker, '/export:$($Entry.FunctionName)=[FORWARD_DLL_HERE].$($Entry.FunctionName),@$($Entry.Ordinal)')"
		}
	}

	# Free buffer
	[Runtime.InteropServices.Marshal]::FreeHGlobal($HModule)
}

function Hunt-SolarBurst{
    Write-host "Starting SolarBurst Hunt"

    $SolarBurstDLL = "C:\Program Files\Solarwinds\Orion\SolarWinds.Core.BusinessLayer.dll"
    $SolarBurstDLLAlt = "C:\Program Files (x86)\Solarwinds\Orion\SolarWinds.Core.BusinessLayer.dll"

    if($localBox){
          if(Test-Path $SolarBurstDLL){
                Get-FileHash -Algorithm SHA256 $SolarBurstDLL | export-csv -NoTypeInformation -Append "$outFolder\Local_SolarBurst.csv" | Out-Null

          }
          if(Test-Path $SolarBurstDLLAlt){
                Get-FileHash -Algorithm SHA256 $SolarBurstDLLAlt | export-csv -NoTypeInformation -Append "$outFolder\Local_SolarBurst.csv" | Out-Null
          }
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock {
                if(Test-Path $SolarBurstDLL){
                    Get-FileHash -Algorithm SHA256 $SolarBurstDLL 
                }
                if(Test-Path $SolarBurstDLLAlt){

                    Get-FileHash -Algorithm SHA256 $SolarBurstDLLAlt
                }
            }  -asjob -jobname "SolarBurst") | out-null
        }
        Create-Artifact
    }

}

function Shipto-Splunk{
    do{
        $splunk = Read-Host "Do you want to ship to splunk?(Y/N)"
        if($splunk -ieq 'y'){
            foreach($file in Get-ChildItem $rawFolder){
                Copy-Item -Path $file.FullName -Destination $jsonFolder
            }
            return $true
        }elseif($splunk -ieq 'n'){
            return $false
        }else{
            Write-Host -ForegroundColor Red "[-]Not a valid option."
        }
    }while($true)
}

function Get-FileName($initialDirectory){
<#
    This was taken from: 
    https://social.technet.microsoft.com/Forums/office/en-US/0890adff-43ea-4b4b-9759-5ac2649f5b0b/getcontent-with-open-file-dialog?forum=winserverpowershell
#>   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Repair-PSSessions{
    $sessions = Get-PSSession
    $sessions | ?{$_.state -eq "Disconnected"} | Connect-PSSession
    $sessions | ?{$_.state -eq "Broken"} | New-PSSession -SessionOption (New-PSSessionOption -NoMachineProfile -MaxConnectionRetryCount 5)
    Get-PSSession | ?{$_.state -eq "Broken"} | Remove-PSSession 
}

<#
    https://www.powershellgallery.com/packages/Binary-Search-WExample/1.0.1/Content/Binary-Search-WExample.ps1
#>
Function Binary-Search {
[CmdletBinding()]

Param (
    [Parameter(Mandatory=$True)
    ]

    $InputArray,
    $SearchVal,
    $Attribute)   #End Param

#==============Begin Main Function============================
#$InputArray = $InputArray |sort $Attribute #remove # if the input array was not sorted before calling the function
#write-host "SearchVal is: "$Searchval
#write-host "Attribute is: "$Attribute
$LowIndex = 0                              #Low side of array segment
$Counter = 0
$TempVal = ""                              #Used to determine end of search where $Found = $False
$HighIndex = $InputArray.count             #High Side of array segment
[int]$MidPoint = ($HighIndex-$LowIndex)/2  #Mid point of array segment
#write-host "Midpoint is: $midPoint Searchval is: $Searchval"
$Global:Found = $False


While($LowIndex -le $HighIndex){
    $MidVal = $InputArray[$MidPoint].$Attribute   
                                                
    If($TempVal -eq $MidVal){              #If identical, the search has completed and $Found = $False
        $Global:Found = $False
        Return
    }
    else{
        $TempVal = $MidVal                 #Update the TempVal. Search continues.
    }
    
#write-host "Midval is: $midval"
    #Write-host "Low is $lowindex, Mid is $midpoint, High is $HighIndex"
    #read-host
        If($SearchVal -lt $MidVal) {
            #write-host "SV < MV"
            $Counter++
            $HighIndex = $MidPoint 
            [int]$MidPoint = (($HighIndex-$LowIndex)/ 2 +$LowIndex)
            }
        If($SearchVal -gt $MidVal) {
            #write-host "SV > MV"
            $Counter++
            $LowIndex = $MidPoint 
            [int]$MidPoint = ($MidPoint+(($HighIndex - $MidPoint) / 2))         
            }
        If($SearchVal -eq $MidVal) {
            $Global:Found = $True 
            #write-host "User $Midval was found. It took $Counter passes"
            
            Return $midpoint
            }
}   #End While
}   #End Function

function Create-Artifact{
<#
    There are two ways to use Create-Artifact. The first way, stores all information on a computer by 
    computer basis in their own individual folders. The second way, which is the primary use case,
    creates a single CSV per artifact for data stacking.
#>
    Repair-PSSessions
    
    $poll = $true
    while($poll){
        
        Get-job | ft
        foreach($job in (get-job)){

            $time = (Get-date)
            $elapsed = ($time - $job.PSBeginTime).minutes
            #write-host "Elapsed:" $elapsed
            
            if(($job.state -eq "completed")){

                $ComputerName = $job.Location

                $Task = $job.name
                
                if($Task -ne "Prefetch"){

                    <#
                        Comment this if-else if you don't want data stacking format.
                    #>
                    $OS = ($nodeList |?{$_.hostname -eq $ComputerName.toUpper()}).operatingsystem
                    if(($OS -like "*pro*") -or ($OS -like "*Enterprise*")){
                    #if($windowsHosts.contains($computername.toUpper())){
                        Receive-Job $job.id | export-csv -force -append -NoTypeInformation -path "$rawFolder\Host_$Task.csv" | out-null
                    }
                    elseif($OS -like "*Server*"){
                    #elseif($windowsServers.Contains($ComputerName.toUpper())){
                        Receive-Job $job.id | export-csv -force -append -NoTypeInformation -path "$rawFolder\Server_$Task.csv" | out-null
                    }else{
                        Receive-Job $job.id | export-csv -force -append -NoTypeInformation -path "$rawFolder\Unknown_$Task.csv" | out-null
                    }
                #This is for prefetch.
                }
                else{
                
                    Receive-Job $job.id | out-null
                }if(!($job.hasmoredata)){
                    remove-job $job.id -force 
                }
                
            }
            <#
                TODO:
                This stores the info of a failed job. Need to implement some form of retrying failed job.
            #>
            elseif($job.state -eq "failed"){

                $job | export-csv -Append -NoTypeInformation "$outFolder\failedjobs.csv"
                Remove-Job $job.id -force

            }
            elseif(($elapsed -ge $jobTimeOutThreshold) -and ($job.state -ne "Completed")){
                $job | export-csv -Append -NoTypeInformation "$outFolder\failedjobs.csv"
                Remove-Job $job.id -force
            }
            elseif($job.state -eq "Running"){
                continue
            }
        }Start-Sleep -Seconds 8
        if((get-job | where state -eq "completed" |measure).Count -eq 0){
            if((get-job | where state -eq "failed" |measure).Count -eq 0){
                if((get-job | where state -eq "Running" |measure).Count -lt $runningJobThreshold){
                    $poll = $false
                }                
            }
        }
    }
}

function Audit-Snort{
<#
    This function asks for a csv with a cve header from an acas vulnerability scan.
    Then, it just needs to be pointed to whatever rules file from security onion.
    It will create two text files, one with the list of mitigated cves, and the other
    with unmitigated cves.
#>
    'Please select cve list:'
    $vulns = Get-FileName
    'Please select snort rule file:'
    $rules = Get-FileName
    $vulns = import-csv $vulns
    foreach($i in $vulns){
        if(Select-String $i.cve.substring(4) $rules){
            Write-Host -ForegroundColor Green $i.cve "has an associated rule."
            $i.cve >> "$outFolder\MitigatedRules $timestamp.txt"

        }else{
            Write-Host -ForegroundColor Red $i.cve "has no associated rule."
            $i.cve >> "$outFolder\UnmitigatedRules $timestamp.txt"
        }
    
    }
}

function Get-SubnetRange {
<#
    Thank you mr sparkly markley for this super awesome cool subnetrange generator.
#>
    [CmdletBinding(DefaultParameterSetName = "Set1")]
    Param(
        [Parameter(
        Mandatory          =$true,
        Position           = 0,
        ValueFromPipeLine  = $false,
        ParameterSetName   = "Set1"
        )]
        [ValidatePattern(@"
^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$
"@
        )]
            [string]$IPAddress,



        [Parameter(
        Mandatory          =$true,
        ValueFromPipeline  = $false,
        ParameterSetName   = "Set1"
        )]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern(@"
^([0-9]|[0-2][0-9]|3[0-2]|\/([0-9]|[0-2][0-9]|3[0-2]))$
"@
        )]
            [string]$CIDR

)
$IPSubnetList = [System.Collections.ArrayList]@()

#Ip Address in Binary #######################################
$Binary  = $IPAddress -split '\.' | Foreach {
    [System.Convert]::ToString($_,2).Padleft(8,'0')}
$Binary = $Binary -join ""


#Host Bits from CIDR #######################################
if ($CIDR -like '/*') {
$CIDR = ($CIDR -split '/')[1] }
$HostBits = 32 - $CIDR



$NetworkIDinBinary = $Binary.Substring(0,$CIDR)
$FirstHost = $NetworkIDinBinary.Padright(32,'0')
$LastHost = $NetworkIDinBinary.padright(32,'1')

#Getting IP of Hosts #######################################
$x = 1


while ($FirstHost -lt $LastHost) {
   
    $Octet1 = $FirstHost[0..7] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet2 = $FirstHost[8..15] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet3 = $FirstHost[16..23] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet4 = $FirstHost[24..31] -join "" | foreach {[System.Convert]::ToByte($_,2)}

    $NewIPAddress = $Octet1,$Octet2,$Octet3,$Octet4 -join "."

    if(!($NewIPAddress -like "*.0")){
        $IPSubnetList.add($NewIPAddress) | out-null
    }
    $NetworkBitsinBinary = $FirstHost.Substring(0,$FirstHost.Length-$HostBits)

    $xInBinary = [System.Convert]::ToString($x,2).padleft($HostBits,'0')

    $FirstHost = $NetworkBitsinBinary+$xInBinary

    ++$x

    }

# Adds Last IP because the while loop refuses to for whatever reason #
    $Octet1 = $LastHost[0..7] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet2 = $LastHost[8..15] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet3 = $LastHost[16..23] -join "" | foreach {[System.Convert]::ToByte($_,2)}
    $Octet4 = $LastHost[24..31] -join "" | foreach {[System.Convert]::ToByte($_,2)}

    
    $NewIPAddress = $Octet1,$Octet2,$Octet3,$Octet4 -join "."
    
    if(!($NewIPAddress -like "*.255")){
        $IPSubnetList.add( $NewIPAddress) | out-null
    }

# Calls on IP List #######################################
   return $IPSubnetList
}

function Enumerator([System.Collections.ArrayList]$iparray){
<#
    TODO: We need to add a method for determining windows hosts past ICMP. 
    we can test-netconnection -port 135 and test-netconnection -commontcpport smb
    TODO: For everything that is left after that, banner grab 22 because that should help
    identify linux devices.
    TODO: Past that, consult something else.
#>
<#
    Enumerator asynchronously pings and asynchronously performs DNS name resolution.
#>
    Build-Directories

    if($adEnumeration){
        Write-host -ForegroundColor Green "[+]Checking Windows OS Type"
   
        foreach($i in $iparray){
            if($i -ne $null){                
                    (Invoke-Command -ComputerName $i -ScriptBlock  {(gwmi win32_operatingsystem).caption} -AsJob -JobName $i) | out-null               
                    }
                }get-job | wait-job | out-null
    }
    else{
        <#
            Asynchronously Ping
        #>
        $task = foreach($ip in $iparray){
            ((New-Object System.Net.NetworkInformation.Ping).SendPingAsync($ip))
        }[threading.tasks.task]::WaitAll($task)

        
        $result = $task.Result
        
        $result = $result | ?{($_.status -eq "Success") -and ($_.address -ne "0.0.0.0") -and ($iparray.Contains($_.Address.IPAddressToString))}
        
        $duplicateIp = [System.Collections.ArrayList]@()
        foreach($i in $result){
            if(!$duplicateIp.Contains($i.Address.IPAddressToString)){

                $duplicateIp.add($i.Address.IPAddressToString) | out-null
                
                $nodeObj = [PSCustomObject]@{
                    HostName = ""
                    IPAddress = ""
                    OperatingSystem = ""
                    TTL = 0
                }
                
                $nodeObj.IPAddress = $i.Address.IPAddressToString
                
                $nodeObj.TTL = $i.Options.ttl
            
                $nodeList.Add($nodeObj) | Out-Null
            }
        }

        write-host -ForegroundColor Green "[+]There are" ($nodeList | measure).count "total live hosts."

        foreach($i in $nodeList){
            $ttl = $i.ttl
            if($ttl -le 64 -and $ttl -ge 45){
                $i.OperatingSystem = "*NIX"
            }elseif($ttl -le 128 -and $ttl -ge 115){
                $i.OperatingSystem = "Windows"
            
            }elseif($ttl -le 255 -and $ttl -ge 230){
                $i.OperatingSystem = "Cisco"
            }
        }

        Write-Host -ForegroundColor Green "[+]Connection Testing Complete beep boop beep"
        Write-Host -ForegroundColor Green "[+]Starting Reverse DNS Resolution"

        <#
            Asynchronously Resolve DNS Names
        #>
        
        $dnsTask = foreach($i in $nodeList){
                    [system.net.dns]::GetHostEntryAsync($i.ipaddress)
                    
        }[threading.tasks.task]::WaitAll($dnsTask) | out-null

        $dnsTask = $dnsTask | ?{$_.status -ne "Faulted"}

        $nodelist = $nodelist | sort ipaddress
        
        foreach($i in $dnsTask){
            $hostname = (($i.result.hostname).split('.')[0]).toUpper()
            $ip = ($i.result.addresslist.Ipaddresstostring)
            if(($ip -ne $null) -and ($hostname -ne $Null) -and ($ip -ne "") -and ($hostname -ne "")){
                $index = Binary-Search $nodeList $ip ipaddress
                if(($index -ne "") -and ($index -ne $null)){
                    $nodeList[$index].hostname = $hostname
                }
            }
            
        }
            
        Write-Host -ForegroundColor Green "[+]Reverse DNS Resolution Complete"   

        Write-host -ForegroundColor Green "[+]Checking Windows OS Type"

        foreach($i in $nodeList){
            if(($i.operatingsystem -eq "Windows")){
                $comp = $i.ipaddress
                Write-Host -ForegroundColor Green "Starting OS ID Job on:" $comp
                if(($i.hostname -ne "") -and ($i.hostname -ne $null)){
                    #Start-Job -Name $comp -ScriptBlock {gwmi win32_operatingsystem -ComputerName $using:comp -ErrorAction SilentlyContinue}|Out-Null
                    
                    Invoke-Command -ComputerName $i.hostname -ScriptBlock {gwmi win32_operatingsystem -ErrorAction SilentlyContinue} -AsJob -JobName $comp | out-null
                }
            }
        }
        
    }
    Write-Host -ForegroundColor Green "[+]All OS Jobs Started"
    
    $poll = $true
    
    $refTime = (Get-Date)
    while($poll){
        foreach($job in (get-job)){
            $time = (Get-Date)
            $elapsed = ($time - $job.PSBeginTime).minutes
            if($job.state -eq "completed"){

                 $osinfo = Receive-Job $job -ErrorAction SilentlyContinue
                 remove-job $job
                 if(($osinfo -ne $null) -and ($osinfo -ne "") -and ($osinfo.csname -ne "") -and ($osinfo.csname -ne $null)){

                    $hostname = (($osinfo.CSName).split('.')[0]).toUpper()
                    
                    foreach($i in $nodeList){
                        if($i.IPAddress -eq $job.name){
                            $i.hostname = $hostname
                            $i.operatingsystem =$osinfo.caption
                        }
                    }
                }
            }
            elseif($job.State -eq "failed"){
                Remove-Job $job.id -Force
            }
            elseif(($elapsed -ge $jobTimeOutThreshold) -and ($job.state -ne "Completed")){
                Write-Host "Stopping Job:" $job.Name
                $job | stop-job
            }
        }Start-Sleep -Seconds 8
        if((get-job | where state -eq "completed" |measure).Count -eq 0){
            if((get-job | where state -eq "failed" |measure).Count -eq 0){
                if((get-job | where state -eq "Running" |measure).Count -lt $runningJobThreshold){
                    $poll = $false
                    Write-Host "Total Elapsed:" ((get-date) - $refTime).Minutes
                }
            }
        }
    }
    <#
        Create the DnsMapper.csv
    #>
    $nodeList.getEnumerator() | Select-Object -Property @{N='HostName';E={$_.hostname}},@{N='IPAddress';E={$_.IPAddress}},@{N='OperatingSystem';E={$_.OperatingSystem}},@{N='TTL';E={$_.TTL}} | Export-Csv -path "$outfolder\NodeList.csv" -NoTypeInformation
    Write-Host -ForegroundColor Green "[+]NodeList.csv created"

    Get-Job
    write-host -ForegroundColor Green "Operating System identification jobs are done."    

    Get-Job | ?{$_.state -ne "Stopped"} | Remove-Job -Force

}

function Memory-Dumper{
    #TODO:Adapt this for other memory dump solutions like dumpit
     <#
        Create individual folders and files under $home\desktop\Meta-Blue
     #>
    foreach($i in $dnsMappingTable.Values){
        if(!(test-path $outFolder\$i)){
            new-item -itemtype directory -path $outFolder\$i -force
        }
    }
    
    Write-host -ForegroundColor Green "Begin Memory Dumping"

    <#
        Create PSSessions
    #>
    foreach($i in $windowsHosts){
        Write-host "Starting PSSession on" $i
        New-pssession -computername $i -name $i | out-null
    }
    foreach($i in $windowsServers){
        Write-host "Starting PSSession on" $i
        New-pssession -computername $i -credential $socreds -name $i | out-null
    }

    if((Get-PSSession | measure).count -eq 0){
        return
    }

    write-host -ForegroundColor Green "There are" ((Get-PSSession | measure).count) "Sessions."

    foreach($i in (Get-PSSession)){
        if(!(invoke-command -session $i -ScriptBlock { Test-Path "c:\winpmem-2.1.post4.exe" })){
            Write-host -ForegroundColor Green "Select winpmem-2.1.post4.exe location:"
            Copy-Item -ToSession $i $(Get-FileName) -Destination "c:\"
        }        
        Invoke-Command -Session $i -ScriptBlock {rm "$home\documents\memory.aff4" -ErrorAction SilentlyContinue}
        Write-Host "Starting Memory Dump on" $i.computername       
        Invoke-Command -session $i -ScriptBlock  {"$(C:\winpmem-2.1.post4.exe -o memory.aff4)" } -asjob -jobname "Memory Dumps" | Out-Null 
        
    }get-job | wait-job

    Write-host "Collecting Memory Dumps"
    foreach($i in (Get-PSSession)){
        $name = $i.computername
        Write-Host "Collecting" $name "'s dump"
        Copy-Item -FromSession $i "$home\documents\memory.aff4" -Destination "$outFolder\$name memorydump"
    }

    Get-PSSession | Remove-PSSession
    get-job | remove-job -Force

}

<#
    MITRE ATT&CK: T1015
#>
function AccessibilityFeature{
    Write-host "Starting AccessibilityFeature Jobs"
    if($localBox){
        Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\*' | export-csv -NoTypeInformation -Append "$outFolder\Local_AccessibilityFeature.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\*' | select DisableExeptionChainValidation,MitigationOptions,PSPath,PSChildName,PSComputerName}  -asjob -jobname "AccessibilityFeature") | out-null
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1100
#>
function WebShell{
    Write-host "Starting WebShell Jobs"
    if($localBox){
        gci -path "C:\inetpub\wwwroot" -recurse -File -ea SilentlyContinue | Select-String -Pattern "runat" | export-csv -NoTypeInformation -Append "$outFolder\Local_WebShell.csv" | Out-Null
        gci -path "C:\inetpub\wwwroot" -recurse -File -ea SilentlyContinue | Select-String -Pattern "eval" | export-csv -NoTypeInformation -Append "$outFolder\Local_WebShell.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                gci -path "C:\inetpub\wwwroot" -recurse -File -ea SilentlyContinue | Select-String -Pattern "runat";
                gci -path "C:\inetpub\wwwroot" -recurse -File -ea SilentlyContinue | Select-String -Pattern "eval"
            }  -asjob -jobname "WebShell") | out-null
        }
        Create-Artifact
    }
}

function Processes{
    Write-host "Starting Process Jobs"
    if($localBox){
        gwmi win32_process | export-csv -NoTypeInformation -Append "$outFolder\Local_Processes.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_process | Select-Object -Property processname,handles,path.pscomputernamename,commandline,creationdate,executablepath,parentprocessid,processid}  -asjob -jobname "Processes") | out-null
        }
        Create-Artifact
    }
}

function DNSCache{
    Write-host "Starting DNSCache Jobs"
    if($localBox){
        Get-DnsClientCache | export-csv -NoTypeInformation -Append "$outFolder\Local_DNSCache.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-DnsClientCache -ErrorAction SilentlyContinue | Select-Object -Property TTL,pscomputername,data,entry,name}  -asjob -jobname "DNSCache") | out-null
        }
        Create-Artifact
    }
}

function ProgramData{
    Write-host "Starting ProgramData Enum"
    if($localBox){
        Get-ChildItem -Recurse C:\ProgramData | export-csv -NoTypeInformation -Append "$outFolder\Local_ProgramData.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-ChildItem -Recurse c:\ProgramData\ | Select-Object -Property Fullname,Pscomputername,creationtimeutc,lastaccesstimeutc,attributes}  -asjob -jobname "ProgramData")| out-null         
        }
        Create-Artifact
    }
}

function AlternateDataStreams{
    Write-host "Starting AlternateDataStreams Enum"
    if($localBox){
        Set-Location C:\Users
        (Get-ChildItem -Recurse).fullname | Get-Item -Stream * | ?{$_.stream -ne ':$DATA'} | export-csv -NoTypeInformation -Append "$outFolder\Local_AlternateDataStreams.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Set-Location C:\Users; (Get-ChildItem -Recurse).fullname | Get-Item -Stream * | ?{$_.stream -ne ':$DATA'} | Select-Object -Property Filename,Pscomputername,stream}  -asjob -jobname "AlternateDataStreams")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1128
#>
function NetshHelperDLL{
    Write-host "Starting NetshHelperDLL Enum"
    if($localBox){
        (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Netsh') | export-csv -NoTypeInformation -Append "$outFolder\Local_NetshHelperDLL.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Netsh')} -asjob -jobname "NetshHelperDLL")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1013
#>
function PortMonitors{
    Write-host "Starting PortMonitors Enum"
    if($localBox){
        (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\*") | export-csv -NoTypeInformation -Append "$outFolder\Local_PortMonitors.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\*") | Select-Object -Property Driver,Pschildname,pscomputername}  -asjob -jobname "PortMonitors")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1038
#>
function KnownDLLs{
    Write-host "Starting KnownDLLs Enum"
    if($localBox){
        (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs\') | export-csv -NoTypeInformation -Append "$outFolder\Local_KnownDLLs.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {(Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\KnownDLLs\')}  -asjob -jobname "KnownDLLs")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1038
#>
function DLLSearchOrderHijacking{
    Write-host "Starting DLLSearchOrderHijacking Enum"
    if($localBox){
        (gci -path C:\Windows\* -include *.dll | Get-AuthenticodeSignature | Where-Object Status -NE "Valid") | export-csv -NoTypeInformation -Append "$outFolder\Local_DLLSearchOrderHijacking.csv" | Out-Null
        (gci -path C:\Windows\System32\* -include *.dll | Get-AuthenticodeSignature | Where-Object Status -NE "Valid") | export-csv -NoTypeInformation -Append "$outFolder\Local_DLLSearchOrderHijacking.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {
                (gci -path C:\Windows\System32\* -include *.dll | Get-AuthenticodeSignature | Where-Object Status -NE "Valid");
                (gci -path C:\Windows\* -include *.dll | Get-AuthenticodeSignature | Where-Object Status -NE "Valid")
             }  -asjob -jobname "DLLSearchOrderHijacking")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1197
#>
function BITSJobs{
    Write-host "Starting BITSJobs Enum"
    if($localBox){
        Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Bits-Client/Operational'; Id='59'} | export-csv -NoTypeInformation -Append "$outFolder\Local_BITSJobs.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Bits-Client/Operational'; Id='59'} | Select-Object -Property message,pscomputername,id,logname,processid,userid,timecreated}  -asjob -jobname "BITSJobs")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1019
#>
function SystemFirmware{
    Write-host "Starting SystemFirmware Enum"
    if($localBox){
        Get-WmiObject win32_bios | export-csv -NoTypeInformation -Append "$outFolder\Local_SystemFirmware.csv" | Out-Null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-WmiObject win32_bios | Select-Object -Property pscomputername,biosversion,caption,currentlanguage,manufacturer,name,serialnumber}  -asjob -jobname "SystemFirmware")| out-null         
        }
        Create-Artifact
    }
}

<#
    T1037.001
#>
function UserInitMprLogonScript{
    Write-host "Starting UserInitMprLogonScript Enum"

    if($localBox){
        $logonScriptsArrayList = [System.Collections.ArrayList]@();
                 
                 New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue | Out-Null;
                 Set-Location HKU: | Out-Null;

                 $SIDS  += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };

                 foreach($SID in $SIDS){
                    $logonscriptObject = [PSCustomObject]@{
                        SID =""
                        HasLogonScripts = ""
                 
                    };
                    $logonscriptObject.sid = $SID; 
                    $logonscriptObject.haslogonscripts = !((Get-ItemProperty HKU:\$SID\Environment\).userinitmprlogonscript -eq $null); 
                    $logonScriptsArrayList.add($logonscriptObject) | out-null
                    }
                    $logonScriptsArrayList | export-csv -NoTypeInformation -Append "$outFolder\Local_UserInitMprLogonScript.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {
                 $logonScriptsArrayList = [System.Collections.ArrayList]@();
                 
                 New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue | Out-Null;
                 Set-Location HKU: | Out-Null;

                 $SIDS  += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };

                 foreach($SID in $SIDS){
                    $logonscriptObject = [PSCustomObject]@{
                        SID =""
                        HasLogonScripts = ""
                 
                    };
                    $logonscriptObject.sid = $SID; 
                    $logonscriptObject.haslogonscripts = !((Get-ItemProperty HKU:\$SID\Environment\).userinitmprlogonscript -eq $null); 
                    $logonScriptsArrayList.add($logonscriptObject) | out-null
                    }
                    $logonScriptsArrayList
                 
             
             }  -asjob -jobname "UserInitMprLogonScript")| out-null         
        }
        Create-Artifact
    }
}

function InstalledSoftware{
    Write-host "Starting InstalledSoftware Jobs"
    if($localBox){
        $(Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*; 
        Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*;
        if(!(test-path HKU:)){
            New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS| Out-Null;
        }
        $UserInstalls += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };
        $(foreach ($User in $UserInstalls){Get-ItemProperty HKU:\$User\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*});
        $UserInstalls = $null;try{Remove-PSDrive -Name HKU}catch{};)|where {($_.DisplayName -ne $null) -and ($_.Publisher -ne $null)} | export-csv -NoTypeInformation -Append "$outFolder\Local_InstalledSoftware.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                $(Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*; 
                Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*;
                if(!(test-path HKU:)){
                    New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS| Out-Null;
                }
                $UserInstalls += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };
                $(foreach ($User in $UserInstalls){Get-ItemProperty HKU:\$User\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*});
                $UserInstalls = $null;)|where {($_.DisplayName -ne $null) -and ($_.Publisher -ne $null)}
            }  -asjob -jobname "InstalledSoftware") | out-null
        }
        Create-Artifact
    }
}

function Registry{
<#
        You can add anything you want here but should reserve it for registry queries. The queries get added in as
        noteproperties to the PSCustomObject so that they can be exported to CSV in a stackable format. If the registry
        key is an array, typecast to a string. Ensure you pick a property name that reflects the forensic relavence of
        the registry location.
#>
    Write-host "Starting Registry Jobs"

    if(!(test-path HKU:)){
        New-PSDrive HKU Registry HKEY_USERS
    }
    Set-Location HKU:

    if($localBox){
        

        $registry = [PSCustomObject]@{
                
                <#
                    MITRE ATT&CK: T1182
                #>
                AppCertDLLs = (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\appcertdlls\')

                

                BootShell = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\' -ErrorAction SilentlyContinue).bootshell

                BootExecute = [String](Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\' -ErrorAction SilentlyContinue).bootexecute

                NetworkList = [String]((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\UnManaged\*' -ErrorAction SilentlyContinue).dnssuffix)
                
                <#
                        MITRE ATT&CK: T1131
                #>
                AuthenticationPackage = [String]((get-itemproperty HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\ -ErrorAction SilentlyContinue).('authentication packages'))

                HKLMRun = [String](get-item 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run\' -ErrorAction SilentlyContinue).property
                HKCURun = [String](get-item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\' -ErrorAction SilentlyContinue).property
                HKLMRunOnce = [String](get-item 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce\' -ErrorAction SilentlyContinue).property
                HKCURunOnce = [String](Get-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce\' -ErrorAction SilentlyContinue).property

                Shell = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\' -ErrorAction SilentlyContinue).shell

                Manufacturer = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\' -ErrorAction SilentlyContinue).manufacturer

                <#
                        MITRE ATT&CK: T1103
                #>
                AppInitDlls = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' -ErrorAction SilentlyContinue).appinit_dlls

                ShimCustom = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Custom' -ErrorAction SilentlyContinue)

                UserInit = [String](Get-ItemProperty ('HKLM:\software\Microsoft\Windows NT\CurrentVersion\Winlogon\') -ErrorAction SilentlyContinue).userinit

                Powershellv2 = if((test-path HKLM:\SOFTWARE\Microsoft\PowerShell\1\powershellengine\)){$true}else{$false}
            }
            $registry | Export-Csv -NoTypeInformation -Append "$outFolder\Local_Registry.csv"
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {  
                

                $registry = [PSCustomObject]@{
                    
                    <#
                        MITRE ATT&CK: T1182
                    #>
                    AppCertDLLs = (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\appcertdlls\')

                    BootShell = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\' -ErrorAction SilentlyContinue).bootshell

                    BootExecute = [String](Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\' -ErrorAction SilentlyContinue).bootexecute

                    NetworkList = [String]((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\UnManaged\*' -ErrorAction SilentlyContinue).dnssuffix)
                    
                    <#
                        MITRE ATT&CK: T1131
                    #>
                    AuthenticationPackage = [String]((get-itemproperty HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\ -ErrorAction SilentlyContinue).('authentication packages'))

                    HKLMRun = [String](get-item 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run\' -ErrorAction SilentlyContinue).property
                    HKCURun = [String](get-item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\' -ErrorAction SilentlyContinue).property
                    HKLMRunOnce = [String](get-item 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce\' -ErrorAction SilentlyContinue).property
                    HKCURunOnce = [String](Get-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce\' -ErrorAction SilentlyContinue).property

                    Shell = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\' -ErrorAction SilentlyContinue).shell

                    Manufacturer = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\' -ErrorAction SilentlyContinue).manufacturer

                    <#
                        MITRE ATT&CK: T1103
                    #>
                    AppInitDlls = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' -ErrorAction SilentlyContinue).appinit_dlls

                    ShimCustom = [String](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Custom' -ErrorAction SilentlyContinue)

                    UserInit = [String](Get-ItemProperty ('HKLM:\software\Microsoft\Windows NT\CurrentVersion\Winlogon\') -ErrorAction SilentlyContinue).userinit

                    Powershellv2 = if((test-path HKLM:\SOFTWARE\Microsoft\PowerShell\1\powershellengine\)){$true}else{$false}
                }
                $registry
        
            }  -asjob -jobname "Registry") | out-null
        }
        Create-Artifact
    }
}

function AVProduct{
    Write-host "Starting AVProduct Jobs"
    if($localBox){
       Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct -ErrorAction SilentlyContinue | select PSComputerName,displayName,pathToSignedProductExe,pathToSignedReportingExe | Export-Csv -NoTypeInformation -Append "$outFolder\Local_AVProduct.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct -ErrorAction SilentlyContinue | select PSComputerName,displayName,pathToSignedProductExe,pathToSignedReportingExe}  -asjob -jobname "AVProduct") | out-null
        }
        Create-Artifact
    }
    
}

function Services{
    Write-host "Starting Services Jobs"
    if($localBox){
        gwmi win32_service | export-csv -NoTypeInformation -Append "$outFolder\Local_Services.csv" | Out-Null
    }else{ 
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_service | Select-Object -Property PSComputerName,caption,description,pathname,processid,startname,state}  -asjob -jobname "Services")| out-null
        }
        Create-Artifact
    }
}

function PoshVersion{
    Write-host "Starting PoshVersion Jobs"
    if($localBox){
        Get-WindowsOptionalFeature -Online -FeatureName microsoftwindowspowershellv2 | export-csv -NoTypeInformation -Append "$outFolder\Local_PoshVersion.csv" | Out-Null
    }else{ 
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {Get-WindowsOptionalFeature -Online -FeatureName microsoftwindowspowershellv2 | Select-Object -Property PSComputerName,FeatureName,State,LogPath}  -asjob -jobname "PoshVersion")| out-null
        }
        Create-Artifact
    }
}

function Startup{
    Write-host "Starting Startup Jobs"
    if($localBox){
        gwmi win32_startupcommand | export-csv -NoTypeInformation -Append "$outFolder\Local_Startup.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){   
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_startupcommand | Select-Object -Property PSComputerName,Caption,Command,Description,Location,User}  -asjob -jobname "Startup")| out-null         
        }
        Create-Artifact
    }
}

<#
    MITRE ATT&CK: T1060
#>
function StartupFolder{
    Write-host "Starting StartupFolder Jobs"
    if($localBox){
        gci -path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*" <#-include *.lnk,*.url#> -ErrorAction SilentlyContinue | export-csv -NoTypeInformation -Append "$outFolder\Local_StartupFolder.csv" | out-null
        gci -path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" <#-include *.lnk,*.url#> -ErrorAction SilentlyContinue | export-csv -NoTypeInformation -Append "$outFolder\Local_StartupFolder.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){   
            (Invoke-Command -session $i -ScriptBlock  {
                gci -path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*"<# -include *.lnk,*.url#> -ErrorAction SilentlyContinue| Select-Object -Property PSComputerName,Length,FullName,Extension,CreationTime,LastAccessTime;
                gci -path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" <#-include *.lnk,*.url#> -ErrorAction SilentlyContinue | Select-Object -Property PSComputerName,Length,FullName,Extension,CreationTime,LastAccessTime
            }  -asjob -jobname "StartupFolder")| out-null         
        }
        Create-Artifact
    }
}

function Drivers{
    Write-host "Starting Driver Jobs"
    if($localBox){
        gwmi win32_systemdriver | export-csv -NoTypeInformation -Append "$outFolder\Local_Drivers.csv" | Out-Null
    }else{ 
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_systemdriver | Select-Object -Property PSComputerName,caption,description,name,pathname,started,startmode,state}  -asjob -jobname "Drivers")| out-null         
        }
        Create-Artifact
    }
}

function DriverHash{
    Write-host "Starting DriverHash Jobs"
    if($localBox){
        $driverPath = (gwmi win32_systemdriver).pathname          
            foreach($driver in $driverPath){                
                    (Get-filehash -algorithm SHA256 -path $driver -ErrorAction SilentlyContinue) | export-csv -NoTypeInformation -Append "$outFolder\Local_DriverHashes.csv" | out-null                
            }
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                $driverPath = (gwmi win32_systemdriver).pathname          
                foreach($driver in $driverPath){                
                        (Get-filehash -algorithm SHA256 -path $driver -ErrorAction SilentlyContinue)                
                }
            }  -asjob -jobname "DriverHash") | out-null
        }
        Create-Artifact
    }
}

function EnvironVars{
    Write-host "Starting EnvironVars Jobs"
    if($localBox){
        gwmi win32_environment|?{$_.name -ne "OneDrive"} | export-csv -NoTypeInformation -Append "$outFolder\Local_EnvironVars.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_environment |?{$_.name -ne "OneDrive"} | Select-Object -Property PSComputerName,name,description,username,variablevalue}  -asjob -jobname "EnvironVars")| out-null         
        }
        Create-Artifact
    }  
}

function NetAdapters{
    Write-host "Starting NetAdapter Jobs"
    if($localBox){
        gwmi win32_networkadapterconfiguration | Export-Csv -NoTypeInformation -Append "$outFolder\Local_NetAdapters.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_networkadapterconfiguration | Select-Object -Property PSComputerName,Description,IPAddress,IPSubnet,MACAddress,servicename}  -asjob -jobname "NetAdapters")| out-null        
        }
        Create-Artifact
    }
}

function SystemInfo{
    Write-host "Starting SystemInfo Jobs"
    if($localBox){
        gwmi win32_computersystem | export-csv -NoTypeInformation -Append "$outFolder\Local_Systeminfo.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {gwmi win32_computersystem | Select-Object -Property PSComputerName,domain,manufacturer,model,primaryownername,totalphysicalmemory,username}  -asjob -jobname "SystemInfo")| out-null         
        }
        Create-Artifact
    }
}

function Logons{
    Write-host "Starting Logon Jobs"
    if($localBox){
        gwmi win32_networkloginprofile | export-csv -NoTypeInformation -Append "$outFolder\Local_Logons.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
            (Invoke-Command -session $i -ScriptBlock  {gwmi win32_networkloginprofile}  -asjob -jobname "Logons")| out-null         
        }
        Create-Artifact
    }
}

function NetConns{
    Write-host "Starting NetConn Jobs"
    if($localBox){
        Get-NetTCPConnection | export-csv -NoTypeInformation -Append "$outFolder\Local_NetConn.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
            (Invoke-Command -session $i -ScriptBlock  {get-NetTcpConnection}  -asjob -jobname "NetConn")| out-null        
        }
        Create-Artifact
    }
}

function SMBShares{
    Write-host "Starting SMBShare Jobs"
    if($localBox){
        Get-SmbShare | export-csv -NoTypeInformation -Append "$outFolder\Local_SMBShares.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {get-Smbshare}  -asjob -jobname "SMBShares")| out-null        
        }
        Create-Artifact
    }
}

function SMBConns{
    Write-host "Starting SMBConn Jobs"
    if($localBox){
        Get-SmbConnection | export-csv -NoTypeInformation -Append "$outFolder\Local_SMBConns.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){         
            (Invoke-Command -session $i -ScriptBlock  {get-SmbConnection}  -asjob -jobname "SMBConns")| out-null      
        }
        Create-Artifact
    }
}

function SchedTasks{
    Write-host "Starting SchedTask Jobs"
    if($localBox){
        Get-ScheduledTask | Export-Csv -NoTypeInformation -Append "$outFolder\Local_SchedTasks.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {get-scheduledtask -ErrorAction SilentlyContinue}  -asjob -jobname "SchedTasks")| out-null         
        }
        Create-Artifact
    }
}

function ProcessHash{
    Write-host "Starting Process Hash Jobs"
    if($localBox){
        $hashes = @()
            $pathsofexe = (gwmi win32_process -ErrorAction SilentlyContinue | select executablepath | sort executablepath -Unique | ?{$_.executablepath -ne ""})
            $execpaths = [System.Collections.ArrayList]@();foreach($i in $pathsofexe){$execpaths.Add($i.executablepath)| Out-Null}
            foreach($i in $execpaths){
                if($i -ne $null){
                    (Get-filehash -algorithm SHA256 -path $i -ErrorAction SilentlyContinue) | export-csv -NoTypeInformation -Append "$outFolder\Local_ProcessHash.csv" | Out-Null
                }
            }
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                $hashes = @()
                $pathsofexe = (gwmi win32_process -ErrorAction SilentlyContinue | select executablepath | sort executablepath -Unique | ?{$_.executablepath -ne ""})
                $execpaths = [System.Collections.ArrayList]@();foreach($i in $pathsofexe){$execpaths.Add($i.executablepath)| Out-Null}
                foreach($i in $execpaths){
                    if($i -ne $null){
                        (Get-filehash -algorithm SHA256 -path $i -ErrorAction SilentlyContinue)
                    }
                }
            }  -asjob -jobname "ProcessHash") | out-null
        }
        Create-Artifact
    }
}

function PrefetchListing{
    Write-host "Starting PrefetchListing Jobs"
    if($localBox){
        Get-ChildItem "C:\Windows\Prefetch" | export-csv -NoTypeInformation -Append "$outFolder\Local_PrefetchListing.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-ChildItem "C:\Windows\Prefetch"}  -asjob -jobname "PrefetchListing")| out-null         
        }
        Create-Artifact
    }
}

function PNPDevices{
    Write-host "Starting PNP Device Jobs"
    if($localBox){
        gwmi win32_pnpentity | export-csv -NoTypeInformation -Append "$outFolder\Local_PNPDevices.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {gwmi win32_pnpentity}  -asjob -jobname "PNPDevices")| out-null         
        }
        Create-Artifact
    }
}

function LogicalDisks{
    Write-host "Starting Logical Disk Jobs"
    if($localBox){
        gwmi win32_logicaldisk | export-csv -NoTypeInformation -Append "$outFolder\Local_LogicalDisks.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {gwmi win32_logicaldisk}  -asjob -jobname "LogicalDisks")| out-null         
        }
        Create-Artifact
    }
}

function DiskDrives{
    Write-host "Starting Disk Drive Jobs"
    if($localBox){
        gwmi win32_diskdrive | export-csv -NoTypeInformation -Append "$outFolder\Local_DiskDrives.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {gwmi win32_diskdrive}  -asjob -jobname "DiskDrives")| out-null         
        }
        Create-Artifact
    }
}

function WMIEventFilters{
    Write-host "Starting WMIEventFilter Jobs"
    if($localBox){
        Get-WMIObject -Namespace root\Subscription -Class __EventFilter | Export-Csv -NoTypeInformation -Append "$outFolder\Local_WMIEventFilters.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-WMIObject -Namespace root\Subscription -Class __EventFilter}  -asjob -jobname "WMIEventFilters")| out-null         
        }
        Create-Artifact
    }
}

function WMIEventConsumers{
    Write-host "Starting WMIEventConsumer Jobs"
    if($localBox){
        Get-WMIObject -Namespace root\Subscription -Class __EventConsumer | export-csv -NoTypeInformation -Append "$outFolder\Local_WMIEventConsumers.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-WMIObject -Namespace root\Subscription -Class __EventConsumer}  -asjob -jobname "WMIEventConsumers")| out-null         
        }
        Create-Artifact
    }
}

function WMIEventConsumerBinds{
    Write-host "Starting WMIEventConsumerBind Jobs"
    if($localBox){
        Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding | Export-Csv -NoTypeInformation -Append "$outFolder\Local_WMIEventConsumerBinds.csv" | Out-Null
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding}  -asjob -jobname "WMIEventConsumerBinds")| out-null         
        }
        Create-Artifact
    }
}

function DLLs{
    Write-host "Starting Loaded DLL Jobs"
    if($localBox){
        Get-Process -Module -ErrorAction SilentlyContinue | Export-Csv -NoTypeInformation -Append "$outFolder\Local_DLLs.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-Process -Module -ErrorAction SilentlyContinue}  -asjob -jobname "DLLs") | out-null
        }
        Create-Artifact
    }
}

<#
    MITRE ATTACK: T1177
#>
function LSASSDriver{
    Write-host "Starting LSASSDriver Jobs"
    if($localBox){
        Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='4614';} -ErrorAction SilentlyContinue | Export-Csv -NoTypeInformation -Append "$outFolder\Local_LSASSDriver.csv" | out-null
        Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='3033';} -ErrorAction SilentlyContinue | Export-Csv -NoTypeInformation -Append "$outFolder\Local_LSASSDriver.csv" | out-null
        Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='3063';} -ErrorAction SilentlyContinue | Export-Csv -NoTypeInformation -Append "$outFolder\Local_LSASSDriver.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
            Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='4614';} -ErrorAction SilentlyContinue;
            Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='3033';} -ErrorAction SilentlyContinue;
            Get-WinEvent -FilterHashtable @{ LogName='Security'; Id='3063';} -ErrorAction SilentlyContinue
            } -asjob -jobname "LSASSDriver") | out-null
        }
        Create-Artifact
    }
}

function DLLHash{
    Write-host "Starting Loaded DLL Hashing Jobs"
    if($localBox){
        $a = (Get-Process -Module -ErrorAction SilentlyContinue | ?{!($_.FileName -like "*.exe")})
            $a = $a.FileName.ToUpper() | sort
            $a = $a | Get-Unique
            foreach($file in $a){
                Get-FileHash -Algorithm SHA256 $file | Export-Csv -NoTypeInformation -Append "$outFolder\Local_DLLHash.csv" | Out-Null
            }
    }else{
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                $a = (Get-Process -Module -ErrorAction SilentlyContinue | ?{!($_.FileName -like "*.exe")})
                $a = $a.FileName.ToUpper() | sort
                $a = $a | Get-Unique
                foreach($file in $a){
                    Get-FileHash -Algorithm SHA256 $file
                }
        
            }  -asjob -jobname "DLLHash") | out-null
        }
        Create-Artifact
    }
}

function UnsignedDrivers{
    Write-host "Starting UnsignedDrivers Jobs"
    if($localBox){
        gci -path C:\Windows\System32\drivers -include *.sys -recurse -ea SilentlyContinue | Get-AuthenticodeSignature | where {$_.status -ne 'Valid'} | Export-Csv -NoTypeInformation -Append "$outFolder\Local_UnsignedDrivers.csv" | out-null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  { gci -path C:\Windows\System32\drivers -include *.sys -recurse -ea SilentlyContinue | Get-AuthenticodeSignature | where {$_.status -ne 'Valid'}}  -asjob -jobname "UnsignedDrivers")| out-null         
        }
        Create-Artifact
    }
}

function Hotfix{
    Write-host "Starting Hotfix Jobs"
    if($localBox){
        Get-HotFix -ErrorAction SilentlyContinue| Export-Csv -NoTypeInformation -Append "$outFolder\Local_Hotfix.csv" | out-null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  { Get-HotFix -ErrorAction SilentlyContinue}  -asjob -jobname "Hotfix")| out-null         
        }
        Create-Artifact
    }
}

function ArpCache{
    Write-host "Starting ArpCache Jobs"
    if($localBox){
        Get-NetNeighbor| Export-Csv -NoTypeInformation -Append "$outFolder\Local_ArpCache.csv" | out-null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  { Get-NetNeighbor -ErrorAction SilentlyContinue}  -asjob -jobname "ArpCache")| out-null         
        }
        Create-Artifact
    }
}

<#
    T1050
#>
function NewlyRegisteredServices{
    Write-host "Starting NewlyRegisteredServices Jobs"
    if($localBox){
       Get-WinEvent -FilterHashtable @{ LogName='System'; Id='7045';} | select timecreated,message | Export-Csv -NoTypeInformation -Append "$outFolder\Local_NewlyRegisteredServices.csv" | out-null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  { Get-WinEvent -FilterHashtable @{ LogName='System'; Id='7045';} | select timecreated,message}  -asjob -jobname "NewlyRegisteredServices")| out-null         
        }
        Create-Artifact
    }
}

<#
    T1546.013
#>
function PowershellProfile{
    Write-host "Starting PowershellProfile Jobs"
    if($localBox){
        test-path "$pshome\profile.ps1" | Export-Csv -NoTypeInformation -Append "$outFolder\Local_PowershellProfile.csv" | out-null
        test-path "$pshome\microsoft.*.ps1" | Export-Csv -NoTypeInformation -Append "$outFolder\Local_PowershellProfile.csv" | out-null
        test-path "c:\users\*\My Documents\powershell\Profile.ps1" | Export-Csv -NoTypeInformation -Append "$outFolder\Local_PowershellProfile.csv" | out-null
        test-path "C:\Users\*\My Documents\microsoft.*.ps1"| Export-Csv -NoTypeInformation -Append "$outFolder\Local_PowershellProfile.csv" | out-null
    
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {
                test-path $pshome\profile.ps1
                test-path $pshome\microsoft.*.ps1
                test-path "c:\users\*\My Documents\powershell\Profile.ps1"
                test-path "C:\Users\*\My Documents\microsoft.*.ps1"
             
             }  -asjob -jobname "PowershellProfile")| out-null         
        }
        Create-Artifact
    }
}

function UACBypassFodHelper{
    Write-host "Starting UACBypassFodHelper Jobs"
    if($localBox){
        if(!(test-path HKU:)){
            New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS| Out-Null;
        }
        $UserInstalls += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };
        foreach($user in $UserInstalls){
            Get-ItemProperty HKU:\$User\SOFTWARE\classes\ms-settings-shell\open\command -ErrorAction SilentlyContinue| Export-Csv -NoTypeInformation -Append "$outFolder\Local_UACBypassFodHelper.csv" | out-null
        }
    }else{
        foreach($i in (Get-PSSession)){
             (Invoke-Command -session $i -ScriptBlock  {
                if(!(test-path HKU:)){
                    New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS| Out-Null;
                }
                $UserInstalls += gci -Path HKU: | where {$_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$'} | foreach {$_.PSChildName };
                foreach($user in $UserInstalls){
                    if(test-path HKU\$user\Software\Classes\ms-settings\shell\open\command){
                        Get-ItemProperty HKU:\$User\SOFTWARE\classes\ms-settings-shell\open\command -ErrorAction SilentlyContinue
                    }
                }
             
             }  -asjob -jobname "UACBypassFodHelper")| out-null         
        }
        Create-Artifact
    }
}

function Update-Sysmon{
<#
    This can update sysmon for everyone. Change the location of the config file and its name 
    per your environment.
#>
    if(!$localBox){
        do{
            $sysmon = Read-Host "Do you want to update sysmon?(Y/N)"
    
            if($sysmon -ieq 'y'){
                foreach($i in (Get-PSSession)){
                    Write-Host "Updating Sysmon Configs From Sysvol on" $i.computername
                    Copy-Item -ToSession $i -Path $home\Downloads\SysinternalsSuite\Sysmon64.exe -Destination $home\Downloads -force 
                    Copy-Item -ToSession $i -Path $home\Downloads\SysinternalsSuite\sysmonconfig-export.xml -Destination $home\Downloads -force  
                    Invoke-Command -session $i -ScriptBlock  {cd $home\Downloads; $(.\sysmon64.exe -accepteula -i .\sysmonconfig-export.xml)} | out-null
                }
                return $true
            }elseif($sysmon -ieq 'n'){
                return $false   
            }else{
                Write-Host "Not a valid option"
            }
        }while($true)
    }
}

function Find-File{
    if(!$localBox){
        do{
            $findfile = Read-Host "Do you want to find some files?(Y/N)"
            if($findfile -ieq 'y'){
                $fileNames = [system.collections.arraylist]@()
                $fileNameFile = Get-FileName
                $fileNameFileImport = import-csv $filenamefile
                foreach($file in $filenamefileimport){$filenames.Add($file.filename) | Out-Null}
                Write-host "Starting File Search Jobs"
                <#
                    OK how the fuck am i gonna do this 
                #> 
                foreach($i in (Get-PSSession)){
                     (Invoke-Command -session $i -ScriptBlock{
                        $files = $using:filenames;
                        cd C:\Users;
                        Get-ChildItem -Recurse | ?{$files.Contains($_.name)}         
                     }  -asjob -jobname "FindFile")| out-null         
                }Create-Artifact
                break
            }elseif($findfile -ieq 'n'){
                return $false
            }else{
                Write-host "Not a valid option"
            }
           }while($true)
       }
}

function Retry-FailedJobs{}

function TearDown-Sessions{
    if(!$localBox){
        if($isRanAsSchedTask -eq $true){
            Remove-PSSession * | out-null
            return $true
        }else{
            do{
            $sessions = Read-Host "Do you want to tear down the PSSessions?(y/n)"

            if($sessions -ieq 'y'){
                    Remove-PSSession * | out-null
                    return $true
            }elseif($sessions -ieq 'n'){
                return $false
            }
            else{
                Write-Host "Not a valid option"
            }
        }while($true)
        }
        
    }
}

function Build-Sessions{
    if(!$localBox){
        $excludedHosts = @()

        Set-Item WSMan:\localhost\Shell\MaxShellsPerUser -Value 10000
        Set-Item WSMan:\localhost\Plugin\microsoft.powershell\Quotas\MaxShellsPerUser -Value 10000
        Set-Item WSMan:\localhost\Plugin\microsoft.powershell\Quotas\MaxShells -Value 10000
        <#
            Clean up and broken PSSessions.
        #>
        $brokenSessions = (Get-PSSession | ?{$_.State -eq "Broken"}).Id
        if($brokenSessions -ne $null){
            Remove-PSSession -id $brokenSessions
        }
        $activeSessions = (Get-PSSession | ?{$_.State -eq "Opened"}).ComputerName

        if(test-path $excludedHostsFile){
            $excludedHosts = import-csv $excludedHostsFile
        }

        <#
            Create PSSessions
        #>
        foreach($i in $nodeList){
            if($activeSessions -ne $null){
                if(!$activeSessions.Contains($i.hostname)){
                    if(($i.hostname -ne "") -and ($i.operatingsystem -like "*Windows*") -and (!$excludedHosts.hostname.Contains($i.hostname))){
                        Write-host "Starting PSSession on" $i.hostname
                        New-pssession -computername $i.hostname -name $i.hostname -SessionOption (New-PSSessionOption -NoMachineProfile -MaxConnectionRetryCount 5) -ThrottleLimit 100| out-null
                    }
                }else{
                    Write-host "PSSession already exists:" $i.hostname -ForegroundColor Red
                }
            }else{
                if(($i.hostname -ne "") -and ($i.operatingsystem -like "*windows*") -and (!$excludedHosts.hostname.Contains($i.hostname))){
                    Write-host "Starting PSSession on" $i.hostname
                    New-pssession -computername $i.hostname -name $i.hostname -SessionOption (New-PSSessionOption -NoMachineProfile -MaxConnectionRetryCount 5) -ThrottleLimit 100| out-null
                }
            }
        }
        
    
        if((Get-PSSession | measure).count -eq 0){
            return
        }    

        write-host -ForegroundColor Green "There are" ((Get-PSSession | measure).count) "Sessions."
    } 

}

function WaitFor-Jobs{
    if((get-job | where state -eq "Running" |measure).Count -ne 0){
            write-host -ForegroundColor Green "There are" ((Get-job |?{$_.state -like "*Run*"} | measure).count) "jobs still running."
        do{
            $waitForJobs = Read-Host "Do you want to wait for more of these jobs to finish?(y/n)"

            if($waitForJobs -ieq 'n'){
                Get-job | Remove-Job -Force
                break
            }elseif($waitForJobs -ieq 'y'){
                $runningJobThreshold--
                Create-Artifact
                #break
            }else{
                Write-Host "Not a valid option"
            }
            
        }while($true)            

    }

}

function VisibleWirelessNetworks{
    Write-host "Starting VisibleWirelessNetwork Jobs"
    if($localBox){
        $netshresults = (netsh wlan show networks mode=bssid);
                $networksarraylist = [System.Collections.ArrayList]@();
                if((($netshresults.gettype()).basetype.name -eq "Array") -and ($netshresults.count -gt 10)){
                    for($i = 4; $i -lt ($netshresults.Length); $i+=11){
                        $WLANobject = [PSCustomObject]@{
                            SSID = ""
                            NetworkType = ""
                            Authentication = ""
                            Encryption = ""
                            BSSID = ""
                            SignalPercentage = ""
                            RadioType = ""
                            Channel = ""
                            BasicRates = ""
                            OtherRates = ""
                        }
                        for($j=0;$j -lt 10;$j++){
                            $currentline = $netshresults[$i + $j]
                            if($currentline -like "SSID*"){
                                $currentline = $currentline.substring(9)
                                if($currentline.startswith(" ")){

                                    $currentline = $currentline.substring(1)
                                    $WLANobject.SSID = $currentline

                                }else{

                                    $WLANobject.SSID = $currentline

                                }

                            }elseif($currentline -like "*Network type*"){

                                $WLANobject.NetworkType = $currentline.Substring(30)

                            }elseif($currentline -like "*Authentication*"){

                                $WLANobject.Authentication = $currentline.Substring(30)

                            }elseif($currentline -like "*Encryption*"){

                                $WLANobject.Encryption = $currentline.Substring(30)

                            }elseif($currentline -like "*BSSID 1*"){

                                $WLANobject.BSSID = $currentline.Substring(30)

                            }elseif($currentline -like "*Signal*"){

                                $WLANobject.SignalPercentage = $currentline.Substring(30)

                            }elseif($currentline -like "*Radio type*"){
        
                                $WLANobject.RadioType = $currentline.Substring(30)
        
                            }elseif($currentline -like "*Channel*"){
            
                                $WLANobject.Channel = $currentline.Substring(30)
                            }elseif($currentline -like "*Basic rates*"){
        
                                $WLANobject.BasicRates = $currentline.Substring(30)

                            }elseif($currentline -like "*Other rates*"){
            
                                $WLANobject.OtherRates = $currentline.Substring(30)

                            }
                        }

                        $networksarraylist.Add($WLANobject) | Out-Null
                    }
                    $networksarraylist | Export-Csv -NoTypeInformation -Append "$outFolder\Local_VisibleWirelessNetworks.csv" | out-null
                }
    }else{
        foreach($i in (Get-PSSession)){
            (Invoke-Command -session $i -ScriptBlock{
                $netshresults = (netsh wlan show networks mode=bssid);
                $networksarraylist = [System.Collections.ArrayList]@();
                if((($netshresults.gettype()).basetype.name -eq "Array") -and ($netshresults.count -gt 10)){
                    for($i = 4; $i -lt ($netshresults.Length); $i+=11){
                        $WLANobject = [PSCustomObject]@{
                            SSID = ""
                            NetworkType = ""
                            Authentication = ""
                            Encryption = ""
                            BSSID = ""
                            SignalPercentage = ""
                            RadioType = ""
                            Channel = ""
                            BasicRates = ""
                            OtherRates = ""
                        }
                        for($j=0;$j -lt 10;$j++){
                            $currentline = $netshresults[$i + $j]
                            if($currentline -like "SSID*"){
                                $currentline = $currentline.substring(9)
                                if($currentline.startswith(" ")){

                                    $currentline = $currentline.substring(1)
                                    $WLANobject.SSID = $currentline

                                }else{

                                    $WLANobject.SSID = $currentline

                                }

                            }elseif($currentline -like "*Network type*"){

                                $WLANobject.NetworkType = $currentline.Substring(30)

                            }elseif($currentline -like "*Authentication*"){

                                $WLANobject.Authentication = $currentline.Substring(30)

                            }elseif($currentline -like "*Encryption*"){

                                $WLANobject.Encryption = $currentline.Substring(30)

                            }elseif($currentline -like "*BSSID 1*"){

                                $WLANobject.BSSID = $currentline.Substring(30)

                            }elseif($currentline -like "*Signal*"){

                                $WLANobject.SignalPercentage = $currentline.Substring(30)

                            }elseif($currentline -like "*Radio type*"){
        
                                $WLANobject.RadioType = $currentline.Substring(30)
        
                            }elseif($currentline -like "*Channel*"){
            
                                $WLANobject.Channel = $currentline.Substring(30)
                            }elseif($currentline -like "*Basic rates*"){
        
                                $WLANobject.BasicRates = $currentline.Substring(30)

                            }elseif($currentline -like "*Other rates*"){
            
                                $WLANobject.OtherRates = $currentline.Substring(30)

                            }
                        }

                        $networksarraylist.Add($WLANobject) | Out-Null
                    }
                    $networksarraylist
                }
                                 
            }  -asjob -jobname "VisibleWirelessNetworks")
        }
        Create-Artifact
    }

}

function HistoricalWiFiConnections{
    Write-host "Starting HistoricalWiFiConnections Jobs"
    if($localBox){
        $netshresults = (netsh wlan show profiles);
                $networksarraylist = [System.Collections.ArrayList]@();
                if((($netshresults.gettype()).basetype.name -eq "Array") -and (!($netshresults[9].contains("<None>")))){
                    for($i = 9;$i -lt ($netshresults.Length -1);$i++){
                        $WLANProfileObject = [PSCustomObject]@{
                            ProfileName = ""
                            Type = ""
                            ConnectionMode = ""
                        }
                        $WLANProfileObject.profilename = $netshresults[$i].Substring(27)
                        $networksarraylist.Add($WLANProfileObject) | out-null
                        $individualProfile = (netsh wlan show profiles name="$($WLANProfileObject.ProfileName)")
                        $WLANProfileObject.type = $individualProfile[9].Substring(29)
                        $WLANProfileObject.connectionmode = $individualProfile[12].substring(29)
                    }
                }
                $networksarraylist | Export-Csv -NoTypeInformation -Append "$outFolder\Local_HistoricalWiFiConnections.csv" | out-null
    }else{
        foreach($i in (Get-PSSession)){
            (Invoke-Command -session $i -ScriptBlock{
                $netshresults = (netsh wlan show profiles);
                $networksarraylist = [System.Collections.ArrayList]@();
                if((($netshresults.gettype()).basetype.name -eq "Array") -and (!($netshresults[9].contains("<None>")))){
                    for($i = 9;$i -lt ($netshresults.Length -1);$i++){
                        $WLANProfileObject = [PSCustomObject]@{
                            ProfileName = ""
                            Type = ""
                            ConnectionMode = ""
                        }
                        $WLANProfileObject.profilename = $netshresults[$i].Substring(27)
                        $networksarraylist.Add($WLANProfileObject) | out-null
                        $individualProfile = (netsh wlan show profiles name="$($WLANProfileObject.ProfileName)")
                        $WLANProfileObject.type = $individualProfile[9].Substring(29)
                        $WLANProfileObject.connectionmode = $individualProfile[12].substring(29)
                    }
                }
                $networksarraylist
            
            } -AsJob -JobName "HistoricalWiFiConnections")
        }
        Create-Artifact
    }
}

function Enable-PSRemoting{
    
    foreach($node in $nodeList){
        wmic /node:$($node.IpAddress) process call create "powershell enable-psremoting -force"
    }
}

function HistoricalFirewallChanges{
    
    Write-host "Starting HistoricalFirewallChanges Jobs"
    if($localBox){
        Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall';} | select timecreated,message | export-csv -NoTypeInformation -Append "$outFolder\Local_HistoricalFirewallChanges.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall';} | select TimeCreated, Message}  -asjob -jobname "HistoricalFirewallChanges") | out-null
        }
        Create-Artifact
    }

}

function PortProxy{
    
    Write-host "Starting PortProxy Jobs"
    if($localBox){
         
    }else{    
        foreach($i in (Get-PSSession)){           
           (Invoke-Command -session $i -ScriptBlock{
                $portproxyResults = (netsh interface portproxy show all);
                $portproxyarraylist = [System.Collections.ArrayList]@();
                if((($portproxyResults.gettype()).basetype.name -eq "Array") -and ($portproxyResults.count -gt 0)){
                    for($i = 5; $i -lt ($portproxyResults.Length); $i++){
                        $portproxyObject = [PSCustomObject]@{
                            proxy = ""
                        }
                        $portproxyObject.proxy = $portproxyResults[$i]

                        $portproxyarraylist.Add($portproxyObject) | Out-Null
                    }
                    $portproxyarraylist
                }
                                 
            } -asjob -jobname "PortProxy") | out-null
        }
        Create-Artifact
    }
}

function CapabilityAccessManager{
    
    Write-host "Starting CapabilityAccessManager Jobs"
    if($localBox){
        (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\*\NonPackaged\*) | export-csv -NoTypeInformation -Append "$outFolder\Local_CapabilityAccessManager.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\*\NonPackaged\*)}  -asjob -jobname "CapabilityAccessManager") | out-null
        }
        Create-Artifact
    }

}

function DnsClientServerAddress{
    
    Write-host "Starting DnsClientServerAddress Jobs"
    if($localBox){
        (get-DnsClientServerAddress) | export-csv -NoTypeInformation -Append "$outFolder\Local_DnsClientServerAddress.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {Get-DnsClientServerAddress}  -asjob -jobname "DnsClientServerAddress") | out-null
        }
        Create-Artifact
    }

}

function ScheduledTaskDetails{
    
    Write-host "Starting ScheduledTaskDetails Jobs"
    if($localBox){
        ((Get-ScheduledTask).actions | ?{$_.execute -ne $null} |select execute,arguments) | export-csv -NoTypeInformation -Append "$outFolder\Local_ScheduledTaskDetails.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {(Get-ScheduledTask).actions | ?{$_.execute -ne $null} |select execute,arguments}  -asjob -jobname "ScheduledTaskDetails") | out-null
        }
        Create-Artifact
    }

}

<#
    T1023 
#>
function ShortcutModification{
    
    Write-host "Starting ShortcutModification Jobs"
    if($localBox){
        Select-String -Path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*.lnk" -Pattern "exe"| export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
        Select-String -Path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*.lnk" -Pattern "dll"| export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
        Select-String -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" -Pattern "dll"| export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
        Select-String -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" -Pattern "exe" | export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
        gci -path "C:\Users\" -recurse -include *.lnk -ea SilentlyContinue | Select-String -Pattern "exe" | export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
        gci -path "C:\Users\" -recurse -include *.lnk -ea SilentlyContinue | Select-String -Pattern "dll" | export-csv -NoTypeInformation -Append "$outFolder\Local_ShortcutModification.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {
                Select-String -Path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*.lnk" -Pattern "exe";
                Select-String -Path "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\*.lnk" -Pattern "dll";
                Select-String -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" -Pattern "dll";
                Select-String -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*" -Pattern "exe";
                gci -path "C:\Users\" -recurse -include *.lnk -ea SilentlyContinue | Select-String -Pattern "exe";
                gci -path "C:\Users\" -recurse -include *.lnk -ea SilentlyContinue | Select-String -Pattern "dll";
            }  -asjob -jobname "ShortcutModification") | out-null
        }
        Create-Artifact
    }

}

function DLLSinTempDirs{
    
    Write-host "Starting DLLSinTempDirs Jobs"
    if($localBox){
        (gps -Module -ea 0).FileName|?{$_ -notlike "*system32*"}|Select-String "Appdata","ProgramData","Temp","Users","public"|unique | export-csv -NoTypeInformation -Append "$outFolder\Local_DLLSinTempDirs.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock  {(gps -Module -ea 0).FileName|?{$_ -notlike "*system32*"}|Select-String "Appdata","ProgramData","Temp","Users","public"|unique}  -asjob -jobname "DLLSinTempDirs") | out-null
        }
        Create-Artifact
    }

}

function NamedPipes{
    
    Write-host "Starting NamedPipes Jobs"
    if($localBox){
        get-childitem \\.\pipe\ | export-csv -NoTypeInformation -Append "$outFolder\Local_NamedPipes.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock {get-childitem \\.\pipe\ | select fullname}  -asjob -jobname "NamedPipes") | out-null
        }
        Create-Artifact
    }

}

<#
    T1547.003
#>
function TimeProviders{
    
    Write-host "Starting TimeProviders Jobs"
    if($localBox){
         (Get-ItemProperty HKLM:\System\CurrentControlSet\Services\W32Time\TimeProviders\*) | export-csv -NoTypeInformation -Append "$outFolder\Local_TimeProviders.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock {(Get-ItemProperty HKLM:\System\CurrentControlSet\Services\W32Time\TimeProviders\*)}  -asjob -jobname "TimeProviders") | out-null
        }
        Create-Artifact
    }

}

function RDPHistoricallyConnectedIPs{
    
    Write-host "Starting RDPHistoricallyConnectedIPs Jobs"
    if($localBox){
        Get-WinEvent -Log 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' | select -exp Properties | where {$_.Value -like '*.*.*.*' } | sort Value -u  | export-csv -NoTypeInformation -Append "$outFolder\Local_RDPHistoricallyConnectedIPs.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock {Get-WinEvent -Log 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' | select -exp Properties | where {$_.Value -like '*.*.*.*' } | sort Value -u }  -asjob -jobname "RDPHistoricallyConnectedIPs") | out-null
        }
        Create-Artifact
    }

}

function CurveBall{
    
    Write-host "Starting CurveBall Jobs"
    if($localBox){
         Get-WinEvent -FilterHashtable @{logname="Application";id='1'} | ?{$_.providername -eq "Audit-CVE"} | export-csv -NoTypeInformation -Append "$outFolder\Local_CurveBall.csv" | Out-Null
    }else{    
        foreach($i in (Get-PSSession)){           
            (Invoke-Command -session $i -ScriptBlock {Get-WinEvent -FilterHashtable @{logname="Application";id='1'} | ?{$_.providername -eq "Audit-CVE"}}  -asjob -jobname "CurveBall") | out-null
        }
        Create-Artifact
    }

}

function PassTheHash{
    Write-host "Starting PassTheHash Jobs"
    if($localBox){
       
    }else{
        foreach($i in (Get-PSSession)){
            (Invoke-Command -session $i -ScriptBlock{
                $regexa = '.+Domain="(.+)",Name="(.+)"$';
                $regexd = '.+LogonId="(\d+)"$';
                $logon_users = @(Get-WmiObject win32_loggedonuser -ComputerName 'localhost');
                if(($logon_users -ne "") -and ($logon_users -ne $null)){
                    $session_user = @{};
                    $logon_users |% {
                        $_.antecedent -match $regexa > $nul;
                        $username = $matches[1] + "\" + $matches[2];
                        $_.dependent -match $regexd > $nul;
                        $session = $matches[1];
                        $sessionHex = ('0x{0:X}' -f [long]$session);
                        $session_user[$sessionHex] += $username ;
                    };

                    $klistsarraylist = [System.Collections.ArrayList]@();

                    foreach($i in $session_user.keys){

                        $item = $session_user.item($i).split("\")[1]    

                        $klistoutput = klist -li $i

                        if(($klistsarraylist -ne $null) -and ($klistoutput.count -gt 7)){
        
                            $numofrecords = $klistoutput[4].split("(")[1]
                            $numofrecords = $numofrecords.Substring(0,$numofrecords.Length-1)        

                            for($j = 0; $j -lt ($numofrecords);$j++){
                                $klistObject = [PSCustomObject]@{
                                                Session = ""
                                                Username = ""
                                                Client = ""
                                                Server = ""
                                                KerbTicketEncryptionType = ""
                                                StartTime = ""
                                                EndTime = ""
                                                RenewTime = ""
                                                SessionKeyType = ""
                                                CacheFlags = ""
                                                KdcCalled = ""
                                            }

                                    $klistObject.session = $i
                                    $klistObject.username = $item
                                    $klistObject.client = $klistoutput[6 + ($j * 11)].substring(12)
                                    $klistobject.server = $klistoutput[7 + ($j * 11)].substring(9)
                                    $klistobject.KerbTicketEncryptionType = $klistoutput[8 + ($j * 11)].substring(29)
                                    $klistobject.StartTime = $klistoutput[10 + ($j * 11)].substring(13)
                                    $klistobject.EndTime = $klistoutput[11 + ($j * 11)].substring(13)
                                    $klistobject.Renewtime = $klistoutput[12 + ($j * 11)].substring(13)
                                    $klistobject.sessionkeytype = $klistoutput[13 + ($j * 11)].substring(13)
                                    $klistobject.cacheflags = $klistoutput[14 + ($j * 11)].substring(14)
                                    $klistobject.kdccalled = $klistoutput[15 + ($j * 11)].substring(13)

                                    $klistsarraylist.Add($klistObject) | out-null
                            }
                        }else{
                            continue
                        }
                    }
                }
                $klistsarraylist

                }  -asjob -jobname "PassTheHash") | out-null
        }
        Create-Artifact
    }

}

function Meta-Blue {    
    <#
        This is the data gathering portion of this script. PSSessions are created on all live windows
        boxes. In order to create a new query, copy and paste and existing one. Change the write-host
        output to reflect the query's actions as well as the jobname parameter. Every 3rd query or so,
        add a call to Create-Artifact. This really impacts machines with small amounts of RAM.
    #>
    Build-Directories
    Build-Sessions
        
    <#
        Begining the artifact collection. Will start one job per session and then wait for all jobs
        of that type to complete before moving on to the next set of jobs.
    #>
    Write-host -ForegroundColor Green "[+]Begin Artifact Gathering"  
    Write-Host -ForegroundColor Yellow "[+]Sometimes the Powershell window needs you to click on it and press enter"
    Write-Host -ForegroundColor Yellow "[+]If it doesn't move on for a while, give it a try!"
    Write-Host -ForegroundColor Yellow "[+]Someone figure out how to make this not happen and I will give you a cookie" 
    
    Hunt-SolarBurst
    CurveBall
    PortProxy
    UACBypassFodHelper
    ArpCache
    #PowershellProfile
    TimeProviders
    DLLSearchOrderHijacking
    StartupFolder
    PassTheHash
    WebShell    
    UnsignedDrivers
    VisibleWirelessNetworks
    HistoricalWiFiConnections
    PoshVersion
    Registry
    SMBConns
    WMIEventFilters
    UserInitMprLogonScript
    WMIEventConsumers
    NetshHelperDLL
    WMIEventConsumerBinds
    LogicalDisks
    KnownDLLs
    DiskDrives
    SystemInfo
    SMBShares
    SystemFirmware
    AVProduct
    PortMonitors
    Startup
    Hotfix
    NetAdapters
    AccessibilityFeature
    DNSCache
    Logons
    LSASSDriver
    ProcessHash
    AlternateDataStreams
    NetConns
    EnvironVars
    DriverHash
    SchedTasks
    ScheduledTaskDetails
    PNPDevices 
    InstalledSoftware
    PrefetchListing
    Processes
    Services
    DLLHash
    Drivers
    BITSJobs
    HistoricalFirewallChanges
    CapabilityAccessManager
    DnsClientServerAddress
    NewlyRegisteredServices
    ShortcutModification
    DLLSinTempDirs
    #NamedPipes
    RDPHistoricallyConnectedIPs
    #ProgramData
    #DLLs
    #Update-Sysmon
    #Find-File
    TearDown-Sessions
    WaitFor-Jobs
    #Shipto-Splunk
    cd $outFolder
         
}

function Show-TitleMenu{
     cls
     Write-Host "================META-BLUE================"
     Write-Host "Tabs over spaces. Ain't nothin but a G thang"
    
     Write-Host "1: Press '1' to run Meta-Blue as enumeration only."
     Write-Host "2: Press '2' to run Meta-Blue as both enumeration and artifact collection."
     Write-Host "3: Press '3' to audit snort rules."
     Write-Host "4: Press '4' to remotely perform dump."
     Write-Host "5: Press '5' to run Meta-Blue against the local box."
     Write-Host "6: Press '6' to generate an IP space text file."
     Write-Host "Q: Press 'Q' to quit."
    
     $input = Read-Host "Please make a selection (title)"
     switch ($input)
     {
           '1' {
                cls
                show-EnumMenu
                break
           } '2' {
                cls
                Show-CollectionMenu
                break
           } '3' {
                cls                
                Audit-Snort
                break    
           } '4'{
                cls
                Show-MemoryDumpMenu
                break
           
           }'5'{
                $localBox = $true
                Meta-Blue
                break
           
           }'6'{
                Generate-IPSpaceTextFile
                Write-Host -ForegroundColor Green "[+]File saved to $($outFolder)\ipspace.txt"
                cd $outFolder
                break
           }

            'q' {
                break 
           } 

     }break
    
}

function Show-EnumMenu{
     
     cls
     Write-Host "================META-BLUE================"
     Write-Host "============Enumeration Only ================"
     Write-Host "      Do you have a list of hosts?"
     Write-Host "1: Yes"
     Write-Host "2: No"
     Write-Host "3: Return to previous menu."
     Write-Host "Q: Press 'Q' to quit."

                do{
                $input = Read-Host "Please make a selection(enum)"
                switch ($input)
                {
                    '1' {
                            $PTL = [System.Collections.arraylist]@()
                            $ptlFile = get-filename                        
                            if($ptlFile -eq ""){
                                Write-warning "Not a valid path!"
                                pause
                                show-enummenu
                            }
                            if($ptlFile -like "*.csv"){
                                $ptlimport = import-csv $ptlFile
                                foreach($ip in $ptlimport){$PTL.Add($ip.ipaddress) | out-null}
                                Enumerator($PTL)
                            }if($ptlFile -like "*.txt"){
                                $PTL = Get-Content $ptlFile
                                Enumerator($PTL)
                            }
                            break
                        }
                    '2'{
                            Write-Host "Running the default scan"
                            $subnets = Read-Host "How many seperate subnets do you want to scan?"

                            $ips = @()

                            for($i = 0; $i -lt $subnets; $i++){
                                $ipa = Read-Host "[$($i +1)]Please enter the network id to scan"
                                $cidr = Read-Host "[$($i +1)]Please enter the CIDR"
                                $ips += Get-SubnetRange -IPAddress $ipa -CIDR $cidr
                            }
                            Enumerator($ips)
                            break
                        }
                    '3'{
                            Show-TitleMenu
                            break
                    }
                    'q' {
                            break
                        }
                }
            }until ($input -eq 'q')
}

function Show-CollectionMenu{
    cls
    Write-Host "================META-BLUE================"
    Write-Host "============Artifact Collection ================"
    Write-Host "          Please Make a Selection               "
    Write-Host "1: Collect from a list of hosts"
    Write-Host "2: Collect from a network enumeration"
    Write-Host "3: Collect from active directory list (RSAT required!!)"
    Write-Host "4: Return to Previous menu."
    Write-Host "Q: Press 'Q' to quit."

                do{
                $input = Read-Host "Please make a selection(collection)"
                switch ($input)
                {
                    '1' {
                            $PTL = [System.Collections.arraylist]@()
                            $ptlFile = get-filename                        
                            if($ptlFile -eq ""){
                                Write-warning "Not a valid path!"
                                pause
                                show-enummenu
                            }
                            if($ptlFile -like "*.csv"){
                                $ptlimport = import-csv $ptlFile
                                foreach($node in $ptlimport){
                                    if($node.OperatingSystem -like "*windows*"){
                                        $nodeObj = [PSCustomObject]@{
                                            HostName = ""
                                            IPAddress = ""
                                            OperatingSystem = ""
                                            TTL = 0
                                        }
                                        $nodeObj.Hostname = $node.hostname
                                        $nodeObj.IPaddress = $node.IPAddress
                                        $nodeObj.OperatingSystem = $node.OperatingSystem
                                        $nodeObj.TTL = $node.TTL
                                        $nodeList.Add($nodeObj) | out-null
                                    }
                                }
                                
                                Build-Sessions
                            }if($ptlFile -like "*.txt"){
                                $PTL = Get-Content $ptlFile
                                Enumerator($PTL)
                            }
                            Meta-Blue
                            break
                        }
                    '2'{
                            Write-Host "Running the default scan"
                            $subnets = Read-Host "How many seperate subnets do you want to scan?"

                            $ips = @()

                            for($i = 0; $i -lt $subnets; $i++){
                                $ipa = Read-Host "[$($i +1)]Please enter the network id to scan"
                                $cidr = Read-Host "[$($i +1)]Please enter the CIDR"
                                $ips += Get-SubnetRange -IPAddress $ipa -CIDR $cidr
                            }
                            Enumerator($ips)
                            Meta-Blue
                            break
                        }
                    '3'{
                            $adEnumeration = $true
                            $iparray = (Get-ADComputer -filter *).dnshostname
                            Enumerator($iparray)
                            Meta-Blue
                            break                           
                    
                        }
                    '4'{
                            Show-TitleMenu
                            break
                    }
                    'q' {
                  
                            break
                        }
                }
            }until ($input -eq 'q')
}

function Show-MemoryDumpMenu{   
    do{
        cls
        Write-Host "================META-BLUE================"
        Write-Host "============Memory Dump ================"
        Write-Host "      Do you have a list of hosts?"
        Write-Host "1: Yes"
        Write-Host "2: Return to previous menu."
        Write-Host "Q: Press 'Q' to quit."
        $input = Read-Host "Please make a selection(dump)"
        switch ($input)
        {
            '1' {
                    $hostsToDump = [System.Collections.arraylist]@()
                    $hostsToDumpFile = get-filename                        
                    if($hostsToDumpFile -eq ""){
                        Write-warning "Not a valid path!"
                        pause
                        }else{
                        $dumpImport = import-csv $hostsToDumpFile
                        foreach($ip in $dumpImport){$hostsToDump.Add($ip.ipaddress) | out-null}
                        Enumerator($hostsToDump)
                        Memory-Dumper
                        break
                    }
                }
            '2'{
                    Show-TitleMenu
                    break
                }
            'q' {
                    break
                }
        }
    }until ($input -eq 'q')
}

show-titlemenu
