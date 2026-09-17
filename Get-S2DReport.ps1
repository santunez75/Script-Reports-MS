#requires -version 5.0

<#
.SYNOPSYS
    This script audit an S2D hyperconverged infrastructure
#>

##### Parameters #####
# ----------------------------------------------------------------

[CmdletBinding()]
param(
    [Parameter(Mandatory=$True, HelpMessage='Specify the name of the cluster')]
    [Alias('S2DCluster')]
    [string]$ClusterName,
    [Parameter(Mandatory=$True, HelpMessage='Specify the domain')]
    [Alias('ComputerDomain')]
    [String]$DomainName,
    [Parameter(Mandatory=$True, HelpMessage='Specify the folder where the report will be exported')]
    [Alias('ExportPath')]
    [String]$Path,
    [Parameter(Mandatory=$True, HelpMessage='Specify Credentials')]
    [Alias('Cred')]
    [PSCredential]$Credential,

    [Parameter(Mandatory=$False, HelpMessage='Specify customer name for the cover page')]
    [string]$Customer,

    [Parameter(Mandatory=$False, HelpMessage='Specify environment: Production, Development or Testing')]
    [ValidateSet('Production','Development','Testing')]
    [string]$Environment,

    [Parameter(Mandatory=$False, HelpMessage='Specify consultant name')]
    [string]$ConsultantName,

    [Parameter(Mandatory=$False, HelpMessage='Specify consultant email')]
    [string]$ConsultantEmail,

    [Parameter(Mandatory=$False, HelpMessage='Specify installation date (e.g. 2026-06-30)')]
    [string]$InstallDate,

    [Parameter(Mandatory=$False, HelpMessage='Specify country')]
    [string]$Country,

    [Parameter(Mandatory=$False, HelpMessage='Specify city')]
    [string]$City,

    [Parameter(Mandatory=$False, HelpMessage='Disable PDF export')]
    [switch]$NoPdf
    )


##### Function #####
# ----------------------------------------------------------------
Function Get-OSLanguage {
# This function returns the human comprehensive os language from the WMI Win32_OperatingSystem OSLanguage method
    Param([int]$Language)

    Switch ($Language){
        9 {$Lang='English'}
        1033 {$Lang='English &#8208; United States'}
        1034 {$Lang='Spanish &#8208; Traditional Sort'}
        2058 {$Lang='Spanish &#8208; Mexico'}
        3082 {$Lang='Spanish &#8208; International Sort'}
        9226 {$Lang=' Spanish &#8208; Colombia'}
        10250 {$Lang='Spanish &#8208; Peru'}
        11274 {$Lang='Spanish &#8208; Argentina'}
        13322 {$Lang='Spanish &#8208; Chile'}
    }
    Return $Lang
}

Function Get-VMHostHwInformation {
# This function collects Hyper-V hardware information and return an array
    Param([Array]$VMHosts,
          [String]$DomainName,
          [PSCredential]$Credential)

    $HWInformationArray = @()
    $i                  = 0

    Foreach ($VMHost in $VMHosts){
        # Show Progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) hardware information"
        $i++
        $CPUInfoArray = @()
        $VMHost       = $VMHost.Name + "." + $DomainName
        $VMHostObject = New-Object System.Object
        
        # Get Processor information
        $ComputerName = $env:ComputerName + "." + $DomainName

        # If the command is run on the local server, credential are not required
        If ($VMHost -like $ComputerName){
            $CPUs = Get-WmiObject -Class Win32_Processor -ComputerName $VMHost
        }
        else{
            $CPUs = Get-WmiObject -Class Win32_Processor -ComputerName $VMHost -Credential $Credential
        }

        # Collecting CPU information for each socket
        Foreach ($CPU in $CPUs){
            $CPUObj = New-Object System.Object
            $CPUObj | Add-Member -Type NoteProperty -Name Name -Value $CPU.Name
            $CPUObj | Add-Member -Type NoteProperty -Name DeviceId -Value $CPU.DeviceId
            $CPUObj | Add-Member -Type NoteProperty -Name NumberOfCores -Value $CPU.NumberOfCores
            $CPUObj | Add-Member -TYpe NoteProperty -Name NumberOfLogicalProcessors -Value $CPU.NumberOfLogicalProcessors
            $CPUInfoArray += $CPUObj
        }
        $VMHostObject | Add-Member -Type NoteProperty -Name CPU -Value $CPUInfoArray

        # Get Physical Memory information
        If ($VMHost -like $ComputerName){
            $PhysicalMemory = ((Get-WmiObject -Class Win32_ComputerSystem -ComputerName $VMHost).TotalPhysicalMemory)/1GB
        }
        else {
            $PhysicalMemory = ((Get-WmiObject -Class Win32_ComputerSystem -ComputerName $VMHost -Credential $Credential).TotalPhysicalMemory)/1GB
        }

        # Memory is round to 0 decimal
        $PhysicalMemory = [Math]::Round($PhysicalMemory, 0)
        $VMHostObject | Add-Member -Type NoteProperty -Name Memory -Value $PhysicalMemory
        $VMHostObject | Add-Member -Type NoteProperty -Name Name   -Value $VMHost
        

        # Get Virtual Machine Information
        $vCPU        = 0
        $MemAssigned = 0
        $WorkloadObj = New-Object System.Object
        # Azure Local 24H2: infrastructure VMs (Arc Resource Bridge, etc.) are excluded from the
        # workload/consolidation figures and reported separately
        $AllVMs      = Get-VM -ComputerName $VMHost
        $InfraVMs    = @($AllVMs | Where-Object { Test-InfrastructureVM -Name $_.Name -Path $_.Path })
        $VMs         = @($AllVMs | Where-Object { -not (Test-InfrastructureVM -Name $_.Name -Path $_.Path) })
        $VMs | Foreach {$vCPU += $_.ProcessorCount}
        $VMs | Foreach {$MemAssigned += $_.MemoryAssigned}
        $WorkloadObj  | Add-Member -Type NoteProperty -Name VMCount -Value $(($VMs | Measure-Object).Count)
        $WorkloadObj  | Add-Member -Type NoteProperty -Name InfraVMCount -Value $(($InfraVMs | Measure-Object).Count)
        $WorkloadObj  | Add-Member -Type NoteProperty -Name vCPU -Value $($vCPU)
        $WorkloadObj  | Add-Member -Type NoteProperty -Name MemAssigned -Value $($MemAssigned/1GB)
        $VMHostObject | Add-Member -Type NoteProperty -Name Workload -Value $WorkloadObj

        $HwInformationArray += $VMHostObject
    }
    Return $HWInformationArray

}

Function Get-VMHostHyperVSettings {
# This function returns Hyper-V settings
    Param([Array]$VMHosts,
          [string]$DomainName,
          [PSCredential]$Credential)
  
    $HyperVConfArray = @()
    $i               = 0
    $ComputerName    = Get-Content ENV:COMPUTERNAME

    # foreach node in the cluster
    Foreach ($VMHost in $VMHosts){

        # Show Progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) Hyper-V information"
       
        $Hostname = $VMHost.Name + "." + $DomainName

        # if the script is executed from the current node,get information locally
        If ($VMHost.Name -like $ComputerName){
           $HyperVSettings = Get-VMHost 
        }
        # else connecting remotely to the node with credential to get Hyper-V settings
        Else {
            $HyperVSettings = Invoke-Command -ComputerName $HostName -Credential $Credential {Get-VMHost}
        }

        # Add information to an array
        $HyperVObj = New-Object System.Object
        $HyperVObj | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost.Name
        $HyperVObj | Add-Member -Type NoteProperty -Name VMPath -Value $HyperVSettings.VirtualMachinePath
        $HyperVObj | Add-Member -Type NoteProperty -Name VHDPath -Value $HyperVSettings.VirtualHardDiskPath
        $HyperVObj | Add-Member -Type NoteProperty -Name MaximumLM -Value $HyperVSettings.MaximumVirtualMachineMigrations
        $HyperVObj | Add-Member -Type NoteProperty -Name MaximumStoMig -Value $HyperVSettings.MaximumStorageMigrations
        $HyperVObj | Add-Member -Type NoteProperty -Name LMAuthentication -Value $HyperVSettings.VirtualMachineMigrationAuthenticationType
        $HyperVObj | Add-Member -Type NoteProperty -Name LMPerformanceOption -Value $HyperVSettings.VirtualMachineMigrationPerformanceOption
        $HyperVConfArray += $HyperVObj

        $i++
    }
    Return $HyperVConfArray
}
    
Function Get-VMHostStorage {
# This function collects local Hyper-V host storage information and return an array
    Param([Array]$VMHosts,
          [String]$DomainName,
          [PSCredential]$Credential)

    $StoInformationArray = @()
    $i                   = 0
    Foreach ($VMHost in $VMHosts){
        # Show a progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) storage information"
        $i++
        $StoArray     = @()
        $VMHost       = $VMHost.Name + "." + $DomainName
        $StoObject = New-Object System.Object

        # if the script is run on a local computer, there is no need of credential
        $ComputerName = $env:ComputerName + "." + $DomainName
        If ($VMHost -like $ComputerName){
            $LocalStorage = Invoke-Command -ComputerName $VMHost {Get-Volume | Where-Object {($_.FileSystem -like "NTFS") -or ($_.FileSystem -like "REFS")}}
        }
        Else{
            $LocalStorage = Invoke-Command -Credential $Credential -ComputerName $VMHost {Get-Volume | Where-Object {($_.FileSystem -like "NTFS") -or ($_.FileSystem -like "REFS")}}
        }
        $StoObject | Add-Member -Type NoteProperty -Name Name -Value $VMHost

        # For each storage device, collecting its information (size, drive label, file system and so on)
        Foreach ($Storage in $LocalStorage){
            $StoObj    = New-Object System.Object
            $StoObj    | Add-Member -Type NoteProperty -Name DriveLetter -Value $Storage.DriveLetter
            $StoObj    | Add-Member -Type NoteProperty -Name FSLabel -Value $Storage.FileSystemLabel
            $StoObj    | Add-Member -Type NoteProperty -Name FileSystem -Value $Storage.FileSystem
            $StoObj    | Add-Member -Type NoteProperty -Name SizeRemaining -Value $Storage.SizeRemaining
            $StoObj    | Add-Member -Type NoteProperty -Name Size -Value $Storage.Size
            $StoObj    | Add-Member -Type NoteProperty -Name HealthStatus -Value $Storage.HealthStatus
            $StoArray += $StoObj       
        }
        $StoObject | Add-Member -Type NoteProperty -Name StorageInformation -Value $StoArray
        $StoInformationArray += $StoObject
        
    }
    Return $StoInformationArray
}

Function Get-VMHostNetwork {
#Get information about VM Host network adapter
    Param([Array]$VMHosts,
          [string]$DomainName,
          [PSCredential]$Credential)

    $VMHostNetworkArray = @()
    $ComputerName = Get-Content ENV:COMPUTERNAME
    $i = 0
    Foreach ($VMHost in $VMHosts){
        $Local = $Null

        # Show Progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) network adapter information"

        $i++
        # if the script is run on a node, don't need any credential and get information locally
        if ($VMHost.Name -like $ComputerName){
            $NICs  = Get-NetAdapter
            $Local = $True
        }

        # if the script is run remotely, get information with credential remotely.
        Else {
            $Cim   = New-CimSession -ComputerName $($VMHost.Name + "." + $DomainName) -Credential $Credential
            $NICs  = Get-NetAdapter -CimSession $Cim
            $Local = $False
        }

        # for each network adapter
        Foreach ($NIC in $NICs) {
            $NicObj = New-Object System.Object
            $NicObj | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost.Name
            $NicObj | Add-Member -Type NoteProperty -Name Name -Value $NIC.Name
            $NicObj | Add-Member -Type NoteProperty -Name Description -Value $NIC.InterfaceDescription
            $NicObj | Add-Member -Type NoteProperty -Name LinkSpeed -Value $Nic.LinkSpeed

            # if the script is run locally
            If ($Local){
                $RDMA = Get-NetAdapterRDMA -Name $Nic.Name -ErrorAction SilentlyContinue
                $MTU  = (Get-NetAdapterAdvancedProperty -Name $NIC.Name |? RegistryKeyword -like *Jumbo*).RegistryValue
                $RSS  = Get-NetAdapterRSS -Name $Nic.Name
                
                $IPaddress = Get-NetIPAddress -InterfaceAlias $NIC.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
                
                $DefaultGw = (Get-NetRoute |? DestinationPrefix -like 0.0.0.0/0 |? InterfaceAlias -like $NIC.Name).NextHop
                $DNS       = Get-DnsClientServerAddress -InterfaceAlias $NIC.Name -ErrorAction SilentlyContinue | Select -Expand ServerAddresses
                $RegisterDNS = Get-DnsClient -InterfaceAlias $NIC.Name  -ErrorAction SilentlyContinue | Select -Expand RegisterThisConnectionsAddress
                
                # if there is no IP address, don't need to show information
                if ($IPAddress.IPaddress -like $Null){
                    $NicObj | Add-Member -Type NoteProperty -Name IPAddress -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name Gateway -Value "GW: N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name DNS -Value "DNS: N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name DNSRegistration -Value "DNS Registration: N/a"
                }
                # else get IP information
                Else {
                    $Address = @()
                    Foreach ($IP in $IPAddress){
                        $Address += "$($IP.IPAddress)/$($IP.PrefixLength)"
                    }
                    $NicObj | Add-Member -Type NoteProperty -Name IPAddress -Value "$Address"
                    $NicObj | Add-Member -Type NoteProperty -Name Gateway -Value "GW: $DefaultGW"
                    $NicObj | Add-Member -Type NoteProperty -Name DNS -Value "DNS: $DNS"
                    $NicObj | Add-Member -Type NoteProperty -Name DNSRegistration -Value "DNS Registration: $RegisterDNS"
                }
                $NicObj | Add-Member -Type NoteProperty -Name RDMAState -Value $RDMA.Enabled
                $NicObj | Add-Member -Type NoteProperty -Name MTU -Value $MTU
                $NicObj | Add-Member -Type NoteProperty -Name RSSState -Value $RSS.Enabled
                $String = "$($RSS.BaseProcessorNumber) - $($RSS.MaxProcessorNumber) ($($RSS.MaxProcessors))"
                $NicObj | Add-Member -Type NoteProperty -Name RSS -Value $String
                

                # if it is a virtual interface
                if ($Nic.InterfaceDescription -like "*Hyper-V*"){

                    # gather information about vNIC
                    $vNIC = Get-VMNetworkAdapter -ManagementOS |? DeviceID -like $Nic.DeviceID
                    $VLAN = Get-VMNetworkAdapterVLAN -VMNetworkAdapterName $vNIC.Name -ManagementOS
                    $NicObj | Add-Member -Type NoteProperty -Name VMQState -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name VMQ -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name Type -Value Virtual
                    $NicObj | Add-Member -Type NoteProperty -Name VMMQ -Value $vNIC.VmmqEnabled
                    $NicObj | Add-Member -Type NoteProperty -Name SwitchName -Value $vNIC.SwitchName
                    $NicObj | Add-Member -Type NoteProperty -Name QoS -Value $vNIC.BandwidthPercentage

                    # Format text about VLAN
                    if ($VLAN.OperationMode -like "Untagged"){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Untagged"
                    }
                    Elseif ($Vlan.OperationMode -Like "Access"){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Access: $($VLAN.AccessVlanId)"
                    }
                    Else {
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Trunk: $($VLAN.AllowedVlanIdList) ($($VLAN.NativeVlanId))"
                    }
                    $TeamMapping = (Get-VMNetworkAdapterTeamMapping -ManagementOS -Name $vNIC.Name).NetAdapterName
                    if ($TeamMapping -like $Null){
                        $TeamMapping = "Not Configured"
                    }
                    $NicObj | Add-Member -Type NoteProperty -Name TeamMapping -Value $TeamMapping

                }

                # if it is a physical interface
                Else {
                    $VLAN = (Get-NetAdapterAdvancedProperty -Name $Nic.Name |? RegistryKeyword -like VlanID).RegistryValue
                    $VMQ  = Get-NetAdapterVMQ -Name $Nic.Name
                    $NicObj | Add-Member -Type NoteProperty -Name VMQState -Value $VMQ.Enabled
                    $String = "$($VMQ.BaseProcessorNumber) - $($VMQ.MaxProcessorNumber) ($($VMQ.MaxProcessors))"
                    $NicObj | Add-Member -Type NoteProperty -Name VMQ -Value $String
                    $NicObj | Add-Member -Type NoteProperty -Name Type -Value Physical
                    $NICObj | Add-Member -Type NoteProperty -Name VMMQ -Value "N/a"
                    $NICObj | Add-Member -Type NoteProperty -Name SwitchName -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name QoS -Value "N/a"


                    # Format VLAN text
                    if ($VLAN -eq 0){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Untagged"
                    }
                    Else {
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Access: $VLAN"
                    }
                    $NicObj | Add-Member -Type NoteProperty -Name TeamMapping -Value "N/a"
                }

            }

            # if it is a remote server, get information with cim session
            Else {
                $RDMA = Get-NetAdapterRDMA -Name $Nic.Name -CimSession $Cim -ErrorAction SilentlyContinue
                $MTU  = (Get-NetAdapterAdvancedProperty -Name $NIC.Name  -CimSession $Cim |? RegistryKeyword -like *Jumbo*).RegistryValue
                $RSS  = Get-NetAdapterRSS -Name $Nic.Name  -CimSession $Cim
                
                $IPaddress = Get-NetIPAddress -InterfaceAlias $NIC.Name -AddressFamily IPv4 -CimSession $Cim -ErrorAction SilentlyContinue

                $DefaultGw = (Get-NetRoute -CimSession $Cim |? DestinationPrefix -like 0.0.0.0/0 |? InterfaceAlias -like $NIC.Name).NextHop
                $DNS       = Get-DnsClientServerAddress -CimSession $Cim -InterfaceAlias $NIC.Name -ErrorAction SilentlyContinue | Select -Expand ServerAddresses
                $RegisterDNS = Get-DnsClient -InterfaceAlias $NIC.Name -ErrorAction SilentlyContinue -CimSession $Cim | Select -Expand RegisterThisConnectionsAddress
                
                # If there is no IP address, no need to gather information
                if ($IPAddress.IPAddress -like $Null){
 
                    $NicObj | Add-Member -Type NoteProperty -Name IPAddress -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name Gateway -Value "GW: N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name DNS -Value "DNS: N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name DNSRegistration -Value "DNS Registration: N/a"
                }

                # if there is an IP address, gather information
                Else {
                    $Address = @()
                    Foreach ($IP in $IPAddress){
                        $Address += "$($IP.IPAddress)/$($IP.PrefixLength)"
                    }
                    $NicObj | Add-Member -Type NoteProperty -Name IPAddress -Value "$Address"
                    $NicObj | Add-Member -Type NoteProperty -Name Gateway -Value "GW: $DefaultGW"
                    $NicObj | Add-Member -Type NoteProperty -Name DNS -Value "DNS: $DNS"
                    $NicObj | Add-Member -Type NoteProperty -Name DNSRegistration -Value "DNS Registration: $RegisterDNS"
                }
                
                $NicObj | Add-Member -Type NoteProperty -Name RDMAState -Value $RDMA.Enabled
                $NicObj | Add-Member -Type NoteProperty -Name MTU -Value $MTU
                $NicObj | Add-Member -Type NoteProperty -Name RSSState -Value $RSS.Enabled
                $String = "$($RSS.BaseProcessorNumber) - $($RSS.MaxProcessorNumber) ($($RSS.MaxProcessors))"
                $NicObj | Add-Member -Type NoteProperty -Name RSS -Value $String

                # if the NIC is a vNIC
                if ($Nic.InterfaceDescription -like "*Hyper-V*"){

                    # Run invoke-command to get remote information. Don't have choice because get-VMNetworkAdapter -cim -computer return error
                    $vNIC = Invoke-Command -ComputerName $($VMHost.Name + "." + $DomainName) -Credential $Credential -ArgumentList $Nic.DeviceId -ScriptBlock {
                            Get-VMNetworkAdapter -ManagementOS |? DeviceID -like $Args[0]}

                    $VLAN = Invoke-Command -ComputerName $($VMHost.Name + "." + $DomainName) -Credential $Credential -ArgumentList $vNIC.Name -ScriptBlock {
                            Get-VMNetworkAdapterVLAN -VMNetworkAdapterName $Args[0] -ManagementOS}
                    $NicObj | Add-Member -Type NoteProperty -Name VMQState -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name VMQ -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name Type -Value Virtual
                    $NicObj | Add-Member -Type NoteProperty -Name VMMQ -Value $vNIC.VmmqEnabled
                    $NicObj | Add-Member -Type NoteProperty -Name SwitchName -Value $vNIC.SwitchName
                    $NicObj | Add-Member -Type NoteProperty -Name QoS -Value $vNIC.BandwidthPercentage

                    # format VLAn text
                    if ($VLAN.OperationMode -like "Untagged"){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Untagged"
                    }
                    Elseif ($Vlan.OperationMode -Like "Access"){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Access: $($VLAN.AccessVlanId)"
                    }
                    Else {
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Trunk: $($VLAN.AllowedVlanIdList) ($($VLAN.NativeVlanId))"
                    }
                    $TeamMapping = Invoke-Command -ComputerName $($VMHost.Name + "." + $DomainName) -Credential $Credential -ArgumentList $vNIC.Name -ScriptBlock {
                                   (Get-VMNetworkAdapterTeamMapping -ManagementOS -Name $Args[0]).NetAdapterName}

                    if ($TeamMapping -like $Null){
                        $TeamMapping = "Not Configured"
                    }
                    $NicObj | Add-Member -Type NoteProperty -Name TeamMapping -Value $TeamMapping

                }

                # if it is a physical NIC
                Else {
                    $VLAN = (Get-NetAdapterAdvancedProperty -Name $Nic.Name -CimSession $Cim |? RegistryKeyword -like VlanID).RegistryValue
                    $VMQ  = Get-NetAdapterVMQ -Name $Nic.Name  -CimSession $Cim
                    $NicObj | Add-Member -Type NoteProperty -Name VMQState -Value $VMQ.Enabled
                    $String = "$($VMQ.BaseProcessorNumber) - $($VMQ.MaxProcessorNumber) ($($VMQ.MaxProcessors))"
                    $NicObj | Add-Member -Type NoteProperty -Name VMQ -Value $String
                    $NicObj | Add-Member -Type NoteProperty -Name Type -Value Physical
                    $NICObj | Add-Member -Type NoteProperty -Name VMMQ -Value "N/a"
                    $NICObj | Add-Member -Type NoteProperty -Name SwitchName -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name QoS -Value "N/a"
                    $NicObj | Add-Member -Type NoteProperty -Name TeamMapping -Value "N/a"
                    
                    # Format VLAN text
                    if ($VLAN -eq 0){
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Untagged"
                    }
                    Else {
                        $NicObj | Add-Member -Type NoteProperty -Name VLAN -Value "Access: $VLAN"
                    }

                }
            }
        $VMHostNetworkArray += $NicObj
        }
        if (!($Local)){
            Remove-CimSession $Cim
        }
    }

    Return $VMHostNetworkArray
}

Function Get-VMHostvSwitch {
# this function get information about VMSwitches
    Param([Array]$VMHosts,
          [string]$DomainName,
          [PSCredential]$Credential)

    $VMHostvSwitchArray = @()
    $i                  = 0

    # for each node in the cluster
    Foreach ($VMHost in $VMHosts){

        # Show Progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) VMSwitches information"

        $i++
              
        $Hostname = $VMHost.Name + "." + $DomainName  
        $ComputerName = Get-Content ENV:COMPUTERNAME

        # if the script is run on a node, run command locally
        if ($VMHost -like $ComputerName){
            $vSwitches = Get-VMSwitch
        }

        # if the script is remote, connect to the node with computername and credential
        Else {
        $vSwitches = Invoke-Command -ComputerName $Hostname -Credential $Credential {Get-VMSwitch}
        }

        # for each VMswitch, get information
        Foreach ($vSwitch in $vSwitches){
            $vSwitchObj = New-Object System.Object
            $vSwitchObj | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost.Name
            $vSwitchObj | Add-Member -Type NoteProperty -Name Name -Value $vSwitch.Name
            $vSwitchObj | Add-Member -Type NoteProperty -Name Type -Value $vSwitch.SwitchType
            $vSwitchObj | Add-Member -Type NoteProperty -Name QoSmode -Value $vSwitch.BandwidthReservationMode
            $vSwitchObj | Add-Member -Type NoteProperty -Name EmbeddedTeaming -Value $vSwitch.EmbeddedTeamingEnabled
            $vSwitchObj | Add-Member -Type NoteProperty -Name PacketDirect -Value $vSwitch.PacketDirectEnabled
            $vSwitchObj | Add-Member -Type NoteProperty -Name IovSupport -Value $vSwitch.IovEnabled

            # trying to get the NIC name instead of the nic description. If local, get information locally
            If ($VMHost -like $ComputerName){
                $Nics = $vSwitch | select -expand NetAdapterInterfaceDescriptions |
                        % {Get-NetAdapter |? InterfaceDescription -like $_ | Select -Expand Name }
            } 

            # if remote, get information with credential
            Else {
                $Nics = $vSwitch | select -expand NetAdapterInterfaceDescriptions
                $Nics = Invoke-Command -ComputerName $Hostname -Credential $Credential -ArgumentList $Nics {
                        $Args |% {Get-NetAdapter |? InterfaceDescription -like $_ | Select -Expand Name} }
            }
        $vSwitchObj | Add-Member -Type NoteProperty -Name NICs -Value $NICs
        $VMHostvSwitchArray += $vSwitchObj
        }
    }
    Return $VMHostvSwitchArray
}

Function Get-ClusterStorage {
# This function collects cluster storage shared between each nodes and return an array
    Param([String]$ClusterName)

    $StoClusterArray  = @()
    $ClusterStoObject = New-Object System.Object
    $ClusterSto       = Get-ClusterSharedVolume -Cluster $ClusterName
    $ClusterStoObject | Add-Member -Type NoteProperty -Name Name -Value $ClusterName
    $StoArray         = @()
    $i                = 0

    # For eachstorage device in $ClusterSto, collecting some information as name, state, size and so on
    Foreach ($Storage in $ClusterSto){
       #Show a Progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete 50 `
                       -CurrentOperation "Collecting $ClusterName storage information"

        $StoObj = New-Object System.Object
        $StoObj | Add-Member -Type NoteProperty -Name Name -Value $Storage.Name
        $StoObj | Add-Member -Type NoteProperty -Name State -Value $Storage.State

        # Get specific information as maintenance mode and friendlyname
        $VolumeState = Get-ClusterSharedVolume -Cluster $ClusterName -Name $Storage.Name | select -Expand SharedVolumeInfo
        $StoObj | Add-Member -Type NoteProperty -Name MaintenanceMode $VolumeState.MaintenanceMode
        $StoObj | Add-Member -Type NoteProperty -Name FriendlyVolumename $VolumeState.FriendlyVolumeName

        # Get some information about the volume (size and used space)
        $PartitionState = Get-ClusterSharedVolume -Cluster $ClusterName -Name $Storage.Name | select -Expand SharedVolumeInfo | select -Expand Partition
        $StoObj | Add-Member -Type NoteProperty -Name Size $PartitionState.Size
        $StoObj | Add-Member -Type NoteProperty -Name UsedSpace $PartitionState.UsedSpace
        $StoArray += $StoObj

    }
    $ClusterStoObject | Add-Member -Type NoteProperty -Name StorageInformation -Value $StoArray
    $StoClusterArray += $ClusterStoObject
    Return $StoClusterArray
}

Function Get-ClusterNetInformation {
# This function collects cluster network information and return an array   
    Param([string]$ClusterName)

    $ClusterNetInfoArray = @()
    # Collect cluster network information
    $ClusterNetInfo      = Get-ClusterNetwork -Cluster $ClusterName
    # Collect network ID which are not allowed to transmit Live-Migration flows
    $LiveMigrationNet    = (Get-ClusterResourceType -Cluster $ClusterName "Virtual Machine" | Get-ClusterParameter -Name MigrationExcludeNetworks).Value
    
    # Split network id in an array
    $LMNetArray          = $LiveMigrationNet -Split ";"
    $i                   = 0  

    # For each cluster network, collect network information
    Foreach ($Network in $ClusterNetInfo){
        # show a progress bar
        Write-Progress -Activity "HTML file construction" -PercentComplete (($i/$ClusterNetInfo.Count)*100) -CurrentOperation "Network information gathering"
        $NetInfoObject = New-object System.Object
        $NetInfoObject | Add-Member -Type NoteProperty -Name Name -Value $Network.Name
        $NetInfoObject | Add-Member -Type NoteProperty -Name Role -Value $Network.Role
        $NetInfoObject | Add-Member -Type NoteProperty -Name Address -Value $Network.Address
        $NetInfoObject | Add-Member -Type NoteProperty -Name AddressMask -Value $Network.AddressMask
        $NetInfoObject | Add-Member -Type NoteProperty -Name State -Value $Network.State
        
        $LMNet = $True
        # Verifying if this network is excluded to transmit Live-Migration flows
        Foreach ($ExcludeNet in $LMNetArray){

            if ($Network.Id -like $ExcludeNet){
                $LMNet = $False
            }
        }
        $NetInfoObject        | Add-Member -Type NoteProperty -Name LMNet -Value $LMNet
        $ClusterNetInfoArray += $NetInfoObject
        
    }
    Return $ClusterNetInfoArray

}

Function Get-VMHostOsInformation {
# Collect Hyper-V Host OS information
    Param([Array]$VMHosts,
          [String]$DomainName,
          [PSCredential]$Credential)

    $OSInformationArray = @()
    $i                  = 0
    $ComputerName = $env:ComputerName + "." + $DomainName

    # For each Hyper-V Host in $VMHosts array, collecting information
    Foreach ($VMHost in $VMHosts){

        # Show a Progress Bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting $($VMHost.Name) OS information"
        $i++
        $VMHost        = $VMHost.Name + "." + $DomainName
        $VMHostOSObj   = New-Object System.Object
        $VMHostOSObj   | Add-Member -Type NoteProperty -Name Name -Value $VMHost
        $FirewallArray = @()

        # Collect Firewall state on remote Hyper-V Host
        $FirewallState = Invoke-Command -ComputerName $VMHost -Credential $Credential -ScriptBlock {Get-NetFirewallProfile}
        
        # For each firewall profile, collect state
        Foreach ($Firewall in $FirewallState){
            $FirewallObj    = New-Object System.object
            $FirewallObj    | Add-Member -type NoteProperty -Name Name -Value $Firewall.Name
            $FirewallObj    | Add-Member -type NoteProperty -Name State -Value $Firewall.Enabled
            $FirewallArray += $FirewallObj
        }
        $VMHostOSObj  | Add-Member -Type NoteProperty -Name Firewall -Value $FirewallArray

        # verifying if the server is in minimal interface mode. If the script is launched on the local server, don't need credential
        If ($VMHost -like $ComputerName){
            $GuiState = (Get-WindowsFeature -Name Server-Gui-Shell -ComputerName $VMHost).Installed
        }
        else{
            $GuiState = (Get-WindowsFeature -Name Server-Gui-Shell -ComputerName $VMHost -Credential $Credential).Installed
        }
   
        $VMHostOSObj  | Add-Member -Type NoteProperty -Name GUIInstalled -Value $GuiState

        # Collecting OS information in Win32_OperatingSystem WMI class
        If ($VMHost -like $ComputerName){
            $OSInfo = Get-WmiObject -ComputerName $VMHost -Class Win32_OperatingSystem
        }
        else{
            $OSInfo = Get-WmiObject -ComputerName $VMHost -Class Win32_OperatingSystem -Credential $Credential
        }
        
        # Count hotfix installed
        If ($VMHost -like $ComputerName){
            $HotFixCount = (Get-HotFix -ComputerName $VMhost).Count
        }
        else{
            $HotFixCount = (Get-HotFix -ComputerName $VMhost -Credential $Credential).Count
        }
        
        $VMHostOSObj  | Add-Member -Type NoteProperty -Name OSVersion -Value $OSInfo.Caption
        $VMHostOSObj  | Add-Member -Type NoteProperty -Name OSLanguage -Value $OSInfo.OSLanguage
        $VMHostOSObj  | Add-Member -Type NoteProperty -Name OSHotfix -Value $HotFixCount
        $OSInformationArray += $VMHostOSObj
    }
    Return $OSInformationArray

}

Function Get-ClusterConfInformation {
# Collecting cluster information
    Param([string]$ClusterName)

    # Show a Progress bar
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                   -PercentComplete 50 `
                   -CurrentOperation "Collecting $ClusterName information"

    $ClustConfArray    = @()

    # Count the number of node in the cluster
    $NodeNbr           = (Get-ClusterNode -Cluster $ClusterName).Count
    # Collect cluster information
    $ClusterInfo       = Get-Cluster -Name $ClusterName
    # Collect Cluster Quorum information
    $QuorumInfo        = Get-ClusterQuorum -Cluster $ClusterName

    $ClustInfoObj      = New-Object System.Object
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name Name -Value $ClusterName
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name NodeNbr -Value $NodeNbr
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name QuorumType -Value $QuorumInfo.QuorumType
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name QuorumResource -Value $QuorumInfo.QuorumResource
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name WitnessDynamicWeight -Value $ClusterInfo.WitnessDynamicWeight
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name DynamicQuorum -Value $ClusterInfo.DynamicQuorum
    $ClustInfoObj      | Add-Member -Type NoteProperty -Name BlockCacheSize -Value $ClusterInfo.BlockCacheSize
    $ClusterConfArray += $ClustInfoObj

    Return $ClusterConfArray
}

Function Get-VMHostWorkloads {
# This function collects virtual machines information on Hyper-V Hosts and return an array
    Param([Array]$VMHosts,
          [string]$DomainName,
          [PSCredential]$Credential)

    $VMHostWorkloadArray = @()
    $i                   = 0

    # For each Hyper-V host in $VMHosts array, collecting VM information
    Foreach ($VMHost in $VMHosts){
        # Show a progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VMHosts.Count)*100) `
                       -CurrentOperation "Collecting VMs information on $($VMHost.Name)"
        $i++
        $VMHost       = $VMHost.Name + "." + $DomainName

        # Get Virtual Machines from remote Hyper-V nodes
        $VMs          = Invoke-Command -ComputerName $VMHost -Credential $Credential -ScriptBlock {Get-VM}

        # For each virtual machine in $VMs, collecting information
        Foreach ($VM in $VMs){
           $VMSnapArray  = @()
           $VMDiskArray  = @()

           $VMObject     = New-Object System.Object
           $VMObject     | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost 
           $VMObject     | Add-Member -Type NoteProperty -Name Name -Value $VM.Name
           $VMObject     | Add-Member -Type NoteProperty -Name State -Value $VM.State
           $VMObject     | Add-Member -Type NoteProperty -Name IsClustered -Value $VM.IsClustered
           $VMObject     | Add-Member -Type NoteProperty -Name IsInfrastructure -Value (Test-InfrastructureVM -Name $VM.Name -Path $VM.Path)
           $VMObject     | Add-Member -Type NoteProperty -Name Generation -Value $VM.Generation
           $VMObject     | Add-Member -Type NoteProperty -Name ProcessorCount -Value $VM.ProcessorCount
           $VMObject     | Add-Member -Type NoteProperty -Name DynamicMemoryEnabled -Value $VM.DynamicMemoryEnabled
           $VMObject     | Add-Member -Type NoteProperty -Name MemoryAssigned -Value $VM.MemoryAssigned
           $VMObject     | Add-Member -Type NoteProperty -Name MemoryDemand -Value $VM.MemoryDemand
           
           # collect VirtualHardDrive information
           $VMDisks = Invoke-Command -ComputerName $VMHost `
                                     -Credential $Credential `
                                     -ArgumentList $VM.Name `
                                     -ScriptBlock {Get-VMHardDiskDrive -VMName $Args[0]}
           
           # Collect Checkpoints information
           $VMSnaps = Invoke-Command -ComputerName $VMHost `
                                     -Credential $Credential `
                                     -ArgumentList $VM.Name `
                                     -ScriptBlock {Get-VMSnapshot -VMName $Args[0]}

           # for each virtual disk, collecting information
           Foreach ($vDisk in $VMDisks){
               $vDiskObject = New-Object System.Object
               $vDiskObject | Add-member -Type NoteProperty -Name ControllerType -Value $vDisk.ControllerType
               $vDiskObject | Add-member -Type NoteProperty -Name Path -Value $vDisk.Path

               # Collecting advanced virtual disk information
               $vDiskInfo    = Invoke-Command -ComputerName $VMHost -Credential $Credential -ArgumentList $vDisk.Path -ScriptBlock {Get-VHD -Path $Args[0]}

               $vDiskObject | Add-member -Type NoteProperty -Name VHDType -Value $vDiskInfo.VHDType
               $vDiskObject | Add-member -Type NoteProperty -Name Size -Value $vDiskInfo.Size

               $VMDiskArray += $vDiskObject
           }
           $VMObject | Add-Member -Type NoteProperty -Name VMDisks -Value $VMDiskArray

           # for each checkpoint, collecting information
           Foreach ($Checkpoint in $VMSnaps){
               $CheckpointObject = New-Object System.Object
               $CheckpointObject | Add-Member -Type NoteProperty -Name Name -Value $Checkpoint.Name
               $CheckpointObject | Add-Member -Type NoteProperty -Name CreationTime -Value $Checkpoint.CreationTime
               $VMSnapArray += $CheckPointObject
           }
           $VMObject | Add-Member -Type NoteProperty -Name Checkpoints -Value $VMSnapArray
           $VMHostWorkloadArray += $VMObject 
              
        }
          
    }
    Return $VMHostWorkloadArray

}

Function Get-StoragePoolInfo {
#this function gets information about Storage Pool
    Param([string]$ClusterName,
          [string]$DomainName,
          [PSCredential]$Credential)
    
    $StoragePoolInfo = @()

    $Cim             = New-CimSession -ComputerName $($ClusterName + "." + $DomainName) -Credential $Credential

    # Get Storage Pool which are not primordial
    $StoragePools     = Get-StorageSubSystem -CimSession $Cim |? Name -like *$ClusterName*| Get-StoragePool |? isPrimordial -like $False
    Foreach ($StoragePool in $StoragePools){   

        # for each storage pool, add information to an array
        $SSObj = New-Object System.Object
        $SSObj | Add-Member -Type NoteProperty -Name FriendlyName $StoragePool.FriendlyName
        $SSObj | Add-Member -Type NoteProperty -Name OperationalStatus $StoragePool.OperationalStatus
        $SSObj | Add-Member -Type NoteProperty -Name HealthStatus $StoragePool.HealthStatus
        $SSObj | Add-Member -Type NoteProperty -Name Size $StoragePool.Size
        $SSObj | Add-Member -Type NoteProperty -Name AllocatedSize $StoragePool.AllocatedSize
        $StoragePoolInfo += $SSObj
    }
    Remove-CimSession -CimSession $CIM
    Return $StoragePoolInfo
}

Function Get-VirtualDiskInfo {
# this function gets information about virtual disks
    Param([string]$ClusterName,
          [string]$DomainName,
          [PSCredential]$Credential)

    $VDArray = @()
    $Cim = New-CimSession -ComputerName $($ClusterName + "." + $DomainName)

    # get virtual disks information from storage subsystem with the cluster name and storage pool not primordial
    $VirtualDisks = Get-StorageSubSystem -CimSession $Cim |? Name -like *$ClusterName* | Get-StoragePool |? isPrimordial -like $False | Get-VirtualDisk
    $i = 0

    Foreach ($VirtualDisk in $VirtualDisks){

        # Show a progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (($i/$VirtualDisks.Count)*100) `
                       -CurrentOperation "Get virtual disk information on $($ClusterName)"
        $i++

        # for each virtual disk, gather information into an array
        $VDObj = New-Object System.Object
        $VDObj | Add-Member -Type NoteProperty -Name FriendlyName -Value $VirtualDisk.FriendlyName
        $VDObj | Add-Member -Type NoteProperty -Name NumberOfColumns -Value $VirtualDisk.NumberOfColumns
        $VDObj | Add-Member -Type NoteProperty -Name ResiliencySettingName -Value $VirtualDisk.ResiliencySettingName
        $VDObj | Add-Member -Type NoteProperty -Name NumberOfDataCopies -Value $VirtualDisk.NumberOfDataCopies
        $VDObj | Add-Member -Type NoteProperty -Name Size -Value $VirtualDisk.Size
        $VDObj | Add-Member -Type NoteProperty -Name FootprintOnPool -Value $VirtualDisk.FootprintOnPool
        $VDObj | Add-Member -Type NoteProperty -Name HealthStatus -Value $VirtualDisk.HealthStatus
        $VDArray += $VDObj
    }
    Remove-CimSession -CimSession $CIM
    Return $VDArray

}

Function Get-PhysicalDiskInfo {
# this function gets information about physical disks
    Param([string]$ClusterName,
          [string]$DomainName,
          [PSCredential]$Credential)

    $PhysicalDiskArray = @()
    $i                 = 0

    # connecting to an online node
    $Node = ((Get-ClusterNode -Cluster ($ClusterName + "." + $DomainName) |? State -Like up)[0]).Name
    $Node = $Node + "." + $DomainName

    # runnin the function remotely
    $PhysicalDiskArray = Invoke-Command -ComputerName $Node -Credential $Credential -Argumentlist $ClusterName -ScriptBlock {
       
        $PDArray      = @()
        $StoragePools = Get-StorageSubSystem |? Name -like *$($Args[0])* | Get-StoragePool |? isprimordial -like $false
       
        # for each storage pool
        Foreach ($StoragePool in $StoragePools){

            # Get physical disk in the storage pool
            $PhysicalDisks = $StoragePool | Get-PhysicalDisk
            # For each physical disk
            Foreach ($PhysicalDisk in $PhysicalDisks){

                # Show a progress bar
                Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                               -PercentComplete (($i/$PhysicalDisks.Count)*100) `
                               -CurrentOperation "Collecting physical disk information on $($ClusterName)"
                $i++
                
                # gather information about Physical disks
                $PDObj = New-Object System.Object
                $PDObj | Add-Member -Type NoteProperty -Name StoragePoolFriendlyName -Value $StoragePool.FriendlyName
                $PDObj | Add-Member -Type NoteProperty -Name FriendlyName -Value $PhysicalDisk.FriendlyName
                $PDObj | Add-Member -Type NoteProperty -Name FirmwareVersion -Value $PhysicalDisk.FirmwareVersion
                $PDObj | Add-Member -Type NoteProperty -Name Model -Value $PhysicalDisk.Model
                $PDObj | Add-Member -Type NoteProperty -Name SerialNumber -Value $PhysicalDisk.SerialNumber
                $PDObj | Add-Member -Type NoteProperty -Name Size -Value $PhysicalDisk.Size
                $PDObj | Add-Member -Type NoteProperty -Name AllocatedSize -Value $PhysicalDisk.AllocatedSize
                $PDObj | Add-Member -Type NoteProperty -Name MediaType -Value $PhysicalDisk.MediaType
                $PDObj | Add-Member -Type NoteProperty -Name BusType -Value $PhysicalDisk.BusType
                $PDObj | Add-Member -Type NoteProperty -Name HealthStatus -Value $PhysicalDisk.HealthStatus
                $PDObj | Add-Member -Type NoteProperty -Name OperationalStatus -Value $PhysicalDisk.OperationalStatus
                $PDObj | Add-Member -Type NoteProperty -Name Usage -Value $PhysicalDisk.Usage
                $PDArray += $PDObj
            }  
        }
        Return $PDArray

    }
    Return $PhysicalDiskArray

}

Function Get-VMHostSMBMultiChannel {
# This function get information about SMB MultiChannel
    Param([Array]$VMHosts,
          [String]$DomainName,
          [PSCredential]$Credential)

    $SMBMultiInfoArray = @()
    $ComputerName      = Get-Content ENV:COMPUTERNAME
    $i                 = 0

    # for each node
    Foreach ($VMHost in $VMHosts){
         # Show a progress bar
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                        -PercentComplete (($i/$VMHosts.Count)*100) `
                        -CurrentOperation "Collecting SMB MultiChannel information on $($VMHost.Name)"
        $i++
        
        # if the a node is running the script, connect locally
        if ($VMHost -like $ComputerName){
            $SBL = Get-SmbMultichannelConnection -SmbInstance SBL
            $CSV = Get-SMBMultiChannelConnection -SmbInstance CSV
        }

        # If the script is run remotely, connecting with credential and CIM session
        Else {
            $HostName = $VMHost.Name + "." + $DomainName
            $Cim = New-CimSession -ComputerName $hostname -Credential $Credential
            $SBL = Get-SmbMultichannelConnection -SmbInstance SBL -CimSession $Cim
            $CSV = Get-SMBMultiChannelConnection -SmbInstance CSV -CimSession $Cim
            
        }

        # for each connection SBL information, add them to array
        Foreach ($Connection in $SBL){
            $SMBInfoObj = New-Object System.Object
            $SMBInfoObj | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost.Name
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ConnectionType -Value "SBL"
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientIP -Value $Connection.ClientIpAddress
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerIP -Value $Connection.ServerIpAddress
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientNIC -Value $Connection.ClientInterfaceFriendlyName
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerNIC -Value $Connection.ServerInterfaceIndex
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientRSS -Value $Connection.ClientRSSCapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerRSS -Value $Connection.ServerRSSCapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientRDMA -Value $Connection.ClientRDMACapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerRDMA -Value $Connection.ServerRDMACapable
            $SMBMultiInfoArray += $SMBInfoObj
       
        }

        # For each CSV connection, add them to array
        Foreach ($Connection in $CSV){
            $SMBInfoObj = New-Object System.Object
            $SMBInfoObj | Add-Member -Type NoteProperty -Name VMHost -Value $VMHost.Name
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ConnectionType -Value "CSV"
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientIP -Value $Connection.ClientIpAddress
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerIP -Value $Connection.ServerIpAddress
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientNIC -Value $Connection.ClientInterfaceFriendlyName
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerNIC -Value $Connection.ServerInterfaceIndex
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientRSS -Value $Connection.ClientRSSCapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerRSS -Value $Connection.ServerRSSCapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ClientRDMA -Value $Connection.ClientRDMACapable
            $SMBInfoObj | Add-Member -Type NoteProperty -Name ServerRDMA -Value $Connection.ServerRDMACapable
            $SMBMultiInfoArray += $SMBInfoObj
        }
        
    }
    Return $SMBMultiInfoArray
}

Function ConvertTo-HtmlSafe {
    Param([AllowNull()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

Function Read-RequiredText {
    Param([string]$PromptText)
    do {
        $Value = Read-Host $PromptText
        if ([string]::IsNullOrWhiteSpace($Value)) {
            Write-Host 'This value is required.' -ForegroundColor Yellow
        }
    } while ([string]::IsNullOrWhiteSpace($Value))
    return $Value.Trim()
}

Function Read-EnvironmentOption {
    Write-Host ''
    Write-Host 'Select Environment:' -ForegroundColor Cyan
    Write-Host '  1. Production'
    Write-Host '  2. Development'
    Write-Host '  3. Testing'
    do {
        $Option = Read-Host 'Enter option [1-3]'
        switch ($Option) {
            '1' { return 'Production' }
            '2' { return 'Development' }
            '3' { return 'Testing' }
            default { Write-Host 'Invalid option. Use 1, 2 or 3.' -ForegroundColor Yellow }
        }
    } while ($true)
}

Function Read-InstallDate {
    do {
        $Value = Read-Host 'Installation date (e.g. 2026-06-30)'
        if ([string]::IsNullOrWhiteSpace($Value)) {
            Write-Host 'This value is required.' -ForegroundColor Yellow
            continue
        }
        $ParsedDate = $null
        if (-not [datetime]::TryParse($Value, [ref]$ParsedDate)) {
            Write-Host "Could not parse '$Value' as a date. Try formats like 2026-06-30 or 30/06/2026." -ForegroundColor Yellow
            $Value = $null
        }
    } while ([string]::IsNullOrWhiteSpace($Value))
    return $ParsedDate.ToString('dd-MMM-yyyy')
}

Function Convert-HtmlReportToPdf {
    Param(
        [Parameter(Mandatory=$true)][string]$HtmlPath,
        [Parameter(Mandatory=$true)][string]$PdfPath
    )

    $BrowserCandidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { $_ -and (Test-Path $_) }

    if (-not $BrowserCandidates -or $BrowserCandidates.Count -eq 0) {
        Write-Warning 'PDF was not generated because Microsoft Edge or Google Chrome was not found. Install Edge/Chrome or run without -NoPdf only where a supported browser exists.'
        return $false
    }

    $Browser = $BrowserCandidates[0]
    $HtmlUri = ([System.Uri](Resolve-Path $HtmlPath).Path).AbsoluteUri
    $Arguments = @(
        '--headless',
        '--disable-gpu',
        '--no-sandbox',
        '--print-to-pdf-no-header',
        "--print-to-pdf=$PdfPath",
        $HtmlUri
    )

    $Process = Start-Process -FilePath $Browser -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
    Start-Sleep -Seconds 2

    if ((Test-Path $PdfPath) -and ((Get-Item $PdfPath).Length -gt 0)) {
        return $true
    }

    Write-Warning "PDF was not generated. Browser exit code: $($Process.ExitCode)"
    return $false
}

<#

Function Get-S2DConfiguration {
    Param([String]$ClusterName,
          [String]$DomainName,
          [PSCredential]$Credential)

    $S2DArray = @()
    $Cim      = New-CimSession -ComputerName $($ClusterName + "." + $DomainName) -Credential $Credential

    $StoragePools = Get-StorageSubSystem -CimSession $Cim |? Name -like *$ClusterName*| Get-StoragePool |? isPrimordial -like $False
    Foreach ($StoragePool in $StoragePools){
        $SPFreeSpace = $StoragePool.Size - $SotragePool.AllocatedSize
        $PhysicalDisks = $StoragePool | Get-PhysicalDisk

         if ($NodeNbr -lt 5){
                $ReservedDisk  = $NodeNbr
        }
        Else {
            $ReservedDisk = 4
        }


        If (($PhysicalDisks |? Usage -like "Journal").Count -gt 0){
            # Mode cache
            $NodeNbr = (Get-ClusterNode -Cluster $($Cluster + "." + $DomainName) -Credential $Credential).Count
           
            $PhysicalDisks |? Usage -like "Auto-Select" |% {$CapaSize += $_.Size}
            $PhysicalDisks |? Usage -like "Journal" |% {$CacheSize += $_.Size}

        }
        Else {
            # No Cache
             $PhysicalDisks |% {$CapaSize += $_.Size}
        }
    }


    Return $S2DArray
}

#>


Function Invoke-NodeCommand {
# Runs a script block on a cluster node: locally when the node is this computer, through WinRM otherwise
    Param([string]$NodeName,
          [string]$DomainName,
          [PSCredential]$Credential,
          [ScriptBlock]$ScriptBlock)

    If ($NodeName -like $env:COMPUTERNAME){
        Return (& $ScriptBlock)
    }
    Return Invoke-Command -ComputerName ($NodeName + "." + $DomainName) -Credential $Credential -ScriptBlock $ScriptBlock -ErrorAction Stop
}

Function Test-InfrastructureVM {
# Returns $True when a VM is part of the Azure Local platform (Arc Resource Bridge, ...) and not a customer workload.
# A VM is considered infrastructure when its name matches $InfraVMNamePatterns or when it lives on $InfraVolumeName
    Param([string]$Name,
          [string]$Path)

    Foreach ($Pattern in $InfraVMNamePatterns){
        If ($Name -like $Pattern){ Return $True }
    }
    If ($InfraVolumeName -and $Path -and ($Path -like "*\ClusterStorage\$InfraVolumeName\*")){
        Return $True
    }
    Return $False
}

Function Get-AzureLocalInfo {
# This function collects Azure Local (24H2) platform information: Azure Arc registration, solution version and
# available updates, Network ATC intents and per node OS build / Arc agent state.
# Every section is collected independently: a failure is stored in the Errors property and the report goes on.
    Param([Array]$VMHosts,
          [string]$DomainName,
          [PSCredential]$Credential)

    $AzureLocalObj = New-Object System.Object
    $Errors        = @{}

    # Cluster-wide information is queried from a single node: this computer when it is part of the cluster, otherwise the first node
    $QueryNode = $VMHosts |? Name -like $env:COMPUTERNAME | Select -First 1
    If (-not $QueryNode){ $QueryNode = $VMHosts | Select -First 1 }
    $QueryNode = $QueryNode.Name

    # ---- Azure Arc registration (Get-AzureStackHCI) ----
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 0 -CurrentOperation "Collecting Azure Local registration information"
    $Registration = $Null
    Try {
        $Registration = Invoke-NodeCommand -NodeName $QueryNode -DomainName $DomainName -Credential $Credential -ScriptBlock {
            $R = Get-AzureStackHCI -ErrorAction Stop
            [PSCustomObject]@{
                ClusterStatus      = "$($R.ClusterStatus)"
                RegistrationStatus = "$($R.RegistrationStatus)"
                RegistrationDate   = "$($R.RegistrationDate)"
                AzureResourceName  = "$($R.AzureResourceName)"
                AzureResourceUri   = "$($R.AzureResourceUri)"
                ConnectionStatus   = "$($R.ConnectionStatus)"
                LastConnected      = "$($R.LastConnected)"
                Region             = "$($R.Region)"
                DiagnosticLevel    = "$($R.DiagnosticLevel)"
                IMDSAttestation    = "$($R.IMDSAttestation)"
            }
        }
        # Subscription and resource group are embedded in the Azure resource URI
        $SubscriptionId = ''
        $ResourceGroup  = ''
        If ($Registration.AzureResourceUri -match '(?i)/subscriptions/([^/]+)/resourceGroups/([^/]+)/'){
            $SubscriptionId = $Matches[1]
            $ResourceGroup  = $Matches[2]
        }
        $Registration | Add-Member -Type NoteProperty -Name SubscriptionId -Value $SubscriptionId
        $Registration | Add-Member -Type NoteProperty -Name ResourceGroup  -Value $ResourceGroup
    }
    Catch {
        $Errors['Registration'] = $_.Exception.Message
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name Registration -Value $Registration

    # ---- Solution version (Lifecycle Manager) ----
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 25 -CurrentOperation "Collecting Azure Local solution update information"
    $UpdateEnvironment = $Null
    Try {
        $UpdateEnvironment = Invoke-NodeCommand -NodeName $QueryNode -DomainName $DomainName -Credential $Credential -ScriptBlock {
            $E = Get-SolutionUpdateEnvironment -ErrorAction Stop
            [PSCustomObject]@{
                CurrentVersion  = "$($E.CurrentVersion)"
                State           = "$($E.State)"
                HealthState     = "$($E.HealthState)"
                LastChecked     = "$($E.LastChecked)"
                LastUpdated     = "$($E.LastUpdated)"
                PackageVersions = (($E.PackageVersions | ForEach-Object { "$($_.PackageType) $($_.Version)" }) -join '<br>')
            }
        }
    }
    Catch {
        $Errors['UpdateEnvironment'] = $_.Exception.Message
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name UpdateEnvironment -Value $UpdateEnvironment

    # ---- Available / installed solution updates ----
    $Updates = @()
    Try {
        $Updates = @(Invoke-NodeCommand -NodeName $QueryNode -DomainName $DomainName -Credential $Credential -ScriptBlock {
            Get-SolutionUpdate -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName      = "$($_.DisplayName)"
                    Version          = "$($_.Version)"
                    State            = "$($_.State)"
                    AvailabilityType = "$($_.AvailabilityType)"
                    PackageType      = "$($_.PackageType)"
                    SbeVersion       = "$($_.SbeVersion)"
                    InstalledDate    = "$($_.InstalledDate)"
                }
            }
        })
    }
    Catch {
        $Errors['Updates'] = $_.Exception.Message
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name Updates -Value $Updates

    # ---- Network ATC intents ----
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 50 -CurrentOperation "Collecting Network ATC intents"
    $Intents = @()
    Try {
        $Intents = @(Invoke-NodeCommand -NodeName $QueryNode -DomainName $DomainName -Credential $Credential -ScriptBlock {
            Get-NetIntent -ErrorAction Stop | ForEach-Object {
                $Types = @()
                If ($_.IsManagementIntentSet){ $Types += 'Management' }
                If ($_.IsComputeIntentSet)   { $Types += 'Compute' }
                If ($_.IsStorageIntentSet)   { $Types += 'Storage' }
                If ($_.IsStretchIntentSet)   { $Types += 'Stretch' }
                [PSCustomObject]@{
                    IntentName  = "$($_.IntentName)"
                    Scope       = "$($_.Scope)"
                    IntentType  = ($Types -join ', ')
                    NetAdapters = "$($_.NetAdapterNamesAsList)"
                }
            }
        })
    }
    Catch {
        $Errors['Intents'] = $_.Exception.Message
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name Intents -Value $Intents

    $IntentStatus = @()
    Try {
        $IntentStatus = @(Invoke-NodeCommand -NodeName $QueryNode -DomainName $DomainName -Credential $Credential -ScriptBlock {
            Get-NetIntentStatus -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    Host                = "$($_.Host)"
                    IntentName          = "$($_.IntentName)"
                    ConfigurationStatus = "$($_.ConfigurationStatus)"
                    ProvisioningStatus  = "$($_.ProvisioningStatus)"
                    Error               = "$($_.Error)"
                    LastUpdated         = "$($_.LastUpdated)"
                }
            }
        })
    }
    Catch {
        $Errors['IntentStatus'] = $_.Exception.Message
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name IntentStatus -Value $IntentStatus

    # ---- Per node: OS build and Azure Connected Machine agent ----
    $Nodes = @()
    $i     = 0
    Foreach ($VMHost in $VMHosts){
        Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" `
                       -PercentComplete (75 + ($i/$VMHosts.Count)*25) `
                       -CurrentOperation "Collecting $($VMHost.Name) Azure Local node information"
        $i++
        $NodeObj = New-Object System.Object
        $NodeObj | Add-Member -Type NoteProperty -Name Name -Value $VMHost.Name
        Try {
            $NodeInfo = Invoke-NodeCommand -NodeName $VMHost.Name -DomainName $DomainName -Credential $Credential -ScriptBlock {
                $OS  = Get-CimInstance -ClassName Win32_OperatingSystem
                $Reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
                $Ubr = $Reg.UBR
                $DisplayVersion = $Reg.DisplayVersion

                # Azure Connected Machine agent (Arc)
                $ArcVersion = ''; $ArcStatus = ''; $ArcResource = ''
                $AzCmAgent  = Join-Path $env:ProgramFiles 'AzureConnectedMachineAgent\azcmagent.exe'
                If (Test-Path $AzCmAgent){
                    Try {
                        $Json = & $AzCmAgent show --json 2>$null | Out-String
                        $Arc  = $Json | ConvertFrom-Json
                        $ArcVersion  = "$($Arc.agentVersion)"
                        $ArcStatus   = "$($Arc.status)"
                        $ArcResource = "$($Arc.resourceName)"
                    }
                    Catch { $ArcStatus = 'Unknown (azcmagent show failed)' }
                }
                Else { $ArcStatus = 'Agent not installed' }
                $Himds = Get-Service -Name himds -ErrorAction SilentlyContinue

                [PSCustomObject]@{
                    Caption        = "$($OS.Caption)"
                    Build          = "$($OS.Version).$Ubr"
                    DisplayVersion = "$DisplayVersion"
                    LastBootUpTime = "$($OS.LastBootUpTime)"
                    ArcVersion     = $ArcVersion
                    ArcStatus      = $ArcStatus
                    ArcResource    = $ArcResource
                    HimdsStatus    = $(If ($Himds){ "$($Himds.Status)" } Else { 'Not found' })
                }
            }
            $NodeObj | Add-Member -Type NoteProperty -Name Caption        -Value $NodeInfo.Caption
            $NodeObj | Add-Member -Type NoteProperty -Name Build          -Value $NodeInfo.Build
            $NodeObj | Add-Member -Type NoteProperty -Name DisplayVersion -Value $NodeInfo.DisplayVersion
            $NodeObj | Add-Member -Type NoteProperty -Name LastBootUpTime -Value $NodeInfo.LastBootUpTime
            $NodeObj | Add-Member -Type NoteProperty -Name ArcVersion     -Value $NodeInfo.ArcVersion
            $NodeObj | Add-Member -Type NoteProperty -Name ArcStatus      -Value $NodeInfo.ArcStatus
            $NodeObj | Add-Member -Type NoteProperty -Name ArcResource    -Value $NodeInfo.ArcResource
            $NodeObj | Add-Member -Type NoteProperty -Name HimdsStatus    -Value $NodeInfo.HimdsStatus
            $NodeObj | Add-Member -Type NoteProperty -Name Error          -Value ''
        }
        Catch {
            $NodeObj | Add-Member -Type NoteProperty -Name Error -Value $_.Exception.Message
        }
        $Nodes += $NodeObj
    }
    $AzureLocalObj | Add-Member -Type NoteProperty -Name Nodes  -Value $Nodes
    $AzureLocalObj | Add-Member -Type NoteProperty -Name Errors -Value $Errors

    Return $AzureLocalObj
}

##### Settings #####
# ----------------------------------------------------------------

# Specify the consolidation rate
$TxConso     = 4

###### 0 = Disabled; 1 = Enabled
# Enable Host hardware collecting information
$HostHwInformation   = 1

# Enable Host network collecting information
$HostNetInformation  = 1

# Enable Host storage collecting information
$HostStoInformation  = 1

# Enable cluster storage collecting information
$ClustStoInformation = 1

# Enable Cluster Network collecting information
$ClustNetInformation = 1

# Enable Host OS collecting information
$HostOSInformation   = 1

# Enable Cluster configuration collecting information
$ClustConfInfo       = 1

# Enable Virtual Machine collecting information
$VMHostWorkloadInfo  = 1

# Enable Storage Spaces Direct Collecting information
$ClusterS2D          = 1

# Enable Azure Local platform collecting information (Arc registration, solution updates, Network ATC, Arc agent)
$AzureLocalInfo      = 1

# Azure Local infrastructure VMs are excluded from the workload/consolidation figures and flagged in the VM list.
# A VM is infrastructure when its name matches one of these patterns or when it is stored on the volume below
# (Infrastructure_1 hosts the Arc Resource Bridge and platform VMs in 23H2/24H2).
$InfraVMNamePatterns = @('*arcbridge*', '*-arb-*')
$InfraVolumeName     = 'Infrastructure_1'


##### Gather Information #####
# ----------------------------------------------------------------

Clear
Write-Host "Get information from $($ClusterName + "." + $DomainName)..." -ForegroundColor Green -BackgroundColor Black
Try {
    # Get Cluster object
    $Cluster  = Get-Cluster -Name $($ClusterName + "." + $DomainName) -ErrorAction Stop

    # Get Hyper-V nodes in the cluster
    $VMHosts  = Get-ClusterNode -Cluster $($ClusterName + "." + $DomainName) -ErrorAction Stop |? State -Like "Up" | Select Name

    $ComputerName = Get-Content ENV:COMPUTERNAME
    $TestHost     = $VMHosts |? Name -NotLike $ComputerName | Select -First 1
    New-CimSession -ComputerName $TestHost.Name -Credential $Credential -ErrorAction Stop | Remove-CimSession
    
}
Catch {
    Write-Error "Can't connect to cluster/Node: $($Error[0].Exception.Message) Exiting"
    throw
}

# Create the folder if doesn't exist
Try {
    Resolve-Path -Path $Path -ErrorAction Stop | Out-Null
}
Catch {
    Try {
        New-Item -Path $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    Catch {
        Write-Error "Can't create the folder $($Path): $($Error[0].Exception.Message). Exiting."
        throw
    }
}

$Date      = Get-Date -Format yyyy-MM-dd_HH-mm
$FileName  = "Audit-S2D-$ClusterName-$Date.html"
$ExportLog = $Path + "\" + $FileName
$PdfLog    = [System.IO.Path]::ChangeExtension($ExportLog, '.pdf')

# Report cover metadata
if ([string]::IsNullOrWhiteSpace($Customer))       { $Customer       = Read-RequiredText 'Customer' }
if ([string]::IsNullOrWhiteSpace($Environment))    { $Environment    = Read-EnvironmentOption }
if ([string]::IsNullOrWhiteSpace($ConsultantName)) { $ConsultantName = Read-RequiredText 'Consultant name' }
if ([string]::IsNullOrWhiteSpace($ConsultantEmail)){ $ConsultantEmail= Read-RequiredText 'Consultant email' }
if ([string]::IsNullOrWhiteSpace($InstallDate))    { $InstallDate    = Read-InstallDate }
if ([string]::IsNullOrWhiteSpace($Country))         { $Country        = Read-RequiredText 'Country' }
if ([string]::IsNullOrWhiteSpace($City))            { $City           = Read-RequiredText 'City' }

$CustomerHtml        = ConvertTo-HtmlSafe $Customer
$EnvironmentHtml     = ConvertTo-HtmlSafe $Environment
$ConsultantNameHtml  = ConvertTo-HtmlSafe $ConsultantName
$ConsultantEmailHtml = ConvertTo-HtmlSafe $ConsultantEmail
$InstallDateHtml     = ConvertTo-HtmlSafe $InstallDate
$CountryHtml         = ConvertTo-HtmlSafe $Country
$CityHtml            = ConvertTo-HtmlSafe $City
$ClusterNameHtml     = ConvertTo-HtmlSafe $ClusterName
$ReportDateHtml      = Get-Date -Format 'dd-MMM-yyyy HH:mm'

## If the module is enabled, run each function to collect information

# If enabled, collect Hyper-V hosts hardware information
if ($HostHwInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 0
    $VMHostHwInformation  = Get-VMHostHwInformation -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect Hyper-V Hosts storage information
if ($HostStoInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 10
    $StorageInformation   = Get-VMHostStorage -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect cluster storage information
if ($ClustStoInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 20
    $StoClusterInfo       = Get-ClusterStorage -ClusterName $ClusterName
}

# If enabled, collect Host network configuration
if ($HostNetInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 30
    $VMHostNicsInfo        = Get-VMHostNetwork -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
    $VMHostvSwitchInfo     = Get-VMHostvSwitch -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect Cluster network ifnformation
if ($ClustNetInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 40
    $ClusterNetInformation = Get-ClusterNetInformation -ClusterName $ClusterName -Credential $Credential
}

# If enabled, collect Hyper-V hosts OS information
if ($HostOSInformation){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 50
    $VMHostOSInformation = Get-VMHostOsInformation -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
    $VMHostHyperVInfo    = Get-VMHostHyperVSettings -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect Cluster configuration information
if ($ClustConfInfo){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 60
    $ClusterConfInformation = Get-ClusterConfInformation -ClusterName $ClusterName
}

# If enabled, collect Virtual Machines information
if ($VMHostWorkloadInfo){
     Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 70
    $VMHostsWorkloadInformation = Get-VMHostWorkloads -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect Storage Spaces Direct information
if ($ClusterS2D){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 80
    $StoragePoolInformation  = Get-StoragePoolInfo -ClusterName $ClusterName -DomainName $DomainName -Credential $Credential
    $VirtualDiskInformation  = Get-VirtualDiskInfo -ClusterName $ClusterName -DomainName $DomainName -Credential $Credential
    $PhysicalDiskInformation = Get-PhysicalDiskInfo -ClusterName $ClusterName -DomainName $DomainName -Credential $Credential
    $VMHostSMBMultiChannel   = Get-VMHostSMBMultiChannel -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}

# If enabled, collect Azure Local platform information
if ($AzureLocalInfo){
    Write-Progress -Activity "Collecting Hyper-V/S2D infrastructure information" -PercentComplete 90
    $AzureLocalInformation = Get-AzureLocalInfo -VMHosts $VMHosts -DomainName $DomainName -Credential $Credential
}





##### CSS Content #####
# ----------------------------------------------------------------

# these variables define the name of CSS Class ofr each purpose
$TdError       = "td_Error"
$TDOK          = "td_OK"
$TableClass    = "table_main"
$UnitClass     = "span_Unit"
$ValueClass    = "span_Value"
$AdvertMessage = "span_advert"
$MiscValue     = "span_Misc"
$ComputerClass = "td_computer"
$TDNoInfo      = "td_empty"




##### HTML Content #####
# ----------------------------------------------------------------
$Date        = Get-Date
$HTMLIntitle = @"
<html> 
<Head>
<Title>Audit of $ClusterName Hyper-V/S2D infrastructure - Date: $Date</title>
<Style>
body
{
    font-family: arial, verdana;
    color: #333;
    background-color: #DDD;
    padding: 10px;
    font-size: 14px;
}

h2
{
    font-size: 1.7em;
    color: #800;
}
.table_main
{
    border-collapse: collapse;
    background-color: #FFF;
    font-size: 1em;
    width: 100%;
}

.table_main th
{
    border: 1px solid #AAA;
    background-color: #555;
    padding: 10px;
    color: #FFF;
}

.table_main td
{
    border: 1px solid #AAA;
    text-align: center;
    padding: 10px;
}

.td_computer
{
    background-color: #F5F5F5;
    color: #333;
    font-weight: bold;
}

.td_Error
{
    background-color: #F1302A;
    color: #FFF;
}

.td_OK
{
    background-color: #46B810;
    color: #FFF;
}

.td_empty
{
    
}

.span_Value
{
    font-size: 1.6em;
}

.span_Unit
{
    font-size: 1.2em;
}

.span_advert
{
    font-size: 0.8em;
    font-weight: bold;
    color: #888;
}

/* ===== Corporate cover and print settings ===== */
@page {
    size: A4 landscape;
    margin: 10mm;
}

@media print {
    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }

    body {
        background-color: #DDD !important;
    }

    .report-cover {
        page-break-after: always;
        break-after: page;
    }

    tr {
        page-break-inside: avoid;
        break-inside: avoid;
    }
}

.report-cover {
    box-sizing: border-box;
    width: 100%;
    min-height: 95vh;
    background: #ffffff;
    color: #222222;
    border-top: 18px solid #E2231A;
    border-bottom: 8px solid #E2231A;
    padding: 55px 70px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    font-family: Arial, Verdana, sans-serif;
}

.cover-brand {
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.cover-logo-text {
    background: #E2231A;
    color: #ffffff;
    font-size: 34px;
    font-weight: bold;
    letter-spacing: 1px;
    padding: 16px 28px;
    border-radius: 2px;
}

.cover-classification {
    font-size: 13px;
    color: #777777;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.cover-title-block { margin-top: 65px; }
.cover-title {
    font-size: 42px;
    font-weight: bold;
    color: #222222;
    margin-bottom: 12px;
}

.cover-subtitle {
    font-size: 25px;
    color: #555555;
    margin-bottom: 45px;
}

.cover-grid {
    width: 720px;
    border-collapse: collapse;
    font-size: 17px;
}

.cover-grid td {
    border: 1px solid #cfcfcf;
    padding: 14px 16px;
    text-align: left;
}

.cover-grid td:first-child {
    width: 230px;
    background: #f3f3f3;
    font-weight: bold;
    color: #333333;
}

.cover-footer {
    display: flex;
    justify-content: space-between;
    align-items: end;
    color: #555555;
    font-size: 14px;
}

.cover-footer strong { color: #222222; }
.report-content { margin-top: 0; }

</style>
<meta charset="utf-8"/>
</Head>
<body>

<div class="report-cover">
    <div>
        <div class="cover-brand">
            <div class="cover-logo-text">LENOVO</div>
            <div class="cover-classification">Infrastructure Assessment</div>
        </div>

        <div class="cover-title-block">
            <div class="cover-title">Azure Local / Hyper-V S2D</div>
            <div class="cover-subtitle">Infrastructure Health Audit Report</div>

            <table class="cover-grid">
                <tr><td>Customer</td><td>$CustomerHtml</td></tr>
                <tr><td>Cluster</td><td>$ClusterNameHtml</td></tr>
                <tr><td>Environment</td><td>$EnvironmentHtml</td></tr>
                <tr><td>Installation Date</td><td>$InstallDateHtml</td></tr>
                <tr><td>Country</td><td>$CountryHtml</td></tr>
                <tr><td>City</td><td>$CityHtml</td></tr>
                <tr><td>Report Date</td><td>$ReportDateHtml</td></tr>
                <tr><td>Scope</td><td>Hyper-V, Failover Cluster, Storage Spaces Direct, CSV, Network and Virtual Machines</td></tr>
            </table>
        </div>
    </div>

    <div class="cover-footer">
        <div>
            Generated from S2D Audit HTML Report<br>
            <strong>$ClusterNameHtml</strong>
        </div>
        <div>
            This script was executed by<br>
            <strong>$ConsultantNameHtml</strong><br>
            $ConsultantEmailHtml
        </div>
    </div>
</div>

<div class="report-content">
"@
 
$HTMLEnding = @"
</div>
<br><br><br>
<center><span class="$AdvertMessage">This board has been generated by PowerShell script</span></center>
<center><span class="$AdvertMessage">This script was executed by - $ConsultantNameHtml $ConsultantEmailHtml</span></center>
</body>
</html>
"@

##### Export #####
# ----------------------------------------------------------------

# Show a Progress Bar
Write-Progress -Activity "HTML file construction" -PercentComplete 0

# Export HTML header (HTML, header and opening body)
Set-Content -Path $ExportLog -Value $HTMLIntitle

# Export Hyper-V hosts hardware information
if ($HostHwInformation){
    Write-Progress -Activity "HTML file construction" -PercentComplete 10 -CurrentOperation "Hardware information gathering"
    Add-Content -Path $ExportLog -Value "<H2>Hyper-V Nodes and Cluster information</H2>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Memory Installed</th>'
    Add-Content -Path $ExportLog -Value '<th>CPU Name</th>'
    Add-Content -Path $ExportLog -Value '<th>CPU deviceID</th>'
    Add-Content -Path $ExportLog -Value '<th># Cores</th>'
    Add-Content -Path $ExportLog -Value '<th># Logical Processors</th>'
    Add-Content -Path $ExportLog -Value '<th>Total Logical Processors</th>'
    Add-Content -Path $ExportLog -Value '<th># Virtual Machines</th>'
    Add-Content -Path $ExportLog -Value '<th>vCPU used</th>'
    Add-Content -Path $ExportLog -Value '<th>Memory used</th>'
    Add-Content -Path $ExportLog -Value '<th>Consolidation rate</th>'
    Add-Content -Path $ExportLog -Value '</tr>'
    Foreach ($Node in $VMHostHwInformation){
        $NodeNbr++

        $CPUColsTemp   = @()
        $ThreadNbr     = 0
        $CPUNbr        = 0
        $NodeTotalCore = 0

        # Add each CPU information in a temporary array
        Foreach ($CPU in $Node.CPU){
            # CPUNbr enables to calculate the RowSpan in the HTML table
            $CPUNbr++
            $NodeTotalCore += $CPU.NumberOfCores
            $TotalPhyCore  += $CPU.NumberOfCores
            $TotalCore     += $CPU.NumberOfLogicalProcessors
            $ThreadNbr     += $CPU.NumberOfLogicalProcessors
            $CPUColsTemp   += "<td>$($CPU.Name)</td>"
            $CPUColsTemp   += "<td>$($CPU.DeviceId)</td>"
            $CPUColsTemp   += "<td><span class=$ValueClass>$($CPU.NumberOfCores)</span><span class=$UnitClass>CORES</span></td>"
            $CPUColsTemp   += "<td><span class=$ValueClass>$($CPU.NumberOfLogicalProcessors)</span><span class=$UnitClass>CORES</span></td>"
        }

        # export HTML content to $ExportLog
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value "<td RowSpan=$CPUNbr NOWRAP class=$ComputerClass>$($Node.name)</td>"
        Add-Content -Path $ExportLog -Value "<td RowSpan=$CPUNbr><span class=$ValueClass>$($Node.Memory)</span><span class=$UnitClass>GB</span></td>"    
        Add-Content -Path $ExportLog -Value $CPUColsTemp[0]
        Add-Content -Path $ExportLog -Value $CPUColsTemp[1]
        Add-Content -Path $ExportLog -Value $CPUColsTemp[2]
        Add-Content -Path $ExportLog -Value $CPUColsTemp[3]
        Add-Content -Path $ExportLog -Value "<td RowSpan=$CPUNbr><span class=$ValueClass>$ThreadNbr</span><span class=$UnitClass>THREADS</span></td>"


        # Show the workload of each Hyper-V nodes
        Foreach ($HostWorkload in $Node.Workload){

            # if 85% of the memory is used change the CSS class to erro
            if ($HostWorkload.MemAssigned -gt (85*$Node.Memory)/100){
                $MemClass = $TDError
            }
            Else{
                $MemClass = $TDOK
            }

            # if the consolidation rate is exceeded, change the CSS class to error
            if ($HostWorkload.vCPU -gt ($TxConso*$NodetotalCore)){
                $CPUClass = $TDError
            }
            Else {
                $CPUClass = $TDOK
            }

            $TotalVMs         += $HostWorkload.VMCount
            $TotalInfraVMs    += $HostWorkload.InfraVMCount
            $TotalvCPU        += $HostWorkload.vCPU
            $TotalMemAssigned += $HostWorkload.MemAssigned
            Add-Content -Path $ExportLog -Value "<td RowSpan=$CPUNbr><span class=$ValueClass>$($HostWorkload.VMCount)</span><span class=$UnitClass>VMs</span><br><span class=$AdvertMessage>+ $($HostWorkload.InfraVMCount) infra</span></td>"
            Add-Content -Path $ExportLog -Value "<td Class=$CPUClass RowSpan=$CPUNbr><span class=$ValueClass>$($HostWorkload.vCPU)</span><span class=$UnitClass>vCPU</span></td>"
            Add-Content -Path $ExportLog -Value "<td Class=$MemClass RowSpan=$CPUNbr><span class=$ValueClass>$([Math]::Round($HostWorkload.MemAssigned, 0))</span><span class=$UnitClass>GB</span></td>"

            # The consolidation rate is round to 2 décimal
            $ConsoRate = [Math]::Round($HostWorkload.vCPU / $NodeTotalCore, 2)
            if ($ConsoRate -ge 4){
                $Class = $TDError
            }
            Else {
                $Class = $TDOK
            }
            Add-Content -Path $ExportLog -Value "<td RowSpan=$CPUNbr Class=$Class><span class=$ValueClass>$ConsoRate</td>"
        
        }
        Add-Content -Path $ExportLog -Value '</tr>'
        #Add row outside the rowspan
        For ($i = 3; $i -lt (($CPUnbr*4)-1); $i += 4){
                Add-Content -Path $ExportLog -Value '<tr>'
                Add-Content -Path $ExportLog -Value $CPUColsTemp[$i+1]
                Add-Content -Path $ExportLog -Value $CPUColsTemp[$i+2]
                Add-Content -Path $ExportLog -Value $CPUColsTemp[$i+3]
                Add-Content -Path $ExportLog -Value $CPUColsTemp[$i+4]
                Add-Content -Path $ExportLog -Value '</tr>'
        }
        $TotalMem    += $Node.Memory
        $TotalThread += $ThreadNbr

    }

    # all this calcul is made at N-1 (Remove one node resource from calcul to take care incidents)
    $TotalCore     = [Math]::Round(($TotalCore/$NodeNbr)*($NodeNbr-1), 0)
    $TotalMem      = ($TotalMem/$NodeNbr)*($NodeNbr-1)
    $TotalThread   = [Math]::Round(($TotalThread/$NodeNbr)*($NodeNbr-1), 0)
    $vCPUAvailable = [Math]::Round(($TxConso*($TotalPhyCore/$NodeNbr))*($NodeNbr-1), 0)
    $ClusterTxRate = [Math]::Round($TotalvCPU / $TotalPhyCore, 2)

    # export the row for cluster information
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value "<td class=$ComputerClass NOWRAP>$ClusterName<br><span class=$AdvertMessage>(N-1 to take care one host down)</span></td>"
    Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$TotalMem</span><span class=$UnitClass>GB</span></td>"
    Add-Content -Path $ExportLog -Value "<td colspan=4><span class=$ValueClass>Consolidation rate $($TxConso):1</span></td>"
    Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$TotalCore</span><span class=$UnitClass>THREADS</span><br><span class=$ValueClass>$vCPUAvailable</span><span class=$UnitClass>vCPU</span></td>"
    Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$TotalVMs</span><span class=$UnitClass>VMs</span><br><span class=$AdvertMessage>+ $TotalInfraVMs infra (excluded)</span></td>"

    # If vCPU allocated is greater than vCPU available, change CSS class to error
    if ($TotalvCPU -ge $vCPUAvailable){
                $Class = $TDError
    }
    Else {
                $Class = $TDOK
    }
    Add-Content -Path $ExportLog -Value "<td Class=$Class><span class=$ValueClass>$TotalvCPU</span><span class=$UnitClass>vCPU</span></td>"

    # if the total memory assigned is greater than memory available,change CSS class to error
    if ($TotalMemAssigned -ge $TotalMem){
                $Class = $TDError
    }
    Else {
                $Class = $TDOK
    }

    Add-Content -Path $ExportLog -Value "<td class=$Class><span class=$ValueClass>$([Math]::Round($TotalMemAssigned, 0))</span><span class=$UnitClass>GB</span></td>"

    # if the cluster rate is greater than expected, change CSS class to error
    if ($ClusterTxRate -ge $TxConso){
                $Class = $TDError
    }
    Else {
                $Class = $TDOK
    }

    Add-Content -Path $ExportLog -Value "<td class=$Class><span class=$ValueClass>$ClusterTxRate</span></td>"
    Add-Content -Path $ExportLog -Value '</table>'
}

# Export Hyper-V hosts network information to HTML
if ($HostNetInformation) {
    Add-Content -Path $ExportLog -Value "<H2>Network information</H2>"
    Foreach ($VMHost in $VMHosts){
        $VMHostNics = $VMHostNicsInfo |? VMHost -like $VMHost.Name

        Add-Content -Path $ExportLog -Value "<H3>Network information on $($VMHost.Name)</H3>"
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>NIC Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Description</th>'
        Add-Content -Path $ExportLog -Value '<th>VLAN</th>'
        Add-Content -Path $ExportLog -Value '<th>TCP/IP</th>'
        Add-Content -Path $ExportLog -Value '<th>Link Speed</th>'
        Add-Content -Path $ExportLog -Value '<th>RSS State</th>'
        Add-Content -Path $ExportLog -Value '<th>RSS</th>'
        Add-Content -Path $ExportLog -Value '<th>VMQ State</th>'
        Add-Content -Path $ExportLog -Value '<th>VMQ</th>'
        Add-Content -Path $ExportLog -Value '<th>RDMA State</th>'
        Add-Content -Path $ExportLog -Value '<th>VMMQ State</th>'
        Add-Content -Path $ExportLog -Value '<th>QoS</th>'
        Add-Content -Path $ExportLog -Value '<th>MTU</th>'
        Add-Content -Path $ExportLog -Value '<th>Switch Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Team Mapping</th>'
        Add-Content -Path $ExportLog -Value '</tr>'

        Foreach ($Nic in $VMHostNics){
            Add-Content -Path $ExportLog -Value "<tr>"
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($Nic.Name)<br><span class=$AdvertMessage>$($Nic.Type)</span></td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.Description)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.Vlan)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($Nic.IPAddress)<br><span class=$AdvertMessage>$($NIC.Gateway)</span><br><span class=$AdvertMessage>$($NIC.DNS)</span><br><span class=$AdvertMessage>$($NIC.DNSRegistration)</span></td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.LinkSpeed)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.RSSState)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.RSS)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.VMQState)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.VMQ)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.RDMAState)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.VMMQ)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.QoS)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($Nic.MTU)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.SwitchName)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Nic.TeamMapping)</td>"
            Add-Content -Path $ExportLog -Value "</tr>"
        }
        Add-Content -Path $ExportLog -Value '</table>'

        $VMHostvSwitch = $VMHostvSwitchInfo |? VMHost -like $VMHost.Name

        Add-Content -Path $ExportLog -Value "<H3>vSwitches on $($VMHost.Name)</H3>"
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Switch Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Type</th>'
        Add-Content -Path $ExportLog -Value '<th>Embedded Teaming</th>'
        Add-Content -Path $ExportLog -Value '<th>Packet Direct</th>'
        Add-Content -Path $ExportLog -Value '<th>SRIOV</th>'
        Add-Content -Path $ExportLog -Value '<th>QoS Mode</th>'
        Add-Content -Path $ExportLog -Value '<th>NICs</th>'
        Add-Content -Path $ExportLog -Value '</tr>'

        Foreach ($vSwitch in $VMHostvSwitch){
            $NICsNbr = 0
            $NetColsTemp = @()
            Foreach ($NIC in $vSwitch.NICs){
                $NICsNbr++
                $NetColsTemp += "<td class=$ComputerClass>$NIC</td>"
            }

            Add-Content -Path $ExportLog -Value "<tr>"
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass RowSpan=$NICsNbr NOWRAP>$($vSwitch.Name)</td>" 
            Add-Content -Path $ExportLog -Value "<td RowSpan=$NICsNbr NOWRAP>$($vSwitch.Type)</td>" 
            Add-Content -Path $ExportLog -Value "<td RowSpan=$NICsNbr NOWRAP>$($vSwitch.EmbeddedTeaming)</td>"  
            Add-Content -Path $ExportLog -Value "<td RowSpan=$NICsNbr NOWRAP>$($vSwitch.PacketDirect)</td>"  
            Add-Content -Path $ExportLog -Value "<td RowSpan=$NICsNbr NOWRAP>$($vSwitch.IOVSupport)</td>"  
            Add-Content -Path $ExportLog -Value "<td RowSpan=$NICsNbr NOWRAP>$($vSwitch.QoSMode)</td>"
            
            Add-Content -Path $ExportLog -Value $NetColsTemp[0]  
            Add-Content -Path $ExportLog -Value "</tr>"
            For ($i = 0; $i -lt ($NICsNbr * 1)-1; $i++){
                Add-Content -Path $ExportLog -Value "<tr>"
                Add-Content -Path $ExportLog -Value $NetColsTemp[$i+1]
                Add-Content -Path $ExportLog -Value "</tr>" 
            }     
        }
        Add-Content -Path $ExportLog -Value '</table>'
    }


}

# Export Hyper-V hosts local storage information to HTML
if ($HostStoInformation){
    # Show progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 20 -CurrentOperation "storage information gathering"

    #export Table header and title
    Add-Content -Path $ExportLog -Value "<H2>Storage information</H2>"
    Add-Content -Path $ExportLog -Value "<H3>Host storage information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Drive Letter</th>'
    Add-Content -Path $ExportLog -Value '<th>File System Label</th>'
    Add-Content -Path $ExportLog -Value '<th>File System</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Size Remaining</th>'
    Add-Content -Path $ExportLog -Value '<th>Percentage free space</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # for each Hyper-V nodes, export information
    Foreach ($Node in $StorageInformation){

        $StorageColsTemp = @()
        $StoNbr          = 0

        # for each storage device, export information in a temporary array
        Foreach ($Storage in $Node.StorageInformation){

            # This value enables to know the rowspan for the html table
            $StoNbr++

            # if it's a drive letter, :\ characters are added
            if ($Storage.DriveLetter -match "[A-Z]"){
                $StorageColsTemp += "<td>$($Storage.DriveLetter):\</td>"
            }
            else{
                $StorageColsTemp += "<td class=$TDNoInfo></td>"
            }
            $StorageColsTemp += "<td>$($Storage.FSLabel)</td>"
            $StorageColsTemp += "<td>$($Storage.FileSystem)</td>"
        

            # the value is converted to GB and rounded to 1 decimal
            $StorageColsTemp += "<td><span class=$ValueClass>$([Math]::Round($Storage.Size/1GB, 1))</span><span class=$UnitClass>GB</span></td>"
            $StorageColsTemp += "<td><span class=$ValueClass>$([Math]::Round($Storage.SizeRemaining/1GB, 1))</span><span class=$UnitClass>GB</span></td>"

            # Calculate the free percentage rounded to 2
            $PercentFreeSpace = [Math]::Round(($($Storage.SizeRemaining)*100)/$($Storage.Size), 2)

            # If there is less or equal 15% free space, CSS class is change to error
            if ($PercentFreeSpace -le 15){
                $Class = $TDError
            }
            Else{
                $Class = $TDOK
            }
            $StorageColsTemp += "<td class=$Class><span class=$ValueClass>$PercentFreeSpace</span><span class=$UnitClass>%</span></td>"
        }
        # Export the first line with the HTML rowspan
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass rowspan=$StoNbr NOWRAP>$($Node.Name)</td>"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[0])"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[1])"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[2])"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[3])"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[4])"
        Add-Content -Path $ExportLog -Value "$($StorageColsTemp[5])"
        Add-Content -Path $ExportLog -Value '</tr>'

        # Export other line outside the HTML rowspan
        For ($i = 5; $i -lt (($StoNbr*6)-1); $i += 6){
            Add-Content -Path $ExportLog -Value '<tr>'
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+1])"
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+2])"
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+3])"
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+4])"
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+5])"
            Add-Content -Path $ExportLog -Value "$($StorageColsTemp[$i+6])"
            Add-Content -Path $ExportLog -Value '</tr>'
        }
    }
    Add-Content -Path $ExportLog -Value '</table>'
}

# Export Cluster storage information
if ($ClustStoInformation){
    # export the title and the header of the table
    Add-Content -Path $ExportLog -Value "<H3>Cluster storage information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Cluster Name</th>'
    Add-Content -Path $ExportLog -Value '<th>CSV Name</th>'
    Add-Content -Path $ExportLog -Value '<th>State</th>'
    Add-Content -Path $ExportLog -Value '<th>Maintenance Mode</th>'
    Add-Content -Path $ExportLog -Value '<th>Volume Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Size Remaining</th>'
    Add-Content -Path $ExportLog -Value '<th>Percentage free space</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # For each cluster, exporting HTML information
    Foreach ($HvCluster in $StoClusterInfo){
        $ClustStoColsTemp = @()
        $CSVNbr           = 0

        # For each storage device in the cluster, exporting HTML in a temporary array
        Foreach ($Storage in $HvCluster.StorageInformation){

            # This variable enables to calculate the HTML rowspan
            $CSVNbr++
            $ClustStoColsTemp += "<td>$($Storage.Name)</td>"

            # If the storage device is not online, change the CSS class to error
            If ($Storage.State -notlike "Online"){
                $Class = $TDError
            }
            Else {
                $Class = $TDOK
            }
            $ClustStoColsTemp += "<td class=$Class><span class=$ValueClass>$($Storage.State)</span></td>"

            # if the storage device is in maintenance mode, change the CSS class to error
            If ($Storage.MaintenanceMode -like "True"){
                $Class           = $TDWarn
                $MaintenanceMode = "Enabled"
            }
            Else{
                $Class           = $TDOK
                $MaintenanceMode = "Disabled"
            }

            $ClustStoColsTemp += "<td class=$Class><span class=$ValueClass>$MaintenanceMode</span></td>"
            $ClustStoColsTemp += "<td>$($Storage.FriendlyVolumeName)</td>"
            $ClustStoColsTemp += "<td><span class=$ValueClass>$([math]::round($Storage.Size/1GB, 1))</span><span class=$UnitClass>GB</span></td>"

            # Calculate the free space on storage device
            $SizeRemaining     = $Storage.Size - $Storage.UsedSpace
            $ClustStoColsTemp += "<td><span class=$ValueClass>$([math]::round($SizeRemaining/1GB, 1))</span><span class=$UnitClass>GB</span></td>"
            $PercentFreeSpace  = [math]::round(($SizeRemaining*100)/$Storage.Size, 2)
        
            # if the storage device size is less or equal to 1GB and if there is only 15% or less of free space, change CSS class to error
            if (($Storage.Size/1GB -le 1) -and ($PercentFreeSpace -le 15)){
                $Class = $TDError
            }

            # if the storage device size is greater than 1GB and if there is only 10% or less of free space, change CSS class to error
            Elseif (($Storage.Size/1GB -gt 1) -and ($PercentFreeSpace -le 10)){
                $Class = $TDError
            }
            Else{
                $Class = $TDOK
            }
            $ClustStoColsTemp += "<td class=$Class><span class=$ValueClass>$PercentFreeSpace</span><span class=$UnitClass>%</span></td>"
        }
    }

    # Export the first line of the table with rowspan
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value "<td class=$ComputerClass rowspan=$CSVNbr NOWRAP>$($HvCluster.Name)</td>"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[0])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[1])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[2])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[3])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[4])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[5])"
    Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[6])"
    Add-Content -Path $ExportLog -Value '</tr>' 

    #export lines outside the rowspan
    For ($i = 6; $i -lt (($CSVNbr*7)-1); $i += 7){
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+1])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+2])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+3])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+4])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+5])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+6])"
        Add-Content -Path $ExportLog -Value "$($ClustStoColsTemp[$i+7])"
        Add-Content -Path $ExportLog -Value '</tr>'
    }
    Add-Content -Path $ExportLog -Value '</table>'
}

# Export Cluster Network Information to HTML
if ($ClustNetInformation){
    # Show a progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 40 -CurrentOperation "Cluster Network information gathering"

    # Export title and table header
    Add-Content -Path $ExportLog -Value "<H3>Cluster network Information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Cluster Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Network Name</th>'
    Add-Content -Path $ExportLog -Value '<th>IP Address</th>'
    Add-Content -Path $ExportLog -Value '<th>Address Mask</th>'
    Add-Content -Path $ExportLog -Value '<th>Role</th>'
    Add-Content -Path $ExportLog -Value '<th>State</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    $ClusterNetColsTemp = @()
    $NetClustNbr      = 0

    # For each network in the cluster, exporting information
    Foreach ($Network in  $ClusterNetInformation){
        # This variable enables to calculate the rowspan of the HTML table
        $NetClustNbr++

        $LMClass =""
        # If the Network is enabled to transmit Live-Migration, verifying if cluster communication are enabled
        if ($Network.LMNet){
            $ClusterNetColsTemp += "<td>$($Network.Name)<br><span class=$AdvertMessage>Live-Migration Network</span></td>"
        }
        Else {
            $ClusterNetColsTemp += "<td>$($Network.Name)</td>"
        }
        $ClusterNetColsTemp += "<td>$($Network.Address)</td>"
        $ClusterNetColsTemp += "<td>$($Network.AddressMask)</td>"
        $ClusterNetColsTemp += "<td class=$LMClass>$($Network.Role)</td>"

        # if the state of the network is UP, export CSS class OK
        if ($Network.State -like "Up"){
            $Class = $TDOK
        }
        Else{
            $Class = $TDError
        }
        $ClusterNetColsTemp += "<td class=$class><span class=$ValueClass>$($Network.State)</span></td>"
    }

    # Exporting the first line of the table with the rowspan
    Add-Content -Path $ExportLog -Value "<tr>"
    Add-Content -Path $ExportLog -Value "<td class=$ComputerClass RowSpan=$NetClustNbr>$ClusterName</td>"
    Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[0])"
    Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[1])"
    Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[2])"
    Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[3])"
    Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[4])"
    Add-Content -Path $ExportLog -Value "</tr>"

    # exporting other line outside the rowspan
    For ($i = 4; $i -lt ($NetClustNbr*5)-1; $i += 5){
        Add-Content -Path $ExportLog -Value "<tr>"
        Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[$i+1])"
        Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[$i+2])"
        Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[$i+3])"
        Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[$i+4])"
        Add-Content -Path $ExportLog -Value "$($ClusterNetColsTemp[$i+5])"
        Add-Content -Path $ExportLog -Value "</tr>"
    }
    Add-Content -Path $ExportLog -Value "</table>"
}

# Exporting Hyper-V Host OS Information to HTML
if ($HostOSInformation){
    # Show a progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 50 -CurrentOperation "Host OS information gathering"

    #export to HTML file the title and the table header
    Add-Content -Path $ExportLog -Value "<H2>Host OS Information</H2>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>OS Version</th>'
    Add-Content -Path $ExportLog -Value '<th>OS Language</th>'
    Add-Content -Path $ExportLog -Value '<th>Firewall State</th>'
    Add-Content -Path $ExportLog -Value '<th>HotFix installed</th>'
    Add-Content -Path $ExportLog -Value '<th>Minimal Interface</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    $i           = 0
    $HotFixError = $False

    # for each node in the cluster
    Foreach ($Node in $VMHostOSInformation) {
        # on the first node, add to variable the number of the hotfix
        if ($i -eq 0){
            $i++
            $HotFixNbr = $Node.OSHotfix
        }
        # if it is not the first node ...
        Else {
            # verifying if the number of Hotfix is the same. If not, break the loop and change $hotfixerror variable to $true
            if ($Node.OSHotFix -ne $HotFixNbr){
                $HotFixClass   = $TDError
                  $HotFixError = $True
                  break
            }
        }
            
    }

    #if there is no error with hotfix, CSS class is OK
    if(!$HotFixError){
        $HotFixClass = $TDOK
    }

    #for each node in the cluster, export HTML information
    Foreach ($Node in $VMHostOSInformation){

        $FirewallColsTemp = @()

        # For each Firewall profile, export HTML to a temporary array. Change CSS class related to firewall profile state
        Foreach ($Firewall in $Node.Firewall){

            
            if ($Firewall.State -eq 1){
                $Class = $TDOK
                $FirewallState = "Enabled"
            }
            Else{
                $Class = $TDError
                $FirewallState = "Disabled"
            }
            $FirewallColsTemp += "<td class=$class>$($Firewall.Name): $FirewallState</td>"
        }

        Add-Content -path $ExportLog -Value "<tr>"
        Add-Content -path $ExportLog -Value "<td RowSpan=3 class=$ComputerClass>$($Node.Name)</td>"
        Add-Content -path $ExportLog -Value "<td RowSpan=3>$($Node.OSVersion)</td>"

        # Get the textual OS language by using Get-OSLanguage function
        $OSlanguage = Get-OSLanguage -Language $Node.OSLanguage

        # if the OS language is not En-US, change the CSS Class to Error
        if ($Node.OSLanguage -ne 1033){
            $Class = $TDError
        }
        Else{
            $Class = $TDOK
        }
        Add-Content -path $ExportLog -Value "<td RowSpan=3 class=$Class><span class=$ValueClass>$OSLanguage</span></td>"
        Add-Content -path $ExportLog -Value $FirewallColsTemp[0]
        Add-Content -path $ExportLog -Value "<td RowSpan=3 class=$HotFixClass><span class=$ValueClass>$($Node.OSHotfix)</span></td>"

        # if The node is not in minimal interface, change the CSS class to Error
        If ($Node.GuiInstalled){
            $MinShell = "No"
            $Class    = $TDError
        }
        Else{
            $MinShell = "Yes"
            $Class    = $TDOK
        }

        Add-Content -path $ExportLog -Value "<td RowSpan=3 class=$Class><span class=$ValueClass>$MinShell</span></td>"
        Add-Content -path $ExportLog -Value "</tr>"

        For ($i = 1; $i -le 2; $i++){
           Add-Content -path $ExportLog -Value "<tr>"
           Add-Content -path $ExportLog -Value $FirewallColsTemp[$i]
           Add-Content -path $ExportLog -Value "</tr>"
        }
        
    }
    # Table about Hyper-V settings
    Add-Content -Path $ExportLog -Value "</table>"
    Add-Content -Path $ExportLog -Value "<H2>Hyper-V Information</H2>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>VM Path</th>'
    Add-Content -Path $ExportLog -Value '<th>VHD Path</th>'
    Add-Content -Path $ExportLog -Value '<th>Simultaneous Live-Migration</th>'
    Add-Content -Path $ExportLog -Value '<th>Simultaneous Storage Live-Migration</th>'
    Add-Content -Path $ExportLog -Value '<th>Live-Migration Authentication</th>'
    Add-Content -Path $ExportLog -Value '<th>Live-Migration Performance option</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # export Hyper-V Settings to the table
    Foreach ($VMHost in $VMHostHyperVInfo){
        Add-Content -Path $ExportLog -Value "<tr>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($VMHost.VMHost)</td>" 
        Add-Content -Path $ExportLog -Value "<td>$($VMHost.VMPath)</td>" 
        Add-Content -Path $ExportLog -Value "<td>$($VMHost.VHDPath)</td>" 
        Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($VMHost.MaximumLM)</td>" 
        Add-Content -Path $ExportLog -Value "<td class=$ValueClass>$($VMHost.MaximumStoMig)</td>" 
        Add-Content -Path $ExportLog -Value "<td>$($VMHost.LMAuthentication)</td>" 
        Add-Content -Path $ExportLog -Value "<td>$($VMHost.LMPerformanceOption)</td>"
        Add-Content -Path $ExportLog -Value '</tr>'
    }
    Add-Content -Path $ExportLog -Value "</table>"
}

# Exporting cluster configuration information to HTML
if ($ClustConfInfo){

    # Show a progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 60 -CurrentOperation "Cluster information gathering"
    
    # Export title and table header
    Add-Content -Path $ExportLog -Value "<H2>Cluster Information</H2>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Node Number</th>'
    Add-Content -Path $ExportLog -Value '<th>Quorum Type</th>'
    Add-Content -Path $ExportLog -Value '<th>Quorum Resource</th>'
    Add-Content -Path $ExportLog -Value '<th>Dynamic Witness</th>'
    Add-Content -Path $ExportLog -Value '<th>Dynamic Quorum</th>'
    Add-Content -Path $ExportLog -Value '<th>Block Cache Size</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # Export in HTML cluster information
    Add-Content -Path $ExportLog -Value "<tr>"
    Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($ClusterConfInformation.Name)</td>"
    Add-Content -Path $ExportLog -Value "<td>$($ClusterConfInformation.NodeNbr)</td>"
    Add-Content -Path $ExportLog -Value "<td>$($ClusterConfInformation.QuorumType)</td>"
    Add-Content -Path $ExportLog -Value "<td>$($ClusterConfInformation.QuorumResource)</td>"
    Add-Content -Path $ExportLog -Value "<td>$($ClusterConfInformation.WitnessDynamicWeight)</td>"
    Add-Content -Path $ExportLog -Value "<td>$($ClusterConfInformation.DynamicQuorum)</td>"

    # If BLockCacheSize is less than 512MB, export CSS class Error
    if ($ClusterConfInformation.BlockCacheSize -lt 512){
        $Class = $TDError
    }
    Else{
        $Class = $TDOK
    }
    Add-Content -Path $ExportLog -Value "<td class=$Class><span class=$ValueClass>$($ClusterConfInformation.BlockCacheSize)</span><span class=$UnitClass>MB</span></td>"
    Add-Content -Path $ExportLog -Value "</tr>"

    Add-Content -Path $ExportLog -Value "</table>"

}

# Export Azure Local platform information to HTML
if ($AzureLocalInfo){
    # Show a progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 65 -CurrentOperation "Azure Local information gathering"

    Add-Content -Path $ExportLog -Value "<H2>Azure Local Platform</H2>"

    # ---- Azure Arc registration ----
    Add-Content -Path $ExportLog -Value "<H3>Azure Arc registration</H3>"
    $Reg = $AzureLocalInformation.Registration
    if ($Reg){
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Cluster Status</th>'
        Add-Content -Path $ExportLog -Value '<th>Registration Status</th>'
        Add-Content -Path $ExportLog -Value '<th>Connection Status</th>'
        Add-Content -Path $ExportLog -Value '<th>Last Connected</th>'
        Add-Content -Path $ExportLog -Value '<th>Azure Resource</th>'
        Add-Content -Path $ExportLog -Value '<th>Subscription</th>'
        Add-Content -Path $ExportLog -Value '<th>Resource Group</th>'
        Add-Content -Path $ExportLog -Value '<th>Region</th>'
        Add-Content -Path $ExportLog -Value '<th>Registration Date</th>'
        Add-Content -Path $ExportLog -Value '<th>Diagnostic Level</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
        Add-Content -Path $ExportLog -Value '<tr>'

        # Cluster must be registered and connected to Azure, otherwise export CSS class error
        if ($Reg.ClusterStatus -like "Clustered")     { $Class = $TDOK } Else { $Class = $TDError }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($Reg.ClusterStatus)</td>"
        if ($Reg.RegistrationStatus -like "Registered"){ $Class = $TDOK } Else { $Class = $TDError }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($Reg.RegistrationStatus)</td>"
        if ($Reg.ConnectionStatus -like "Connected")   { $Class = $TDOK } Else { $Class = $TDError }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($Reg.ConnectionStatus)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Reg.LastConnected)</td>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$(ConvertTo-HtmlSafe $Reg.AzureResourceName)</td>"
        Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Reg.SubscriptionId)</td>"
        Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Reg.ResourceGroup)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Reg.Region)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Reg.RegistrationDate)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Reg.DiagnosticLevel)</td>"
        Add-Content -Path $ExportLog -Value '</tr>'
        Add-Content -Path $ExportLog -Value '</table>'
    }
    Else {
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $AzureLocalInformation.Errors['Registration'])</span></p>"
    }

    # ---- Solution version and updates (Lifecycle Manager) ----
    Add-Content -Path $ExportLog -Value "<H3>Solution version (Lifecycle Manager)</H3>"
    $Env = $AzureLocalInformation.UpdateEnvironment
    if ($Env){
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Current Solution Version</th>'
        Add-Content -Path $ExportLog -Value '<th>State</th>'
        Add-Content -Path $ExportLog -Value '<th>Health State</th>'
        Add-Content -Path $ExportLog -Value '<th>Package Versions</th>'
        Add-Content -Path $ExportLog -Value '<th>Last Checked</th>'
        Add-Content -Path $ExportLog -Value '<th>Last Updated</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass><span class=$ValueClass>$($Env.CurrentVersion)</span></td>"

        # An update environment which is not AppliedSuccessfully (e.g. UpdateAvailable, UpdateFailed) is highlighted
        if ($Env.State -like "AppliedSuccessfully"){ $Class = $TDOK } Else { $Class = $TDError }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($Env.State)</td>"
        if ($Env.HealthState -like "Success")      { $Class = $TDOK } Else { $Class = $TDError }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($Env.HealthState)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Env.PackageVersions)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Env.LastChecked)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($Env.LastUpdated)</td>"
        Add-Content -Path $ExportLog -Value '</tr>'
        Add-Content -Path $ExportLog -Value '</table>'
    }
    Else {
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $AzureLocalInformation.Errors['UpdateEnvironment'])</span></p>"
    }

    Add-Content -Path $ExportLog -Value "<H3>Solution updates</H3>"
    if ($AzureLocalInformation.Errors.ContainsKey('Updates')){
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $AzureLocalInformation.Errors['Updates'])</span></p>"
    }
    Elseif (($AzureLocalInformation.Updates | Measure-Object).Count -eq 0){
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>No solution update discovered.</span></p>"
    }
    Else {
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Update</th>'
        Add-Content -Path $ExportLog -Value '<th>Version</th>'
        Add-Content -Path $ExportLog -Value '<th>State</th>'
        Add-Content -Path $ExportLog -Value '<th>Package Type</th>'
        Add-Content -Path $ExportLog -Value '<th>Availability</th>'
        Add-Content -Path $ExportLog -Value '<th>SBE Version</th>'
        Add-Content -Path $ExportLog -Value '<th>Installed Date</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
        Foreach ($Update in $AzureLocalInformation.Updates){
            # Installed updates are OK; anything pending (Ready, Preparing, ...) or failed is highlighted
            Switch -Wildcard ($Update.State){
                "Installed"      { $Class = $TDOK }
                "*Fail*"         { $Class = $TDError }
                "*Ready*"        { $Class = $TDError }
                default          { $Class = $TDNoInfo }
            }
            Add-Content -Path $ExportLog -Value '<tr>'
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$(ConvertTo-HtmlSafe $Update.DisplayName)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Update.Version)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$Class>$($Update.State)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Update.PackageType)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Update.AvailabilityType)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Update.SbeVersion)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Update.InstalledDate)</td>"
            Add-Content -Path $ExportLog -Value '</tr>'
        }
        Add-Content -Path $ExportLog -Value '</table>'
    }

    # ---- Network ATC intents ----
    Add-Content -Path $ExportLog -Value "<H3>Network ATC intents</H3>"
    if ($AzureLocalInformation.Errors.ContainsKey('Intents')){
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $AzureLocalInformation.Errors['Intents'])</span></p>"
    }
    Elseif (($AzureLocalInformation.Intents | Measure-Object).Count -eq 0){
        # Network ATC is mandatory on Azure Local 23H2/24H2, so a cluster without intent is an error
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>No Network ATC intent found on this cluster.</span></p>"
    }
    Else {
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Intent Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Scope</th>'
        Add-Content -Path $ExportLog -Value '<th>Intent Type</th>'
        Add-Content -Path $ExportLog -Value '<th>Network Adapters</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
        Foreach ($Intent in $AzureLocalInformation.Intents){
            Add-Content -Path $ExportLog -Value '<tr>'
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$(ConvertTo-HtmlSafe $Intent.IntentName)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Intent.Scope)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Intent.IntentType)</td>"
            Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Intent.NetAdapters)</td>"
            Add-Content -Path $ExportLog -Value '</tr>'
        }
        Add-Content -Path $ExportLog -Value '</table>'
    }

    Add-Content -Path $ExportLog -Value "<H3>Network ATC intent status per node</H3>"
    if ($AzureLocalInformation.Errors.ContainsKey('IntentStatus')){
        Add-Content -Path $ExportLog -Value "<p><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $AzureLocalInformation.Errors['IntentStatus'])</span></p>"
    }
    Elseif (($AzureLocalInformation.IntentStatus | Measure-Object).Count -gt 0){
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Intent Name</th>'
        Add-Content -Path $ExportLog -Value '<th>Configuration Status</th>'
        Add-Content -Path $ExportLog -Value '<th>Provisioning Status</th>'
        Add-Content -Path $ExportLog -Value '<th>Error</th>'
        Add-Content -Path $ExportLog -Value '<th>Last Updated</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
        Foreach ($Status in ($AzureLocalInformation.IntentStatus | Sort-Object Host, IntentName)){
            if ($Status.ConfigurationStatus -like "Success"){ $Class = $TDOK } Else { $Class = $TDError }
            Add-Content -Path $ExportLog -Value '<tr>'
            Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($Status.Host)</td>"
            Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Status.IntentName)</td>"
            Add-Content -Path $ExportLog -Value "<td class=$Class>$($Status.ConfigurationStatus)</td>"
            if ($Status.ProvisioningStatus -like "Completed"){ $Class = $TDOK } Else { $Class = $TDError }
            Add-Content -Path $ExportLog -Value "<td class=$Class>$($Status.ProvisioningStatus)</td>"
            Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Status.Error)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Status.LastUpdated)</td>"
            Add-Content -Path $ExportLog -Value '</tr>'
        }
        Add-Content -Path $ExportLog -Value '</table>'
    }

    # ---- Per node OS build and Arc agent ----
    Add-Content -Path $ExportLog -Value "<H3>Node OS build and Azure Connected Machine agent</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Node Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Operating System</th>'
    Add-Content -Path $ExportLog -Value '<th>Version</th>'
    Add-Content -Path $ExportLog -Value '<th>OS Build</th>'
    Add-Content -Path $ExportLog -Value '<th>Last Boot</th>'
    Add-Content -Path $ExportLog -Value '<th>Arc Agent Version</th>'
    Add-Content -Path $ExportLog -Value '<th>Arc Agent Status</th>'
    Add-Content -Path $ExportLog -Value '<th>Arc Resource Name</th>'
    Add-Content -Path $ExportLog -Value '<th>himds Service</th>'
    Add-Content -Path $ExportLog -Value '</tr>'
    Foreach ($Node in $AzureLocalInformation.Nodes){
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($Node.Name)</td>"
        if ($Node.Error){
            Add-Content -Path $ExportLog -Value "<td colspan=8><span class=$AdvertMessage>Not available: $(ConvertTo-HtmlSafe $Node.Error)</span></td>"
        }
        Else {
            Add-Content -Path $ExportLog -Value "<td>$($Node.Caption)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Node.DisplayVersion)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Node.Build)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Node.LastBootUpTime)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($Node.ArcVersion)</td>"
            if ($Node.ArcStatus -like "Connected"){ $Class = $TDOK } Else { $Class = $TDError }
            Add-Content -Path $ExportLog -Value "<td class=$Class>$(ConvertTo-HtmlSafe $Node.ArcStatus)</td>"
            Add-Content -Path $ExportLog -Value "<td>$(ConvertTo-HtmlSafe $Node.ArcResource)</td>"
            if ($Node.HimdsStatus -like "Running"){ $Class = $TDOK } Else { $Class = $TDError }
            Add-Content -Path $ExportLog -Value "<td class=$Class>$($Node.HimdsStatus)</td>"
        }
        Add-Content -Path $ExportLog -Value '</tr>'
    }
    Add-Content -Path $ExportLog -Value '</table>'
}

# Exporting Virtual Machines information to HTML
if ($VMHostWorkloadInfo){
    # Show a progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 70 -CurrentOperation "VM Hosts workload gathering"

    # Export title and table header
    Add-Content -Path $ExportLog -Value "<H2>Virtual Machines</H2>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>VM Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Hyper-V Host</th>'
    Add-Content -Path $ExportLog -Value '<th>Role</th>'
    Add-Content -Path $ExportLog -Value '<th>State</th>'
    Add-Content -Path $ExportLog -Value '<th>Gen</th>'
    Add-Content -Path $ExportLog -Value '<th>Clustered</th>'
    Add-Content -Path $ExportLog -Value '<th>vCPU count</th>'
    Add-Content -Path $ExportLog -Value '<th>Dynamic Memory</th>'
    Add-Content -Path $ExportLog -Value '<th>Memory Demand</th>'
    Add-Content -Path $ExportLog -Value '<th>Memory Assigned</th>'
    Add-Content -Path $ExportLog -Value '<th>Disk Controller</th>'
    Add-Content -Path $ExportLog -Value '<th>Disk Path</th>'
    Add-Content -Path $ExportLog -Value '<th>Type</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Checkpoint</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # For each virtual machines, exporting information
    Foreach ($VM in $VMHostsWorkloadInformation){
        if ($VM.Name -notlike $Null){
            $vDiskColsTemp = @()
            $vDiskNbr      = 0
            $SnapTemp      = $Null

            # For each VHD, export HTML to temporary array
            Foreach ($vDisk in $VM.VMDisks){

                # this variable enables to calculate the HTML rowspan
                $vDiskNbr++
                # If VM is in Gen2 and the VHD controller are not SCSI, export CSS Class error
                if (($VM.Generation -eq 2) -and ($vDisk.ControllerType -notlike "SCSI")){
                    $Class = $TDError
                }
                Else {
                    $Class = $TDOK
                }
                $vDiskColsTemp += "<td class=$class>$($vDisk.ControllerType)</td>"

                $vDiskName      = Split-Path $vDisk.Path -Leaf

                # If the Virtual Disk is not a VHDX, export CSS Class error
                if ($vDiskName -notlike "*.vhdx"){
                    $Class = $TDError
                }
                Else{
                    $Class = $TDOK
                }
                $vDiskColsTemp += "<td class=$Class>$($vDisk.Path)</td>"
                $vDiskColsTemp += "<td>$($vDisk.VHDType)</td>"
                $vDiskColsTemp += "<td><span class=$ValueClass>$([math]::Round($vDisk.Size/1GB, 1))</span><span class=$UnitClass>GB</span></td>"
            }

            # for each checkpoint, exporting to temporary array
            Foreach ($SnapShot in $VM.CheckPoints){
                $SnapTemp += "$($SNapShot.Name) ($($SnapShot.CreationTime))<br>"
            }
            Add-Content -Path $ExportLog -Value "<tr>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr class=$ComputerClass>$($VM.Name)</td>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr>$($VM.VMHost)</td>"

            # Azure Local platform VMs (Arc Resource Bridge, ...) are flagged so they are not mistaken for customer workloads
            If ($VM.IsInfrastructure){
                Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr><span class=$AdvertMessage>Azure Local infrastructure</span></td>"
            }
            Else {
                Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr>Workload</td>"
            }

            # if the VM is not running, export CSS class error
            If ($VM.State -like "Running"){
                $Class = $TDOK
            }
            Else{
                $Class = $TDError
            }
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr class=$class><span class=$ValueClass>$($VM.State)</span></td>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr><span class=$ValueClass>$($VM.Generation)</span></td>"

            # if the VM is not clustered, export CSS class error and change the true/false result by Yes/No
            if ($VM.IsClustered -like "true"){
                $Class = $TDOK
                $IsClustered = "Yes"
            }
            Else{
                $Class = $TDError
                $IsClustered = "No"
            }
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr>$IsClustered</td>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr><span class=$ValueClass>$($VM.ProcessorCount)</span><span class=$UnitClass>vCPU</span></td>"

            # Changing the True/False value by Yes/No
            If ($VM.DynamicMemoryEnabled -like "true"){
                $DynamicMemory = "Yes"
            }
            Else {
                $DynamicMemory = "No"
            }

            # Exporting the first line of the row with the RowSpan
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr>$DynamicMemory</td>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr><span class=$ValueClass>$([math]::round($VM.MemoryDemand/1GB, 1))</span><span class=$UnitClass>GB</span></td>"
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr><span class=$ValueClass>$([math]::round($VM.MemoryAssigned/1GB, 1))</span><span class=$UnitClass>GB</span></td>"
            Add-Content -Path $ExportLog -Value $vDiskColsTemp[0]
            Add-Content -Path $ExportLog -Value $vDiskColsTemp[1]
            Add-Content -Path $ExportLog -Value $vDiskColsTemp[2]
            Add-Content -Path $ExportLog -Value $vDiskColsTemp[3]
            Add-Content -Path $ExportLog -Value "<td RowSpan=$vDiskNbr>$SnapTemp</td>"
            Add-Content -Path $ExportLog -Value "</tr>"

            # Exporting the other line outside rowspan
            For ($i =3; $i -lt ($vDiskNbr*4)-1; $i += 4){
                Add-Content -Path $ExportLog -Value "<tr>"
                Add-Content -Path $ExportLog -Value $vDiskColsTemp[$i+1]
                Add-Content -Path $ExportLog -Value $vDiskColsTemp[$i+2]
                Add-Content -Path $ExportLog -Value $vDiskColsTemp[$i+3]
                Add-Content -Path $ExportLog -Value $vDiskColsTemp[$i+4]
                Add-Content -Path $ExportLog -Value "</tr>"
            }
        }
    }
    Add-Content -Path $ExportLog -Value "</table>"
}

# Export Storage Spaces Direct information to HTML
if ($ClusterS2D){
    # Show progress bar
    Write-Progress -Activity "HTML file construction" -PercentComplete 70 -CurrentOperation "Storage Spaces Direct information gathering"

    #export Table header and title
    Add-Content -Path $ExportLog -Value "<H2>Storage Spaces Direct</H2>"
    
    # Export title and table header for Storage Pool
    Add-Content -Path $ExportLog -Value "<H3>Storage Pool Information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Friendly Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Health</th>'
    Add-Content -Path $ExportLog -Value '<th>Operational Status</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Allocated Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Free Space</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # Add a row for each Storage Pool
    Foreach ($StoragePool in $StoragePoolInformation){

        Add-Content -Path $ExportLog -Value "<tr>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($StoragePool.FriendlyName)</td>"
        
        # change the class depending on healthy or not
        If ($StoragePool.HealthStatus -notlike "Healthy"){
            $Class = $TDError
        }
        Else {
            $Class = $TDOK
        }

        # change the class depending on healthy or not
        Add-Content -Path $ExportLog -Value "<td class=$class>$($StoragePool.HealthStatus)</td>"
        If ($StoragePool.OperationalStatus -notlike "OK"){
            $Class = $TDError
        }
        Else {
            $Class = $TDOK
        }
        # export value formatted for human (yes you !)
        Add-Content -Path $ExportLog -Value "<td class=$class>$($StoragePool.OperationalStatus)</td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($StoragePool.Size/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($StoragePool.AllocatedSize/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round((($StoragePool.Size - $StoragePool.AllocatedSize)/1TB), 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "</tr>"
    }

    Add-Content -Path $ExportLog -Value "</table>"

    # export information about virtual disks
    Add-Content -Path $ExportLog -Value "<H3>Virtual disk Information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Friendly Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Health</th>'
    Add-Content -Path $ExportLog -Value '<th>Number of columns</th>'
    Add-Content -Path $ExportLog -Value '<th>Resiliency</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Footprint on pool</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # export virtual disks sorted by virtual disk name (ASC)
    Foreach ($VirtualDisk in ($VirtualDiskInformation | sort -Property FriendlyName)){

        Add-Content -Path $ExportLog -Value "<tr>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($VirtualDisk.FriendlyName)</td>"
        If ($VirtualDisk.HealthStatus -notlike "Healthy"){
            $Class = $TDError
        }
        Else {
            $Class = $TDOK
        }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($VirtualDisk.HealthStatus)</td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$($VirtualDisk.NumberOfColumns)</span></td>"
       
        # format text for resiliency
        If (($VirtualDisk.ResiliencySettingName -like "Mirror") -and ($VirtualDisk.NumberOfDataCopies -eq 2)){
            $Resiliency = "2-Way Mirroring"
        }
        ElseIf (($VirtualDisk.ResiliencySettingName -like "Mirror") -and ($VirtualDisk.NumberOfDataCopies -eq 3)){
            $Resiliency = "3-Way Mirroring"
        }
        ElseIf (($VirtualDisk.ResiliencySettingName -like "Parity") -and ($VirtualDisk.NumberOfDataCopies -eq 2)){
            $Resiliency = "Simple Parity"
        }
        ElseIf (($VirtualDisk.ResiliencySettingName -like "Parity") -and ($VirtualDisk.NumberOfDataCopies -eq 3)){
            $Resiliency = "Dual Parity"
        }
        Add-Content -Path $ExportLog -Value "<td>$Resiliency</td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($Virtualdisk.Size/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($Virtualdisk.FootprintOnPool/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "</tr>"
    }

    Add-Content -Path $ExportLog -Value "</table>"

    # Export physical disk information
    Add-Content -Path $ExportLog -Value "<H3>Physical disk information</H3>"
    Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
    Add-Content -Path $ExportLog -Value '<tr>'
    Add-Content -Path $ExportLog -Value '<th>Friendly Name</th>'
    Add-Content -Path $ExportLog -Value '<th>Storage Pool</th>'
    Add-Content -Path $ExportLog -Value '<th>Health</th>'
    Add-Content -Path $ExportLog -Value '<th>Operational Status</th>'
    Add-Content -Path $ExportLog -Value '<th>Firmware version</th>'
    Add-Content -Path $ExportLog -Value '<th>Model</th>'
    Add-Content -Path $ExportLog -Value '<th>Serial number</th>'
    Add-Content -Path $ExportLog -Value '<th>Media type</th>'
    Add-Content -Path $ExportLog -Value '<th>Bus type</th>'
    Add-Content -Path $ExportLog -Value '<th>Usage</th>'
    Add-Content -Path $ExportLog -Value '<th>Size</th>'
    Add-Content -Path $ExportLog -Value '<th>Allocated size</th>'
    Add-Content -Path $ExportLog -Value '<th>Free Size</th>'
    Add-Content -Path $ExportLog -Value '</tr>'

    # Physical disk are exported sorted by the storage pool name
    Foreach ($PhysicalDisk in ($PhysicalDiskInformation | sort -Property StoragePoolFriendlyName)){

        Add-Content -Path $ExportLog -Value "<tr>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($PhysicalDisk.FriendlyName)</td>"
        Add-Content -Path $ExportLog -Value "<td class=$ComputerClass>$($PhysicalDisk.StoragePoolFriendlyName)</td>"

        If ($PhysicalDisk.HealthStatus -notlike "Healthy"){
            $Class = $TDError
        }
        Else {
            $Class = $TDOK
        }
        Add-Content -Path $ExportLog -Value "<td class=$Class>$($PhysicalDisk.HealthStatus)</td>"

         If ($PhysicalDisk.OperationalStatus -notlike "OK"){
            $Class = $TDError
        }
        Else {
            $Class = $TDOK
        }

        Add-Content -Path $ExportLog -Value "<td class=$Class>$($PhysicalDisk.OperationalStatus)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.FirmwareVersion)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.Model)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.SerialNumber)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.MediaType)</td>"
        Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.BusType)</td>"
        if ($($PhysicalDisk.Usage) -like "Journal"){
            Add-Content -Path $ExportLog -Value "<td>Cache<br><Span class=$AdvertMessage>Journal</Span></td>"   
        }
        Elseif ($($PhysicalDisk.Usage) -like "Auto-Select"){
            Add-Content -Path $ExportLog -Value "<td>Capacity<br><Span class=$AdvertMessage>Auto-Select</Span></td>"  
        }
        Else {
            Add-Content -Path $ExportLog -Value "<td>$($PhysicalDisk.Usage)</td>"  
        }
        
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($PhysicalDisk.Size/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round($PhysicalDisk.AllocatedSize/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "<td><span class=$ValueClass>$([math]::round(($PhysicalDisk.Size - $PhysicalDisk.AllocatedSize)/1TB, 2))</span><span class=$UnitClass>TB</span></td>"
        Add-Content -Path $ExportLog -Value "</tr>"
    }

    Add-Content -Path $ExportLog -Value "</table>"

    # Export SMB MultiChannel information
    Foreach ($VMHost in $VMHosts){
        $VMHostConnection = $VMHostSMBMultiChannel |? VMHost -like $VMHost.Name
        Add-Content -Path $ExportLog -Value "<H3>SMB MultiChannel SBL connection for $($VMHost.Name)</H3>"
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Client IP Address</th>'
        Add-Content -Path $ExportLog -Value '<th>Client NIC name</th>'
        Add-Content -Path $ExportLog -Value '<th>Client RSS Capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Client RDMA capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Server IP Address</th>'
        Add-Content -Path $ExportLog -Value '<th>Server NIC index</th>'
        Add-Content -Path $ExportLog -Value '<th>Server RSS Capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Server RDMA Capable</th>'
        Add-Content -Path $ExportLog -Value '</tr>'

        Foreach ($SBLConnection in ($VMHostConnection |? ConnectionType -like "SBL")){
            Add-Content -Path $ExportLog -Value "<tr>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ClientIP)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ClientNIC)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ClientRSS)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ClientRDMA)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ServerIP)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ServerNIC)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ServerRSS)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($SBLConnection.ServerRDMA)</td>"
            Add-Content -Path $ExportLog -Value "</tr>"  
         
        
        }
        Add-Content -Path $ExportLog -Value "</table>"
        Add-Content -Path $ExportLog -Value "<H3>SMB MultiChannel CSV connection for $($VMHost.Name)</H3>"
        Add-Content -Path $ExportLog -Value "<table Class=$TableClass>"
        Add-Content -Path $ExportLog -Value '<tr>'
        Add-Content -Path $ExportLog -Value '<th>Client IP Address</th>'
        Add-Content -Path $ExportLog -Value '<th>Client NIC name</th>'
        Add-Content -Path $ExportLog -Value '<th>Client RSS Capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Client RDMA capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Server IP Address</th>'
        Add-Content -Path $ExportLog -Value '<th>Server NIC index</th>'
        Add-Content -Path $ExportLog -Value '<th>Server RSS Capable</th>'
        Add-Content -Path $ExportLog -Value '<th>Server RDMA Capable</th>'
        Add-Content -Path $ExportLog -Value '</tr>'
  
        Foreach ($CSVConnection in ($VMHostConnection |? ConnectionType -like "CSV")){
            Add-Content -Path $ExportLog -Value "<tr>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ClientIP)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ClientNIC)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ClientRSS)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ClientRDMA)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ServerIP)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ServerNIC)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ServerRSS)</td>"
            Add-Content -Path $ExportLog -Value "<td>$($CSVConnection.ServerRDMA)</td>"
            Add-Content -Path $ExportLog -Value "</tr>"  
         
        
        }
        Add-Content -Path $ExportLog -Value "</table>" 
    }
}

# Add HTML footer (closing body and html)
Add-Content -Path $ExportLog -Value $HTMLEnding

Write-Host "You can find the HTML File here: $ExportLog" -ForegroundColor Green -BackgroundColor Black
Write-Host "This script was executed by - $ConsultantName $ConsultantEmail" -ForegroundColor Green -BackgroundColor black

if (-not $NoPdf) {
    Write-Host "Generating PDF report..." -ForegroundColor Cyan -BackgroundColor Black
    $PdfCreated = Convert-HtmlReportToPdf -HtmlPath $ExportLog -PdfPath $PdfLog
    if ($PdfCreated) {
        Write-Host "You can find the PDF File here: $PdfLog" -ForegroundColor Green -BackgroundColor Black
    }
}
