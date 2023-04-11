<#
PoweShell script to find hardware information for w32 system via WMI calls, dumping data to JSON file
#>



<#
### attempts to make script networkable ###
# Get the target hostname or IP address from command-line arguments
if ($args.Count -gt 0) {
    $target = $args[0]
} else {
    $target = "localhost"
}
# Prompt for credentials if needed
if (!(Test-Connection $target -Quiet)) {
    Write-Host "Error: $target is offline or unreachable"
    return
}
if ($cred -eq $null) {
    $cred = Get-Credential
}
# Define the script block to execute remotely
$scriptBlock = {
    $system = Get-CimInstance CIM_ComputerSystem
    $bios   = Get-CimInstance CIM_BIOSElement
    return $system, $bios
}
# Invoke the script block remotely with the specified credentials
#$system, $bios = Invoke-Command -ComputerName $target -Credential $cred -ScriptBlock $scriptBlock
#>



###     Get system information
$system =   Get-CimInstance CIM_ComputerSystem
$bios   =   Get-CimInstance CIM_BIOSElement

$spec   =   @{"_comment" = "System Hardware Specification"}     #   System spec hashtable/dictionary

#####   Base system info    #####   
$mobo   =   @{
    "_comment"      =   "Motherboard information"
    "Hostname"      =   $system.Name
    "Manufacturer"  =   $system.Manufacturer
    "Model"         =   $system.Model
    "SerialNumber"  =   $bios.SerialNumber
    "SystemSKUNumber" = $system.SystemSKUNumber
}  
$spec["BIOS"]       =   $mobo


#####   CPU     #####   
$processor  =   Get-CimInstance CIM_Processor
$cpu = @{
    "_comment"  =   "CPU information"
    "Name"      =   $processor.Name
    "Cores"     =   $processor.NumberOfCores
    "Speed"     =   [math]::Round($processor.MaxClockSpeed / 1e3, 2)    
}
$spec["CPU"]    =   $cpu


#####   Memory  #####
$ram    =   @{"_comment"  =   "RAM information"}     #   ram hashtable/disctionary object
#$ramModule  =   @()     #   init array for individual RAM modules    

$memory =   Get-CimInstance -ClassName Win32_PhysicalMemory
$slots  =   Get-CimInstance -ClassName Win32_PhysicalMemoryArray

#$ramMax     =   [math]::Round($system.TotalPhysicalMemory/1GB)
$ram["Total"]   =   [int][math]::Round($system.TotalPhysicalMemory / 1GB)    #   Maximum RAM board can handle
$ram["Maximum"] =   [int][math]::Round($slots.MaxCapacityEx / 1MB)         #   Maximum RAM board can handle
#$ramType    =   $slots.MemoryType                                      # Get the type of RAM modules the board requires

$ram["Type"]    =   switch($slots.MemoryType) {
                        default {''}        #   say nothing if unsure
                        20      {'DDR'}
                        21      {'DDR2'}
                        24      {'DDR3'}
                        26      {'DDR4'}
                    }  #   what kinda RAM do we like

#$ramSlots   =   $memory.MemoryDevices
$ram["Count"]  =   $memory.Count               #   number of modules loaded
$ram["Slots"]  =   $slots.MemoryDevices     # Total number of memory slots

# Get the memory modules installed in the memory slots
$ram["Modules"] =   @()     #   init array for individual RAM modules
foreach($module in $memory) {
        $ram["Modules"]     +=  @{
            "_comment"          = "Module information"
            "BankLabel"         = $module.BankLabel
            "DeviceLocator"     = $module.DeviceLocator
            "Manufacturer"      = $module.Manufacturer 
            "PartNumber"        = $module.PartNumber
#            "SerialNumber"      = $module.SerialNumber      ?? ""
#            "SerialNumber"      = $module.SerialNumber -ne $null ? $module.SerialNumber : ""
            "SerialNumber"      = if ($module.SerialNumber -ne $null) { $module.SerialNumber } else { "" }
            "Capacity"          = [int][math]::Round($module.Capacity / 1GB)
            "Speed"             = $module.Speed
<#
            "MemoryType"        = switch ($module.MemoryType) {
                                    default { '' }  # say nothing if unsure
                                    20 {'DDR'}
                                    21 {'DDR2'}
                                    24 {'DDR3'}
                                    26 {'DDR4'}
                                    }
            "TypeDetail"        =   switch ($module.TypeDetail) {
                                    default { '' }  # say nothing if unsure
                                    1 { 'Other' }
                                    2 { 'Unknown' }
                                    11 { 'Cache DRAM' }
                                    12 { 'Non-volatile' }
                                    13 { 'Registered (Buffered)' }
                                    14 { 'Unbuffered (Unregistered)' }
                                    15 { 'LRDIMM' }
                                    16 { 'DDR2 FB-DIMM' }
                                    17 { 'DDR-SDRAM' }
                                    18 { 'DDR2-SDRAM' }
                                    19 { 'DDR2-SDRAM FB-DIMM' }
                                    20 { 'DDR3-SDRAM' }
                                    21 { 'FBD2-SDRAM' }
                                    22 { 'DDR4-SDRAM' }
                                    24 { 'LPDDR-SDRAM' }
                                    25 { 'LPDDR2-SDRAM' }
                                    26 { 'LPDDR3-SDRAM' }
                                    27 { 'LPDDR4-SDRAM' }
                                    28 { 'Logical non-volatile device' }
                                    29 { 'HBM (High Bandwidth Memory)' }
                                    30 { 'HBM2 (High Bandwidth Memory 2)' }
                                }
#>
            "SMBIOSMemoryType"  = switch($module.SMBIOSMemoryType) {
                                    default {''}
                                    1     {'Other'}
                                    2     {'Unknown'}
                                    3     {'DRAM'}
#                                    4     {'EDRAM'}
#                                    7     {'RAM'}
#                                    8     {'ROM'}
#                                    9     {'Flash'}
#                                    14    {'3DRAM'}
                                    15    {'SDRAM'}
                                    17    {'RDRAM'}
                                    18    {'DDR'}
                                    19    {'DDR2'}
                                    20    {'DDR2 FB-DIMM'}
                                    21    {'DDR3'}
                                    22    {'FBD2'}
                                    24    {'DDR4'}
                                    25    {'LPDDR'}
                                    26    {'LPDDR2'}
                                    27    {'LPDDR3'}
                                    28    {'LPDDR4'}
                                    29    {'Logical non-volatile device'}
                                    30    {'HBM2'}
                                    }
<#
            "FormFactor" = switch($module.FormFactor) {
                                    1 {'Other'}
                                    6 {'Proprietary'}
                                    7 {'SIMM'}
                                    8 {'DIMM'}
                                    12 {'SODIMM'}
                                    14 {'FB-DIMM'}
                                    default {''}
                                }
#>
#            "Rank"              = $module.Rank              ?? ""
#            "OtherTimingInfo"   = $module.OtherTimingInfo   ?? ""
            "Rank"              = if ($module.Rank -ne $null) { $module.Rank } else { "" }
            "OtherTimingInfo"   = if ($module.OtherTimingInfo -ne $null) { $module.OtherTimingInfo } else { "" }

        }
}
$spec["RAM"]    =   $ram


#####   HDD     #####
$diskDrive      =   Get-CimInstance CIM_DiskDrive | `
                        Where-Object {$_.MediaType -ne "Removable Media" -and $_.MediaType -ne $null -and $_.Size -ne 0}
$hdd            =   @()     #   store drives as array
foreach($disk in $diskDrive) {
    $hdd   += @{
        "_comment"  = "HDD information"
        "Model"     = $disk.Model
        "Size"      = [int][math]::Round($disk.Size / 1gb, 2)
        "MediaType" = $disk.MediaType   #   need to determine if SSD or platter
    }
}
$spec["HDD"]    =   $hdd


#####   Networking  #####
#$adapters   =   Get-CimInstance -ClassName Win32_NetworkAdapter | `
#                    Where-Object { $_.PhysicalAdapter -eq $true -and $_.AdapterType -ne $null `
#                        -and $_.MacAddress -ne "00-00-00-00-00-00" -and $_.MacAddress -ne "00:00:00:00:00:00"  `
#                        -and $_.Name -notlike "*Loopback*"  -and $_.Name -notlike "*Virtual*"}
$adapters   =   Get-NetAdapter -Physical -IncludeHidden | `
                        Where-Object { `
                                $_.MACAddress -ne "00-00-00-00-00-00" -and $_.MACAddress -ne "00:00:00:00:00:00"  `
                            -and $_.Name -notlike "*Loopback*"  -and $_.Name -notlike "*Virtual*"}

$nic        =   @{"_comment"  =  "Networking information"}     #   NIC's as dict
foreach($interface in $adapters) {
    $nic[$interface.MacAddress]    +=  @{
        "_comment"      =   "Interface/card information"
        "MAC"           =   $interface.MacAddress
        "Name"          =   $interface.Name
#        "ServiceName"   =   $interface.ServiceName
#        "Description"           =   $interface.Description
        "InterfaceDescription"  =   $interface.InterfaceDescription        
        #        "Speed"         =   [int][math]::Round($interface.Speed / 1MB, 2)    #   NEEDS WORK - CAN'T gGET MAX LINKSPEED, only current
        "LinkSpeed"     =   $interface.LinkSpeed    #   NEEDS WORK
        }
}
$spec["NIC"]    =   $nic


# GPU information
$graphics   =   Get-CimInstance CIM_VideoController
$gpu        =   @()     #   GPU's as array/list
foreach($card in $graphics) {
    $gpu   +=  @{
        "_comment"  =   "GPU information"
        "Model"     =   $card.Description
        "RAM"       =   [int][math]::Round($card.AdapterRAM / 1GB, 2)
#        "Resolution"    =   $graphics[$card].VideoModeDescription   #   find max resolution by GPU
    }
}
$spec["GPU"]    =   $gpu


# Dump Data
$specFileName   =   $spec["BIOS"]["SystemSKUNumber"] +"-"+ $spec["BIOS"]["SerialNumber"]            #   base filename
#$specFileDB     =   "specdata.db"                            #   SQLite db

#JSON
#$spec | ConvertTo-Json -Depth 100
$spec | ConvertTo-Json -Depth 100 | Out-File -Encoding UTF8 -FilePath "$specFileName.json"
Write-Host "System specification saved to $specFileName.JSON"

Pause