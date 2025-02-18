$matcheddevices = ("DMPS-200-C", "TSW-760", "TSW-750")
$targetDevices = ("DMPS3", "TSW-750")

    $setupblock = 
    {
        Param([string]$dev, [string[]]$targets)
            if($dev -match $targets[0]) #if the device description contains the device we are looking to target, allows us to target multiple versions of devices.
            {
                echo "$($targets[0]) == $dev"
            }
            elseif($dev -match $targets[1])
            {
                echo "$($targets[1]) == $dev"
            }
    }

    foreach($device in $matcheddevices) 
    {
        Start-Job -Name $device -ScriptBlock $setupblock -ArgumentList $device, $targetDevices > $null
    }

    Get-Job | Receive-Job -Wait