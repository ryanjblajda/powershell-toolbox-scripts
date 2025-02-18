﻿try 
{
    Import-Module PSCrestron
    Import-Module ImportExcel
}
catch 
{
    Write-Warning "Run the _powershell_crestron_setup script you must!! - Yoda, probably..."
    Write-Error $_
}

$SendCommand = { 
    param ([PSObject]$device, [string]$command, [string]$user, [string]$pw)

    Invoke-CrestronCommand -Device $device.IP -Command $command -Secure -Username $user -Password $pw | Write-Host -Foreground Green
}

function Get-Flattened {
    param([string]$ToFlatten)
       
    $result = $ToFlatten.ToLower()
    $result = $result -replace "`r`n", ""

    return($result)
}
function Get-Devices {
    do {
        $showRetryExitPrompt = $false
        try {
            Write-Host "Discovering Crestron Devices...." -ForegroundColor DarkGreen
            $discovered = Get-AutoDiscovery -ShowProgress

            if($null -eq $discovered)
            {
                Write-Warning -Message "No Crestron devices were found on the network!`nPlease make sure you are connected to the right network, and you have an IP Address in the correct subnet`n`n**sad trombone noises**"
                $showRetryExitPrompt = $true
            }
            else {    
                $exitResponse = 'continue'
            }
        }
        catch {
            Write-Warning -Message "Looks like your aren't connected to any networks`nMake sure your are actually connected a network switch or Crestron device`n`n**sad trombone noises**"
            Write-Error $_
            $showRetryExitPrompt = $true
        }
        
        if ($showRetryExitPrompt) {
            $exitResponse = Read-Host -Prompt "Enter [e] to exit, or [r] re-run discovery again"
            $exitResponse = Get-Flattened -ToFlatten $exitResponse
        }
    }
    while(($exitResponse -ne "e") -and ($exitResponse -ne "continue"))

    if ($exitResponse -eq "e") { Exit }

    return ($discovered)
}

#loop our entire main loop for updating stuff often.
do {
    #discover devices and handle if the computer discovers nothing
    $discoveredCrestronDevices = Get-Devices
    #display the output from Get-Devices
    Write-Output $discoveredCrestronDevices | Out-Default

    #get the desired devices to configure
    do {
        #get the list of devices that the user wants to configure
        $response = Read-Host -Prompt "Enter the hostnames of the devices you wish to update/configure, separated by a comma"
        #get rid of the hidden crap we dont want
        $response -replace "`r`n", "" > $null
        #get rid of any spaces, and split the entries based on the comma to create a list
        $targetDevices = $response.Replace(" ", "").Split(",")
        #spit out a list to show the user exactly what devices we are going to configure so that we can make sure what they wanted is correct
        Write-Host "You wish to configure the following devices:" -ForegroundColor Yellow
        #loop through all the target devices
        $targetDeviceObjects = [System.Collections.ArrayList]::new()
        foreach($item in $targetDevices) {
            $item = Get-Flattened $item
            #find where we have a matching device in the discovered devices table
            $results = $discoveredCrestronDevices | Where-Object { $_.Hostname.ToLower() -eq $item.ToLower() }
            #if we find a matching result, print its details to the console
            if ($null -ne $results) {
                $targetDeviceObjects.Add($results[0]) > $null
                Write-Host "$item @ $($results[0].IP)" -ForegroundColor DarkGreen
                $results | Add-Member -MemberType AliasProperty -Name Command -Value "info"
            }
        }
        #prompt the use to confirm this is correct.
        $desiredDevicesCorrect = Read-Host -Prompt "Enter [y] to confirm, [n] to re-enter desired device list"
    }
    while($desiredDevicesCorrect -ne "y")

    #determine the credentials desired
    do {
        #store default credentials by default
        $username = "admin"
        $password = "CCS$erv!ce"
        
        $response = Read-Host -Prompt "Use CCS default credentials? --- Enter [y] to confirm, [n] to enter custom SSH credentials"
        $useDefaultCredentials = Get-Flattened $response
        
        if($useDefaultCredentials -eq "y") {
            #we want to exit this loop, so set variable to y
            $credentialsConfirmed = "y"
        }
        elseif($useDefaultCredentials -eq "n")
        {
            #if the user needs to use custom credentials, we will loop through making sure that everything is correct before continuing
            do {
                $username = Read-Host -Prompt 'Please enter the desired username'
                $password = Read-Host -Prompt 'Please enter the desired password'
                Write-Host "You desire the following credentials | username: $username // password: $password" -ForegroundColor DarkGreen
                $response = Read-Host -Prompt "Enter [y] to confirm, [n] to re-enter credentials" 
                $credentialsConfirmed = Get-Flattened $response
            }
            while ($credentialsConfirmed -ne "y")
        }
        else 
        {
            Write-Warning -Message "Danger Will Robinson!`nYou've entered an invalid response. I know not what to do with this nonsense: $credentialsConfirmed`nEnter [y] to use the default crestron credentials, or [n] to use SSH credentials you provide"
        }       
    }
    while($credentialsConfirmed -ne "y")

    #we will ask the user how they wish to configure the devices, whether they want to send a batch of commands to all devices, or a unique command to each device
    do {
        Write-Host "How would you like to send commands to devices?"
        Write-Host "Enter [b] to send a single command to all devices [batch mode]" -ForegroundColor DarkGreen
        Write-Host "Enter [u] to send a unique command to each device [unique mode]" -ForegroundColor Cyan
        $response = Read-Host
        $mode = Get-Flattened $response
        $modeString = "Invalid!"
        if ($mode -eq 'b') { $modeString = 'Batch'}
        elseif ($mode -eq 'u') { $modeString = 'Unique'}
        
        if ($mode -eq 'b' -or $mode -eq 'u') { 
            Write-Host "You desire to send commands in $modeString mode?" -ForegroundColor Yellow
            $response = Read-Host -Prompt "Enter [y] to confirm, [n] to choose a different option"
            $modeConfirmed = Get-Flattened $response
        }
        else { 
            Write-Warning -Message "Danger Will Robinson!`nYou've entered an invalid response. I know not what to do with this nonsense: $mode!!"
        }
    }
    while($modeConfirmed -ne "y")

    #batch mode stuff
    if ($mode -eq 'b') {    
        do {
            #request the command from the user
            Write-Host "Please enter the command you would like to send as a batch to all devices you've selected" -Foreground Cyan
            $response = Read-Host
            $command = Get-Flattened $response
            #confirm the command entry
            Write-Host "You want to send: $command to -> ALL DEVICES <-" -ForegroundColor Yellow
            $response = Read-Host -Prompt "Enter [y] to confirm this command, [n] to re-type your desired command"
            $commandConfirmed = Get-Flattened $response
        }
        while($commandConfirmed -ne "y")
    }
    elseif($mode -eq 'u') {
        foreach ($device in $targetDeviceObjects) {
            #Write-Host $device
            do {
                #request the command from the user
                Write-Host "Please enter the command you would like to send to $($device.Hostname) @ $($device.IP)" -Foreground Cyan
                $response = Read-Host
                $command = Get-Flattened $response
                #confirm the command entry
                Write-Host "You want to send: $command to $($device.Hostname) @ $($device.IP)" -ForegroundColor Yellow
                $response = Read-Host -Prompt "Enter [y] to confirm this command, [n] to re-type your desired command"
                $commandConfirmed = Get-Flattened $response
            }
            while($commandConfirmed -ne "y")
        }
    }

    #for each command/device create a job
    Write-Host "Beginning Jobs..." -Foreground DarkGreen

    $actions = {
        if ($mode -eq "b") {
            foreach ($device in $targetDeviceObjects) {
                #Write-Host "Sending Command To $($device.Hostname) @ $($device.IP) $($device.Command) // Username: $username, Password: $password"
                Start-Job -ScriptBlock $SendCommand -ArgumentList $device, $command, $username, $password -Name $device.Hostname
            } 
        }
        elseif ($mode -eq "u") {
            foreach ($device in $targetDeviceObjects) {
                #Write-Host "Sending Command To $($device.Hostname) @ $($device.IP) $($device.Command) // Username: $username, Password: $password"
                Start-Job -ScriptBlock $SendCommand -ArgumentList $device, $device.Command, $username, $password -Name $device.Hostname
            }
        }

        Get-Job | Receive-Job -Wait
    }

    #invoke script block
    & $actions

    #allow us to quickly just re-do exactly what we just did for quick troubleshooting updates
    do {
        Write-Host "All Jobs Completed!" -ForegroundColor Yellow
        Write-Host "Enter [e] to exit this script immediately" -ForegroundColor Red
        Write-Host "Enter [n] to re-run through the discovery & configuration process anew" -ForegroundColor Cyan
        Write-Host "Enter [r] to re-run this script exactly as you have it configured once more." -ForegroundColor DarkGreen
        $exit = Read-Host
        $exit = $exit.ToLower()
        $exit -replace "`r`n", "" > $null
        if ($exit -eq 'r') {
            #invoke script block
            & $actions
        }
    } while(($exit -ne 'n') -and ($exit -ne 'e'))

} while($exit -ne 'e')

exit