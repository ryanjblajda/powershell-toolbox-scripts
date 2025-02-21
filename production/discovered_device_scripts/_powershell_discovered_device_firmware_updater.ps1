try 
{
    Import-Module PSCrestron
    Import-Module ImportExcel
}
catch 
{
    Write-Warning "Run the _powershell_crestron_setup script you must!! - Yoda, probably..."
    Write-Error $_
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

$updateDeviceFirmware = {
    param([PSObject]$dev, [string]$firmware, [string]$usr, [string]$pw)

    $firmwareFileName = $firmware | Split-Path -Leaf
    $model = $dev.Description.Split(" ")[0].ToLower()
    Write-Host -ForegroundColor Cyan "Uploading $firmwareFileName to: $($dev.Hostname) @ $($dev.IP) -> [$model]"

    Send-FTPFile -Device $dev.IP -LocalFile $firmware -RemoteFile "\firmware\$firmwareFileName" -Secure -Username $usr -Password $pw

    if($model -like "*DMPS3*" -or $model -like "*HD-DM*")
    {
        Write-Host "Invoking PUSHUPDATE FULL on $model @ $($dev.IP)"
        Invoke-CrestronCommand $dev.IP -Command "pushupdate full" -Secure -Username $usr -Password $pw
    }
    else {
        Write-Host "Invoking PUF on $model @ $($dev.IP)"
        Invoke-CrestronCommand $dev.IP -Command "puf" -Secure -Username $usr -Password $pw
    }
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
        $response = Read-Host -Prompt "Enter the model names of the devices you wish to update/configure, separated by a comma"
        #get rid of the hidden crap we dont want
        $response = Get-Flattened $response
        #get rid of any spaces, and split the entries based on the comma to create a list
        $targetDevices = $response.Replace(" ", "").Split(",")
        #spit out a list to show the user exactly what devices we are going to configure so that we can make sure what they wanted is correct
        Write-Host "You wish to update firmware on the following types of devices: $targetDevices" -ForegroundColor Yellow
        #loop through all the target devices
        $targetDeviceObjects = [System.Collections.ArrayList]::new()
        foreach($item in $targetDevices) {
            $item = Get-Flattened $item
            #find where we have a matching device in the discovered devices table
            $results = $discoveredCrestronDevices | Where-Object { $_.Description.Split(" ")[0].ToLower().Contains($item) }
            #show the devices that have been added to the list of devices to update
            if ($null -ne $results) {
                foreach($added in $results) {
                    $targetDeviceObjects.Add($added) > $null
                    Write-Host "$($added.Hostname) @ $($added.IP)" -ForegroundColor DarkGreen
                }
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

    $firmwareFileDictionary = @{}

    #determine the correct firmware file
    foreach ($deviceModel in $targetDevices) {
        Write-Host "Please select the firmware file to upload to $deviceModel devices" -ForegroundColor Yellow
        do {
            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $dialog.AddExtension = $true
            $dialog.Filter = 'All Files |*.*'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = "$HOME\Downloads"
            $dialog.RestoreDirectory = $true
            $dialog.Title = 'Select Firmware File'
            $dialog.ShowDialog()

            $firmwareFilePath = $dialog.FileName
            $firmwareFileName = $firmwareFilePath | Split-Path -Leaf

            Write-Host "You have selected the following file to upload to $deviceModel devices: $firmwareFileName" -ForegroundColor DarkGreen
            $response = Read-Host "Enter [y] to confirm this is the correct file, [n] to select a different file"
            $firmwareFileConfirmed = Get-Flattened $response
        }
        while($firmwareFileConfirmed -ne 'y')
        #assign the firmware file to the model dictionary for later use
        $model = $deviceModel.ToLower()
        $firmwareFileDictionary[$model] = $firmwareFilePath
    }

    #for each command/device create a job
    Write-Host "Beginning Jobs..." -Foreground DarkGreen
    Write-Host $firmwareFileDictionary

    $actions = {
        foreach($device in $targetDeviceObjects) {
            $model = $device.Description.Split(" ")[0].ToLower()
            $firmware = $firmwareFileDictionary[$model]
            #if the user entered just tsw, or dge, then 
            if ($firmware -eq $null) {
                do {
                    Write-Warning "There is no dictionary entry for $model devices!"
                    $response = Read-Host -Prompt "Enter [y] to specify the correct firmware file for $model devices, or [n] to skip these devices."
                    $update = Get-Flattened $response
                } 
                while(($update -ne 'y') -and ($update -ne 'n'))
                
                if ($update -eq 'y') {
                    Write-Host "Please select the firmware file to upload to $model devices" -ForegroundColor Yellow
                    do {
                        $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
                        $dialog.AddExtension = $true
                        $dialog.Filter = 'All Files |*.*'
                        $dialog.Multiselect = $false
                        $dialog.InitialDirectory = "$HOME\Downloads"
                        $dialog.RestoreDirectory = $true
                        $dialog.Title = 'Select Firmware File'
                        $dialog.ShowDialog()
            
                        $firmwareFilePath = $dialog.FileName
                        $firmwareFileName = $firmwareFilePath | Split-Path -Leaf
            
                        Write-Host "You have selected the following file to upload to $model devices: $firmwareFileName" -ForegroundColor DarkGreen
                        $response = Read-Host "Enter [y] to confirm this is the correct file, [n] to select a different file"
                        $firmwareFileConfirmed = Get-Flattened $response
                    }
                    while($firmwareFileConfirmed -ne 'y')

                    $firmwareFileDictionary[$model] = $firmwareFilePath
                }
            }
            else { 
                $update = 'y'
            }
            
            if ($update -eq 'y') {
                Write-Host -ForegroundColor DarkGreen "Beginning Send $($device.Hostname) @ $($device.IP) --> $firmware"
                Start-Job -ScriptBlock $updateDeviceFirmware -ArgumentList $device, $firmware, $username, $password -Name $device.Hostname 
            }
        }
        Get-Job | Receive-Job -Wait
    }

    #invoke script block
    Invoke-Command $actions

    #allow us to quickly just re-do exactly what we just did for quick troubleshooting updates
    do {
        Write-Host "All Jobs Completed!" -ForegroundColor Yellow
        Write-Host "Enter [e] to exit this script immediately" -ForegroundColor Red
        Write-Host "Enter [n] to re-run through the discovery & configuration process anew" -ForegroundColor Cyan
        Write-Host "Enter [r] to re-run this script exactly as you have it configured once more." -ForegroundColor DarkGreen
        $response = Read-Host
        $exit = Get-Flattened $response
        if ($exit -eq 'r') {
            #invoke script block
            Invoke-Command $actions
        }
    } while(($exit -ne 'n') -and ($exit -ne 'e'))

} while($exit -ne 'e')

exit