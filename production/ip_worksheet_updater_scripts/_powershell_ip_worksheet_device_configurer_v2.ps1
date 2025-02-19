Import-Module ImportExcel
Import-Module PSCrestron
Add-Type -AssemblyName Microsoft.VisualBasic

try 
{
    echo "Discovering Crestron Devices...."
    $discovered = Get-AutoDiscovery
    if($null -eq $discovered)
    {
        Write-Warning -Message "No Crestron devices were found on the network!"
        Write-Warning -Message "Please resolve this, and re-run the script"
        Read-Host "Press Enter To Exit"
        Exit
    }
    else 
    {
        echo "We Discovered Crestron Devices, Please Find The File Browser Window, and select an IP Worksheet"
    }
}
catch 
{
    Write-Warning -Message "**insert sad trombone noise** `nLooks like your aren't connected to any networks...`nFix that and run me again!"
    Read-Host "Press Enter To Exit"
    Exit
}

echo "Available Network Devices`n-----------------------------"
foreach($device in $discovered)
{
    echo "$($device.Hostname) @ $($device.IP)"
}

#do stuff with our list of devices
if ($null -ne $matcheddevices)
{
    $confirmation = $null
    do {
        $defaultcreds = Read-Host -Prompt "Use Crestron default credentials? --- Enter [y] to confirm, [n] to enter custom SSH credentials"
        $defaultcreds.ToLower()
        if($defaultcreds -eq "y") {
            $username = "crestron"
            $password = ""
            $confirmation = "y"
        }
        elseif($defaultcreds -eq "n")
        {
            do {
                $username = Read-Host -Prompt 'Username'
                $userpw = Read-Host -Prompt 'Password'
                $confirmation = Read-Host -Prompt "USER: $username | PW: $userpw --- Enter [y] to confirm, [n] to re-enter credentials"
                $confirmation.ToLower() > $null
            }
            while ($confirmation -ne "y")
        }
        else 
        {
            Write-Warning -Message "Yo...type [y] or [n], not $defaultcreds"
        }       
    }
    while($confirmation -ne "y")

    $updatenetwork = $null
    $updateconfig = $null
    $updateprogram = $null
    $updatetouchpanel = $null

    #ask the end user if they want to update the network settings on the targeted devices
    do {
        $networkresponse = Read-Host -Prompt "Configure Networking Settings? --- Enter [y] to confirm, [n] to skip"
        $networkresponse.ToLower() > $null
        if($networkresponse -eq "y") {
            echo "The network settings on the targeted devices WILL be updated"
            $updatenetwork = $true
        }
        elseif($networkresponse -eq "n") {
            echo "The network settings on the targeted devices will NOT be updated"
            $updatenetwork = $false
        }
        else 
        {
            Write-Warning -Message "Yo...type [y] or [n], not $networkresponse"
        }       
    }
    while($null -eq $updatenetwork)

    #ask the end user if they want to update the program on the targeted devices
    do {
        $progresponse = Read-Host -Prompt "Update SIMPL Programming? --- Enter [y] to confirm, [n] to skip"
        $progresponse.ToLower() > $null
        if($progresponse -eq "y") {
            $updateprogram = $true
            do {
                $response = Read-Host -Prompt "Enter the model names of the Control Systems you wish to push SIMPL programming to, separated by a comma"
                $controlsys -replace "`r`n", "" > $null
                $controlsys = $response.Split(',').Trim()
                $targetconfirmation = Read-Host -Prompt "You wish to send new programming to $controlsys type devices? --- Enter [y] to confirm, [n] to try again"
                $targetconfirmation.ToLower() > $null
            }
            while($targetconfirmation -ne "y")
            
            echo "Use the file dialog to select a SIMPL program file"
            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $dialog.AddExtension = $true
            $dialog.Filter = 'Crestron Program | *.lpz; *.spz'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = "$HOME\Downloads"
            $dialog.RestoreDirectory = $true
            $dialog.Title = 'Select Program File'
            $dialog.ShowDialog()
            $progFile = $dialog.FileName
            
        }
        elseif($progresponse -eq "n") {
            Write-Warning "SIMPL Programming will not be updated on any device"
            $updateprogram = $false
        }
        else 
        {
            Write-Warning -Message "Yo...type [y] or [n], not $progresponse"
        }       
    }
    while($null -eq $updateprogram)

    #ask end user if they want to update config files
    do {
        $configresponse = Read-Host -Prompt "Update Config File? --- Enter [y] to confirm, [n] to skip"
        $configresponse.ToLower() > $null
        if($configresponse -eq "y")
        {
            $updateconfig = $true

            if($updateprogram -eq $false) #only ask this if we are updating configs, but not programs. if we are updating both we dont need to ask again.
            {
                do {
                    $response = Read-Host -Prompt "Enter the model names of the Control Systems you wish to push new config files to, separated by a comma"
                    $controlsys -replace "`r`n", "" > $null
                    $controlsys = $response.Split(',').Trim()
                    $targetconfirmation = Read-Host -Prompt "You wish to send new config files to $controlsys type devices? --- Enter [y] to confirm, [n] to try again"
                    $targetconfirmation.ToLower() > $null
                }
                while($targetconfirmation -ne "y")
            }
            echo "Config files WILL be updated on $controlsys`nNavigate to the Folder Browser; Select Config File Folder"
            Write-Warning "Dont forget to name config files according to device hostname!!!"
            
            $configFolder = New-Object System.Windows.Forms.FolderBrowserDialog
            $configFolder.Description = 'Select Config File Folder Location'
            $configFolder.ShowDialog()
            $configPath = $configFolder.SelectedPath
            
        }
        elseif($configresponse -eq "n")
        {
            $updateconfig = $false
            Write-Warning "Config files will NOT be updated"
        }
        else {
            Write-Warning -Message "Yo...type [y] or [n], not $progresponse"
        }
    }
    while($null -eq $updateconfig)

    #ask the end user if they want to update the ui on the targeted devices
    do {
        $touchpanelresponse = Read-Host -Prompt "Update UI File? --- Enter [y] to confirm, [n] to skip"
        $touchpanelresponse.ToLower() > $null
        if($touchpanelresponse -eq "y") {
            $updatetouchpanel = $true
            do {
                $response = Read-Host -Prompt "Enter the model names of the Interfaces you wish to send UI files to, separated by a comma"
                $interfaces -replace "`r`n", "" > $null
                $interfaces = $response.Split(',').Trim()
                $targetconfirmation = Read-Host -Prompt "You wish to send new programming to $interfaces type devices? --- Enter [y] to confirm, [n] to try again"
                $targetconfirmation.ToLower() > $null
            }
            while($targetconfirmation -ne "y")
            
            echo "Use the file dialog to select a UI file"
            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $dialog.AddExtension = $true
            $dialog.Filter = 'Crestron Touchpanel | *.vtz*'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = "$HOME\Downloads"
            $dialog.RestoreDirectory = $true
            $dialog.Title = 'Select Touchpanel File'
            $dialog.ShowDialog()
            $uiFile = $dialog.FileName
            
        }
        elseif($touchpanelresponse -eq "n") {
            Write-Warning "UI files will not be updated on any device"
            $updatetouchpanel = $false
        }
        else 
        {
            Write-Warning -Message "Yo...type [y] or [n], not $touchpanelresponse"
        }       
    }
    while($null -eq $updatetouchpanel)

    if($updateconfig){
        echo "$controlsys will be sent config files from $configPath"
    }
    if($updateprogram) {
        echo "$controlsys devices will have $progFile sent to them"
    }
    if($updatetouchpanel) {
        echo "$interfaces devices will have $uiFile sent to them"
    }

    $updateActions = ($updatenetwork, $updateprogram, $updateconfig, $updatetouchpanel)
    $targetDevices = ($controlsys, $interfaces)

    $setupblock = 
    {
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [PSObject]$dev,
            [string[]]$targets,
            [bool[]]$todo,
            [string]$cfg,
            [string]$prog,
            [string]$ui,
            [string]$usr, 
            [string]$pw)

            Write-Host -ForegroundColor Cyan "Starting setup for $($dev.Building) $($dev.'Room #') | $($dev.Model) : $($dev.'Serial #') | USR: $usr / PW: $pw"

            if($todo[0] -eq $true) { #if the end user wants to update network settings on devices
                Write-Host -ForegroundColor Cyan "Starting $($dev.Hostname) Network Setup"
                if($dev.'DHCP/Static'.ToLower() -eq "static")
                {
                    Write-Host -ForegroundColor Cyan "Setting $($dev.Model) @ $($dev.'IP Address') | $($dev.'Subnet Mask') | $($dev.Gateway)"
                    Invoke-CrestronCommand $dev.'Discovered IP' -Command "dhcp 0 off"
                    Invoke-CrestronCommand $dev.'Discovered IP' -Command "ipaddr 0 $($dev.'IP Address')" -Secure -Username $usr -Password $pw
                    Invoke-CrestronCommand $dev.'Discovered IP' -Command "defrouter 0 $($dev.Gateway)" -Secure -Username $usr -Password $pw
                    Invoke-CrestronCommand $dev.'Discovered IP' -Command "ipmask 0 $($dev.'Subnet Mask')" -Secure -Username $usr -Password $pw
                    $newip = $dev.'IP Address'
                }
                elseif($dev.'DHCP/Static'.ToLower() -eq "dhcp")
                {
                    Write-Host -ForegroundColor Cyan "Setting $($dev.Model) @ $($dev.'Discovered IP') to DHCP"
                    Invoke-CrestronCommand $dev.'Discovered IP' -Command "dhcp 0 on" -Secure -Username $usr -Password $pw
                    $newip = $dev.'Discovered IP'
                }

                Invoke-CrestronCommand $dev.'Discovered IP' -Command "hostname $($dev.Hostname)" -Secure -Username $usr -Password $pw
            }
            
            if($todo[1] -eq $true) { #if the end user wants to update programming
                
                if($dev.Model -match $targets[0]) #if the device description contains the device we are looking to target, allows us to target multiple versions of devices.
                {
                    Write-Host -ForegroundColor Cyan "Starting program update on $($dev.Model) @ $($dev.'Discovered IP')"
                    Send-CrestronProgram $dev.'Discovered IP' -LocalFile $prog -SendSIGFile -Secure -Username $usr -Password $pw -ShowProgress
                }
            }

            if($todo[2] -eq $true) { #if the end user wants to send new config file
                if($dev.Model -match $targets[0]) { #this will make sure that only the right devices get configs...not everything.
                    if([IO.File]::Exists("$cfg\$($dev.Hostname).json") -ne $false)
                    {
                        Write-Host -ForegroundColor Cyan "Starting config file update on $($dev.Model) @ $($dev.'Discovered IP')"
                        Send-FTPFile $dev.'Discovered IP' -LocalFile "$cfg\$($dev.Hostname).json" -RemoteFile '\USER\config.json' -Secure -Username $usr -Password $pw
                        Invoke-CrestronCommand $dev.'Discovered IP' -Command "progreset -p:1" -Secure -Username $usr -Password $pw
                    }
                    else 
                    {
                        Write-Warning "$($dev.Hostname).json DOES NOT EXIST!! Config file not uploaded, please create one and try again..."   
                    }
                }
            }

            if($todo[3] -eq $true) { #if the end user wants to update UI files
                if($dev.Model -match $targets[1]) { #if device is a match to the device we want to send UI files to
                    Write-Host -ForegroundColor Cyan "Starting UI update on $($dev.Model) @ $($dev.'Discovered IP')"
                    Send-CrestronProject $dev.'Discovered IP' -LocalFile $ui -Secure -Username $usr -Password $pw
                }
            }

            Write-Host -ForegroundColor Cyan "Completing setup for $($dev.Building) $($dev.'Room #') | $($dev.Model) : $($dev.'Serial #')"

            if($todo[0] -eq $true)
            {   
                Write-Host -ForegroundColor Cyan "Attempting to reboot $($dev.Model) : $($dev.'Serial #') to apply network settings. To continue configuration, see $newip"
                Invoke-CrestronCommand $dev.'Discovered IP' -Command "reboot" -Secure -Username $usr -Password $pw
            }
    }

    foreach($device in $matcheddevices) 
    {
        Start-Job -Name $device.Hostname -ScriptBlock $setupblock -ArgumentList $device, $targetDevices, $updateActions, $configPath, $progFile, $uiFile, $username, $password > $null
    }
    Get-Job | Receive-Job -Wait
}
else 
{
    Write-Warning -Message "**insert sad trombone noise**`nWe couldn't find any of the stuff you wanted to update, sucks to be you.`nCheck your network settings, make sure the devices are actually ON and try again."
}

Read-Host -Prompt "Press Enter To Exit"