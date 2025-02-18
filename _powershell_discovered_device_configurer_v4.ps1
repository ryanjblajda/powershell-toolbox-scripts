try 
{
    Import-Module PSCrestron
    Add-Type -AssemblyName Microsoft.VisualBasic
}
catch 
{
    Write-Warning "Install Crestron Powershell EDK First You Must!! - Yoda, probably..."
}

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

    Write-Host "Available Devices`n-------------------------------------------"
    Write-Host $discovered | Out-Default
}
catch 
{
    Write-Warning -Message "**insert sad trombone noise** `nLooks like your aren't connected to any networks...`nFix that and run me again!"
    Read-Host "Press Enter To Exit"
    Exit
}
#loop our entire main loop for updating stuff often.
do {
        #do stuff with our list of devices

    do {
        $response = Read-Host -Prompt "Enter the hostnames of the devices you wish to update/configure, separated by a comma"
        $response -replace "`r`n", "" > $null
        $targetdevices = $response.Replace(" ", "").ToUpper().Split(",")
        $hostconfirmation = Read-Host -Prompt "You wish to update/confgure $targetdevices ? --- Enter [y] to confirm, [n] to re-enter desired devices"
        $hostconfirmation.ToLower() > $null
    }
    while($hostconfirmation -ne "y")

    do {
        $defaultcreds = Read-Host -Prompt "Use Crestron default credentials? --- Enter [y] to confirm, [n] to enter custom SSH credentials"
        $defaultcreds.ToLower()
        if($defaultcreds -eq "y") {
            $username = "crestron"
            $password = ""
            $credsresponse = "y"
        }
        elseif($defaultcreds -eq "n")
        {
            do {
                $username = Read-Host -Prompt 'Username'
                $password = Read-Host -Prompt 'Password'
                $credsresponse = Read-Host -Prompt "USER: $username | PW: $password --- Enter [y] to confirm, [n] to re-enter credentials"
                $credsresponse.ToLower() > $null
            }
            while ($credsresponse -ne "y")
        }
        else 
        {
            Write-Warning -Message "Yo...type [y] or [n], not $defaultcreds"
        }       
    }
    while($credsresponse -ne "y")

    $updateconfig = $null
    $updateprogram = $null
    $updatetouchpanel = $null
    $updatehostname = $null

    #ask the end user if they want to update the program on the targeted devices
    do {
        $progresponse = Read-Host -Prompt "Update SIMPL Programming? --- Enter [y] to confirm, [n] to skip"
        $progresponse.ToLower() > $null
        if($progresponse -eq "y") {
            $updateprogram = $true
            do {
                $response = Read-Host -Prompt "Enter the model names of the Control Systems you wish to push SIMPL programming to, separated by a comma"
                $controlsys -replace "`r`n", "" > $null
                $controlsys = $response.Replace(" ", "").ToUpper().Split(",")
                $targetconfirmation = Read-Host -Prompt "You wish to send new programming to $controlsys type devices? --- Enter [y] to confirm, [n] to try again"
                $targetconfirmation.ToLower() > $null
            }
            while($targetconfirmation -ne "y")
            
            Write-Host "Use the file dialog to select a SIMPL program file"
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
                    $controlsys = $response.Replace(" ", "").ToUpper().Split(",")
                    $targetconfirmation = Read-Host -Prompt "You wish to send new config files to $controlsys type devices? --- Enter [y] to confirm, [n] to try again"
                    $targetconfirmation.ToLower() > $null
                }
                while($targetconfirmation -ne "y")
            }
            Write-Host "Config files WILL be updated on $controlsys`nNavigate to the Folder Browser; Select Config File Folder"
            Write-Warning "Dont forget to name config files according to device hostname, in ALL UPPERCASE!!!"
            
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
                $interfaces = $response.Replace(" ", "").ToUpper().Split(",")
                $targetconfirmation = Read-Host -Prompt "You wish to send a new UI file to $interfaces type devices? --- Enter [y] to confirm, [n] to try again"
                $targetconfirmation.ToLower() > $null
            }
            while($targetconfirmation -ne "y")
            
            Write-Host "Use the file dialog to select a UI file"
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

    #notify the user about what will be done
    if($updateconfig){
        Write-Output "$controlsys will be sent config files from $configPath"
    }
    if($updateprogram) {
        Write-Output "$controlsys devices will receive $progFile"
    }
    if($updatetouchpanel) {
        Write-Output "$interfaces devices will receive $uiFile"

        do
        {
            $iptablesend =  Read-Host "Do you want to send an IP Table Entry? --- Enter [y] to confirm, [n] to continue without" 
            $iptablesend = $iptablesend.ToLower()

            if($iptablesend -eq "y") { $updateiptable = $true }
            else { $updateiptable = $false }
        }
        while ($iptablesend -ne "y" -and $iptablesend -ne "n")
    }

    $actions = ($updateprogram, $updateconfig, $updatetouchpanel, $updateiptable)
    $devices = ($controlsys, $interfaces)

    $setupblock = {
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [PSObject]$dev,
            [PSObject]$targets,
            [bool[]]$todo,
            [string]$cfg,
            [string]$prog,
            [string]$ui,
            [string]$ipt,
            [string]$usr, 
            [string]$pw)

            Write-Host -ForegroundColor Cyan "Starting setup for $($dev.Hostname) : $($dev.IP) | USR: $usr / PW: $pw"

            $m = $dev.Description.Split()
            $model = $m[0]
        
            
            if($todo[0] -eq $true) { #if the end user wants to update programming
                if($targets[0] -contains $model) #if the device description contains the device we are looking to target, allows us to target multiple versions of devices.
                {
                    Write-Host -ForegroundColor Cyan "Starting program update on $($dev.Hostname) : $($dev.IP)"
                    Send-CrestronProgram $dev.IP -LocalFile $prog -SendSIGFile -Secure -Username $usr -Password $pw -ShowProgress
                }
            }

            if($todo[1] -eq $true) { #if the end user wants to send new config file
                if($targets[0] -contains $model) { #this will make sure that only the right devices get configs...not everything.
                    if([IO.File]::Exists("$cfg\$($dev.Hostname).json") -ne $false)
                    {
                        Write-Host -ForegroundColor Cyan "Starting config file update on $($dev.Hostname) : $($dev.IP)"
                        Send-FTPFile $dev.IP -LocalFile "$cfg\$($dev.Hostname).json" -RemoteFile '\USER\config.json' -Secure -Username $usr -Password $pw
                    Invoke-CrestronCommand $dev.IP -Command "progreset -p:1" -Secure -Username $usr -Password $pw
                }
                else 
                {
                    Write-Warning "$($dev.Hostname).json DOES NOT EXIST!! Config file not uploaded, please create one and try again..."   
                }
            }
        }

        if($todo[2] -eq $true) { #if the end user wants to update UI files
            if($targets[1] -contains $model) { #if device is a match to the device we want to send UI files to
               if($todo[3] -eq $true) 
                { 
                    Write-Host -ForegroundColor Cyan "Setting IP Table on $($dev.Hostname) : $($dev.IP) | $ipt"
                    Invoke-CrestronCommand $dev.IP -Command "addm $ipt" -Secure -Username $usr -Password $pw
                }
                
                Write-Host -ForegroundColor Cyan "Starting UI update on $($dev.Hostname) : $($dev.IP)"
                Send-CrestronProject $dev.IP -LocalFile $ui -Secure -Username $usr -Password $pw
            }
        }

        Write-Host -ForegroundColor Cyan "Completing setup for $($dev.Hostname) | $($dev.IP)"
    }

    $match = New-Object System.Collections.ArrayList

    foreach($dev in $discovered) {
        if ($targetdevices -match $dev.Hostname.ToUpper())
        {
            $model = $dev.Description.Split()
            #Write-Host $model[0]
            if ($null -ne $controlsys)
            {
                if ($controlsys -match $model[0])
                {
                    $match.Add($dev) > $null
                }
            }

            if ($null -ne $interfaces)
            {
                if ($interfaces -match $model[0])
                {
                    if (!$match.Contains($dev))
                    {
                        $match.Add($dev) > $null
                    }
                }
            }
        }
    }

    $iptcmds = @{}

    foreach($device in $match) {
        if ($updateiptable)
        {
            Write-Host "You specified you wish to set the IP Table on devices, we will now generate those IP Table commands..."

            $uimodels = ("TSW", "TS", "TST", "DGE", "TPMC")
            
            $iptable = $null

            if ($updateiptable)
            {
                $foundmatch = $false

                foreach($type in $uimodels)
                {   
                    #Write-Host $device.Description.Split()[0]
                    if($device.Description.Split()[0].Contains($type))
                    {
                        $foundmatch = $true

                        do
                        {
                            $iptable = Read-Host "Please Provide the Master IP Table Entry for $($device.Hostname), formatted as <ID> <XXX.XXX.XXX.XXX> <RM>"
                            $iptable.ToLower() > $null

                            $iptconfirmation = Read-Host "You would like to create an IP Table Entry pointing $($device.Hostname) @ $($device.IP) to < $iptable >, enter [y] to confirm, or [n] to correct your entry"
                            $iptconfirmation.ToLower() > $null
                            #assign the command
                            $iptcmds[$device] = $iptable

                        }
                        while($iptconfirmation -ne "y")
                    }
                    if ($foundmatch -eq $true) { break }
                }
            }
        }
    }

    $runConfig = {
        Write-Host -ForegroundColor Cyan "There are $($match.Count) devices to update, starting update process"

        foreach($device in $match) {
            $ipcmd = $iptcmds[$device]
            Start-Job -Name $device.Hostname -ScriptBlock $setupblock -ArgumentList $device, $devices, $actions, $configPath, $progFile, $uiFile, $ipcmd, $username, $password > $null
            #print output
            Get-Job | Receive-Job -Wait
        }
    }

    Invoke-Command $runConfig

    #allow us to quickly just re-do exactly what we just did for quick troubleshooting updates
    do {
        $exit = Read-Host -Prompt "All Jobs Completed, Enter [e] to Exit the script, [n] to re-run through the discovery & configuration process anew, [r] to re-do what you just did immediately."
        $exit = $exit.ToLower()
        $exit -replace "`r`n", "" > $null
        if ($exit -eq 'r') {
            Invoke-Command $runConfig
        }
    } while(($exit -ne 'n') -and ($exit -ne 'e'))

} while($exit -ne 'e')

exit