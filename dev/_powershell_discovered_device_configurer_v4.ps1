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
        $password = 'CCS$erv!ce'
        
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