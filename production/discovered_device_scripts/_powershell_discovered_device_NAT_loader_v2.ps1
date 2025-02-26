try 
{
    Import-Module PSCrestron
    Import-Module ImportExcel
}
catch 
{
    Write-Warning "Install Crestron Powershell EDK & ImportExcel Module First You Must!! - Yoda, probably..."
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

$SetRule = {
    param(
        [string]$dev,
        [psobject]$rule, 
        [string]$usr,
        [string]$pw
    )

    #Write-Information "Setting NAT/Port Fowarding Rule`nInternal IP Address: $($Rule.'IP Address') // Internal Port: $($Rule.'Internal Port [1-65535]') // $($Rule.'External Port [1-65535]') // $($Rule.'Transport Type')"

    $cmdType = $Rule.'Action'.ToLower()

    $cmd = "$($cmdType)portmap $($Rule.'External Port [1-65535]') $($Rule.'Internal Port [1-65535]') $($Rule.'IP Address') $($Rule.'Transport Type')"
    Write-Host "Invoking Command on Device @ $($dev) // $cmd"
    Invoke-CrestronCommand -Device $dev -Command $cmd -Username $usr -Password $pw -Secure
}

#loop our entire main loop for updating stuff often.
do {
    #discover devices and handle if the computer discovers nothing
    $discoveredCrestronDevices = Get-Devices
    #display the output from Get-Devices
    Write-Output $discoveredCrestronDevices | Out-Default

    do {
        $response = Read-Host -Prompt "Enter the hostname of the device you wish to update/configure"
        $formattedResponse = Get-Flattened -ToFlatten $response
        $desiredDevice = $formattedResponse.ToUpper().trim()

        #$desiredDeviceIPAddress = $discoveredCrestronDevices | Where-Object {$_.'Hostname' -eq $desiredDevice }

        $desiredDeviceIPAddress = 'UNKNOWN'

        ForEach ($device in $discoveredCrestronDevices) {
            if($device.'Hostname' -eq $desiredDevice) { 
                $desiredDeviceIPAddress = $device.'IP'
            }
        }

        Write-Host "The hostname of the device you wish to configure is: $desiredDevice @ $desiredDeviceIPAddress" -ForegroundColor DarkGreen
        $desiredDeviceCorrect = Read-Host -Prompt "Enter [y] to confirm, [n] to re-enter desired device"
    }
    while($desiredDeviceCorrect -ne "y")

    #Write-Host $desiredDeviceIPAddress

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

    do {
        #ask the end user if they want to update the program on the targeted devices
        do {
            Write-Host "Use this file dialog to select the Excel document we will extract NAT Port Fowarding Rules from"
            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $dialog.AddExtension = $true
            $dialog.Filter = 'Excel Workbook | *.xlsx; *.xls'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = "$HOME\Downloads"
            $dialog.RestoreDirectory = $true
            $dialog.Title = 'Select Excel File'
            $dialog.ShowDialog()
            $excelDocument = $dialog.FileName
            
            Write-Host "You wish to use the contents of: $excelDocument" -ForegroundColor DarkGreen
            $fileCorrect = Read-Host "Type [y] to confirm, or [n] to select a different file"
            $fileCorrect = Get-Flattened $fileCorrect
        }
        while($fileCorrect -ne "y")


        $documentData = Import-Excel $excelDocument 
        $documentData = $documentData | Where-Object { $_.'Transport Type' -eq "TCP" -or  $_.'Transport Type' -eq "UDP" }

        Write-Host "The following NAT/Port Forwarding Routes will be created." -ForegroundColor DarkGreen
        $documentData | Out-Default

        $fileContentsCorrect = Read-Host "Type [y] to confirm, or [n] to re-load an excel file"
        $fileContentsCorrect = Get-Flattened $fileContentsCorrect
    }
    while($fileContentsCorrect -ne "y")

    #for each nat rule, start a job
    $actions = {
        ForEach ($rule in $documentData) {
            Start-Job -ScriptBlock $SetRule -ArgumentList $desiredDeviceIPAddress, $rule, $username, $password > $null
        }

        Get-Job | Receive-Job -Wait
    }

    Invoke-Command $actions

    #allow us to quickly just re-do exactly what we just did for quick troubleshooting updates
    do {
        $exit = Read-Host -Prompt "All Jobs Completed, Enter [e] to Exit the script, [n] to re-run through the discovery & configuration process anew, [r] to re-do what you just did immediately."
        $exit = $exit.ToLower()
        $exit -replace "`r`n", "" > $null
        if ($exit -eq 'r') {
            Invoke-Command $actions
        }
    } while(($exit -ne 'n') -and ($exit -ne 'e'))

} while($exit -ne 'e')

exit