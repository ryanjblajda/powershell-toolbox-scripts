Import-Module PSCrestron

try 
{
    Write-Host "Discovering Crestron Devices...."
    $discovered = Get-AutoDiscovery

    if($null -eq $discovered)
    {
        Write-Warning -Message "No Crestron devices were found on the network!"
        Write-Warning -Message "Please resolve this, and re-run the script"
        Read-Host "Press Enter To Exit"
        Exit
    }

    Write-Host "Available Devices`n-------------------------------------------"
    Write-Host $discovered | Out-Default Format-Table
}
catch 
{
    Write-Warning -Message "**insert sad trombone noise** `nLooks like your aren't connected to any networks...`nFix that and run me again!"
    Read-Host "Press Enter To Exit"
    Exit
}

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
            $username -replace "`r`n", "" > $null
            $password = Read-Host -Prompt 'Password'
            $password -replace "`r`n", "" > $null
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

do {
    $response = Read-Host -Prompt "Enter the model names of the devices you wish to update/configure, separated by a comma"
    $response -replace "`r`n", "" > $null
    $targetdevices = $response.Split(',').Trim().ToUpper()
    $targetconfirmed = Read-Host -Prompt "You wish to update/confgure $targetdevices ? --- Enter [y] to confirm, [n] to re-enter desired devices"
    $targetconfirmed.ToLower() > $null
}
while($targetconfirmed -ne "y")


do {
    $devFirmware = @{}
    #ask the end user if they want to update the firmware on the targeted devices
    foreach($device in $targetdevices)
    {
        do {
            Write-Host -ForegroundColor Yellow "Select the firmware file you wish to send to $device devices"

            $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $dialog.AddExtension = $true
            $dialog.Filter = 'Firmware Files | *.zip; *.puf'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = "$HOME\Downloads"
            $dialog.RestoreDirectory = $true
            $dialog.Title = 'Select Firmware File'
            $dialog.ShowDialog()
            $firmware = $dialog.FileName
    
            Write-Host -ForegroundColor Yellow "You wish to send $firmware to $device devices? --- Enter [y] to confirm, [n] to select a different firmware package"
            $confirm = Read-Host
            $confirm.ToLower() > $null
        }
        while($confirm -ne "y")

        $devFirmware.Add($device, $dialog.FileName)
    }


    Write-Host -ForegroundColor Yellow "The following firmware will be sent to the devices you selected"
    echo $devFirmware
    Write-Host -ForegroundColor Yellow "Do you wish to begin the update process? Once this begins, it should not be stopped. --- Enter [y] to confirm, [n] to begin firmware selection again"
    $batchconfirm = Read-Host
    $batchconfirm.ToLower() > $null
}
while($batchconfirm -ne "y")


$updateDevice = 
{
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$dev,
        [string]$model,
        [string]$firmware,
        [string]$usr, 
        [string]$pw)

        $firmwarefile = $firmware | Split-Path -Leaf

        Write-Host -ForegroundColor Cyan "Uploading $firmwarefile to: $model @ $($dev.IP)"
    
        Send-FTPFile -Device $dev.IP -LocalFile $firmware -RemoteFile "\firmware\$firmwarefile" -Secure -Username $usr -Password $pw

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

foreach($device in $discovered) 
{
    $devDesc = $device.Description.Split()
    
    if($devFirmware.ContainsKey($devDesc[0])) {
        Start-Job -Name $device.Hostname -ScriptBlock $updateDevice -ArgumentList $device, $devDesc[0], $devFirmware[$devDesc[0]], $username, $password > $null
    }
}

Get-Job | Receive-Job -Wait
Read-Host -Prompt "Press Enter To Exit"