Import-Module ImportExcel
Import-Module PSCrestron
Add-Type -AssemblyName Microsoft.VisualBasic

try 
{
    Write-Output "Discovering Crestron Devices...."
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
        Write-Output "We Discovered Crestron Devices, Please Find The File Browser Window, and select an IP Worksheet"
    }
}
catch 
{
    Write-Warning -Message "**insert sad trombone noise** `nLooks like your aren't connected to any networks...`nFix that and run me again!"
    Read-Host "Press Enter To Exit"
    Exit
}

#have the user select an excel file to import
do
{
    $dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $dialog.AddExtension = $true
    $dialog.Filter = 'Excel Files | *.xlsx; *.csv'
    $dialog.Multiselect = $false
    $dialog.InitialDirectory = "$HOME\OneDrive"
    $dialog.RestoreDirectory = $true
    $dialog.Title = 'Select IP Worksheet'
    $dialog.ShowDialog()
}
while(!$dialog.FileName)

$excel = $dialog.FileName
#have the user select a worksheet from the excel file to import
$sheets = Get-ExcelSheetInfo $excel
Write-Output "The Following Worksheets Are Available To Import:`n------------------------------------------------"
Write-Output $sheets.Name

do
{
    $worksheet = Read-Host -Prompt "Specify The Worksheet You Wish To Import"
    $confirmation = Read-Host -Prompt "You entered: ""$worksheet"" is this correct? --- Enter [y] to confirm, [n] to re-type the worksheet name"
    $confirmation.ToLower() > $null
}
while($confirmation -ne "y")

$targetdevices = ImportExcel\Import-Excel $excel -Worksheet $worksheet -HeaderRow 11 -HeaderName "Room Type", "Building", "Room #", "Manufacturer", "Model", "Serial #", "Location", "PoE", "MAC Address", "Hostname", "FW Ver.", "DHCP/Static", "VLAN", "Jack ID", "SW/Port", "IP Address", "Subnet Mask", "Gateway", "Notes" | Where-Object -Property "Manufacturer" -eq "Crestron"

echo "Buildings:`n------------------------------------------------"
echo $targetdevices.Building | Sort-Object | Get-Unique

do
{
    $bldgs = Read-Host -Prompt "Enter the Building(s) you wish to target, seperated by a comma --- enter [all] to target all buildings"
    $bldgs -replace "`r`n", "" > $null
    $bldgarray = $bldgs.Split(',').Trim()
    $confirmation = Read-Host -Prompt "You entered, $bldgarray is this correct? --- Enter [y] to confirm, [n] to re-enter a list of buildings" 
}
while($confirmation -ne "y")

echo "Rooms:`n------------------------------------------------"
echo $targetdevices.'Room #' | Sort-Object | Get-Unique

do {
    $rooms = Read-Host -Prompt "Enter Room Numbers you wish to target, seperated by a comma --- enter [all] to target all rooms"
    $rooms -replace "`r`n", "" > $null
    $rmarray = $rooms.Split(',').Trim()
    $confirmation = Read-Host -Prompt "You entered, $rmarray is this correct? --- Enter [y] to confirm, [n] to re-enter a list of rooms"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

if($bldgarray -ne "all")
{
    try { $targetdevices.Building.ToLower() }
    catch{}
    $targetdevices = $targetdevices | Where-Object { $_.Building -contains $bldgarray }
}

if($rmarray -ne "all")
{
    #echo "Targeting Only $rmarray"
    try { $targetdevices.'Room #'.ToLower() }
    catch { }
    $targetdevices = $targetdevices | Where-Object { $_.'Room #' -in $rmarray}
}

#trim all whitespace from serial # fields so comparison actually works
$targetdevices | ForEach-Object { try { $_.'Serial #' = $_.'Serial #'.replace(' ','') } catch { } }

#echo $targetdevices

#convert the TSID shown in the -ver command into the right serial number, and add it to the discovered devices to be compared later. 
foreach($device in $discovered)
{
    $serialhex = $device.Description.Substring($device.Description.IndexOf("#") + 1, $device.Description.IndexOf("]") - $device.Description.IndexOf("#") - 1)
    $serial = Convert-TsidToSerial -TSID $serialhex
    $device | Add-Member -MemberType NoteProperty -Name "Serial #" -Value $serial
}

#compare the available online devices with the targeted devices from our IP worksheet
$matcheddevices = Compare-Object -ReferenceObject $targetdevices -DifferenceObject $discovered -Property "Serial #" -ExcludeDifferent -IncludeEqual -PassThru 
#add the current IP address to our list of matched devices so we can use that to connect in the initial setup. 
$matcheddevices | ForEach-Object { try { $_ | Add-Member -Force -MemberType NoteProperty -Name "Discovered IP" -Value ($discovered | Where-Object -Property "Serial #" -eq $_.'Serial #').'IP'} catch { echo "Error Adding Discovered IP"} } 

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
            $confirmation.ToLower()
        }
        while ($confirmation -ne "y")
    }
    else 
    {
        Write-Warning -Message "Yo...type [y] or [n], not $defaultcreds"
    }       
}
while($confirmation -ne "y")

do
{
    $targets = Read-Host -Prompt 'Enter the device model you wish to target with this firmware update'
    $targets -replace "`r`n", ""
    $confirmation = Read-Host -Prompt "You want to target $targets devices -- Enter [y] to continue, [n] to re-enter target devices"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

do
{
    $updatecmd = Read-Host -Prompt "What is the firmware update command on these devices?? [""PUSHUPDATE FULL"" for DM products] [""PUF"" for anything using a .PUF file]"
    $updatecmd -replace "`r`n", ""
    $confirmation = Read-Host -Prompt "The firmware update command is $updatecmd for $target devices? -- Enter [y] to continue, [n] to re-enter update command"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

$devices = $matcheddevices | Where-Object { $_.Model -match $targets}

Write-Output $devices

$dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
$dialog.AddExtension = $true
$dialog.Filter = 'All Files |*.*'
$dialog.Multiselect = $false
$dialog.InitialDirectory = "$HOME\Downloads"
$dialog.RestoreDirectory = $true
$dialog.Title = 'Select Firmware File'
$dialog.ShowDialog()

$firmwareFilePath = $dialog.FileName

$updateScript = 
{
    Param (
    [PSObject]$dev,
    [string]$upd,
    [string]$fw, 
    [string]$usr, 
    [string]$pw)
    
    $firmwarefile = $fw | Split-Path -Leaf

    Write-Output "Beginning Firmware Update: $firmwarefile to: $($dev.Model) @ $($dev.'Discovered IP')"

    Send-FTPFile -Device $dev.'Discovered IP' -LocalFile $fw -RemoteFile "\firmware\$firmwarefile" -Secure -Username $usr -Password $pw
    Invoke-CrestronCommand -Device $dev.'Discovered IP' -Command $upd -Secure -Username $usr -Password $pw
}

foreach ($device in $devices) 
{
    Start-Job -Name $device.Hostname -ScriptBlock $updateScript -ArgumentList $device, $updatecmd, $firmwareFilePath, $username, $password > $null
}

Get-Job |  Receive-Job -Wait
$response = Read-Host "Press Enter to Exit"