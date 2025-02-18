Import-Module PSCrestron
Add-Type -AssemblyName Microsoft.VisualBasic

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
    $hostname = Read-Host -Prompt 'Use Hostname to Connect? [y]es or [n]o?'
    $hostname.ToLower()
    if($hostname -ne "y" -and $hostname -ne "n")
    {
        Write-Warning -Message "Yo...type [y] or [n], not $defaultcreds"
    }
}
while($hostname -ne "y" -and $hostname -ne "n")

$dialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
$dialog.AddExtension = $true
$dialog.Filter = 'All Files |*.*'
$dialog.Multiselect = $false
$dialog.InitialDirectory = "$HOME\Downloads"
$dialog.RestoreDirectory = $true
$dialog.Title = 'Select Device List'
$dialog.ShowDialog()

do
{
    $targets = Read-Host -Prompt 'Enter the device models you wish to target with this firmware update'
    $targets -replace "`r`n", ""
    $confirmation = Read-Host -Prompt "You want to target $targets devices -- Enter [y] to continue, [n] to re-enter target devices"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

do
{
    $updatecmd = Read-Host -Prompt "What is the firmware update command on these devices?? [""PUSHUPDATE FULL"" for DM products] [""PUF"" for anything using a .PUF file]"
    $updatecmd -replace "`r`n", ""
    $confirmation = Read-Host -Prompt "The firmware update command is $updatecmd for $target devices? -- Enter [y] to continue, [n] to re-enter target devices"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

$devices = Import-Csv $dialog.FileName | Where-Object { $_.Description -match $targets }

echo "Devices"
echo $devices

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
    [string]$usehost,
    [string]$upd,
    [string]$fw, 
    [string]$usr, 
    [string]$pw)
    
    $firmwarefile = $fw | Split-Path -Leaf

    if($usehost = "y")
    {
        $endpoint = $dev.Hostname
    }
    elseif($usehost = "n")
    {
        $endpoint = $dev.IP
    }

    echo "Beginning Firmware Update: $firmwarefile to: $($dev.Description) @ $endpoint"

    Send-FTPFile -Device $endpoint -LocalFile $fw -RemoteFile "\firmware\$firmwarefile" -Secure -Username $usr -Password $pw
    Invoke-CrestronCommand -Device $endpoint -Command $upd -Secure -Username $usr -Password $pw
}

$devices | ForEach-Object { Start-Job -Name $_.Hostname -ScriptBlock $updateScript -ArgumentList $_, $hostname, $updatecmd, $firmwareFilePath, $username, $password}

Get-Job |  Receive-Job -Wait
$response = Read-Host "Press Enter to Exit"