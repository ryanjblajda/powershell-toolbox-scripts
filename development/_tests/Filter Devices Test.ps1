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
echo "The Following Worksheets Are Available To Import:"
echo "------------------------------------------------"
echo $sheets.Name

do
{
    $worksheet = Read-Host -Prompt "Specify The Worksheet You Wish To Import"
    $confirmation = Read-Host -Prompt "You entered: ""$worksheet"" is this correct? --- Enter [y] to confirm, [n] to re-type the worksheet name"
    $confirmation.ToLower()
}
while($confirmation -ne "y")


do
{
    $bldgs = Read-Host -Prompt "Enter the Building(s) you wish to target, seperated by a comma"
    $bldgs -replace "`r`n", ""
    $bldgarray = $bldgs.Split(',').Trim()
    $confirmation = Read-Host -Prompt "You entered, $bldgarray is this correct? --- Enter [y] to confirm, [n] to re-enter a list of buildings" 
}
while($confirmation -ne "y")

do {
    $rooms = Read-Host -Prompt "Enter Room Numbers you wish to target, seperated by a comma --- enter [all] to target all rooms"
    $bldgs -replace "`r`n", ""
    $rmarray = $rooms.Split(',').Trim()
    $confirmation = Read-Host -Prompt "You entered, $rmarray is this correct? --- Enter [y] to confirm, [n] to re-enter a list of rooms"
    $confirmation.ToLower()
}
while($confirmation -ne "y")

<#>
$bldgarray = ("Boyde", "Rounds Hall")
$rmarray = ("103", "005")
#>

$worksheetdevices = ImportExcel\Import-Excel $excel -Worksheet $worksheet -HeaderRow 11 -HeaderName "Room Type", "Building", "Room #", "Manufacturer", "Model", "Serial #", "Location", "PoE", "MAC Address", "Hostname", "FW Ver.", "DHCP/Static", "VLAN", "Jack ID", "SW/Port", "IP Address", "Subnet Mask", "Gateway", "Notes" | Where-Object -Property "Manufacturer" -eq "Crestron"

$bldgdevices = $worksheetdevices | Where-Object { $bldgarray -contains $_.Building }
$rmdevices = $bldgdevices | Where-Object { $_.'Room #' -in $rmarray }
echo $rmdevices