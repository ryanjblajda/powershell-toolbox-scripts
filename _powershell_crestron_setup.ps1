Write-Host "Beginning ImportExcel Install"

try {
    Install-Module ImportExcel -AllowClobber -Scope CurrentUser
}
catch {
    Write-Error $_
}

Write-Host "Begining Download Of Powershell EDK"

try {
    $url = 'https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.7.5.exe'
    $path_to_file = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
    $powershell_edk = "$($path_to_file)\powershell_edk.exe"
    Invoke-WebRequest $url -OutFile $powershell_edk
}
catch {
    Write-Error $_
}

try {
    Write-Host "Attempting To Start Powershell EDK Installer"
    Invoke-Expression $powershell_edk
}
catch {
    Write-Error $_
}

Read-Host "Script Complete, Press Any Key To Exit, Follow Instructions On Powershell EDK Installer"
