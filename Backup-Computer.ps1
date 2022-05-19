[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
$log = [PSCustomObject]@{
    Warnings = New-Object System.Collections.ArrayList
    Errors = New-Object System.Collections.ArrayList
    runtime = (get-date -Format "yyyy.MM.dd HH:mm:ss")
    backupLocation = ""
    Choices = [System.Collections.Generic.List[object]]::new()
    Passwords = New-Object System.Collections.ArrayList
}
$outputDate = (get-date -Format "yyyy.MM.dd")

if (Test-Path $env:OneDriveCommercial) {    
    $backupLocation = ($env:OneDriveCommercial) + "\Laptop Backup " + ($outputDate)
    $log.backupLocation = $backupLocation
    Write-Host "Saving to " $backupLocation -BackgroundColor DarkGreen
    $null = $log.Warnings.Add("Saving to OneDrive, please ensure user's OneDrive is syncing.")
    
}
else {
    $backupLocation = (New-Item -Path ("$env:USERPROFILE\Laptop Backup " + ($outputDate)) -ItemType Directory -Force).fullName
    $null = $log.Warnings.add("Note unable to save to OneDrive Location. Saving to $backuplocation")
    Write-host "Note unable to save to OneDrive Location. Saving to $backuplocation" -BackgroundColor DarkYellow

    $log.backupLocation = $backupLocation
}

"Chrome","Firefox","Edge","Wifi","Printers"| ForEach-Object {
    $choice = New-Variable -Name "Bool$_" -Value (Read-Host "Would you like to back up $_ ? y / n") -Force -PassThru
    if ($choice.Value -eq "y"){$choice.Value = $true}
    else {$choice.Value = $false}
    $null = $log.Choices.add([PSCustomObject]@{
        $choice.Name = $choice.Value
    })
}

Write-Host "`n=======================================`n" -ForegroundColor Black
Write-Host "Backing up Microsoft AppData" -BackgroundColor DarkGreen
$null = New-Item "$backupLocation\Microsoft" -ItemType Directory -ErrorAction SilentlyContinue
"Signatures","Speech","Stationery","Sticky Notes","Templates" | ForEach-Object {
    if (Test-Path $env:APPDATA\Microsoft\$_){
        $null = New-Item $backupLocation\Microsoft\$_ -Force -ItemType Directory 
        Write-Host "Backing up $_" -BackgroundColor DarkCyan
        Copy-Item $env:APPDATA\Microsoft\$_\* -Destination $backupLocation\Microsoft\$_ -Force
    }
}

#Optional Choices
Write-Host "`n=======================================`n" -ForegroundColor Black
Write-Host "Backing up Optional data selected:" -BackgroundColor DarkGreen

#TODO: add bookmarks.bak
if ($log.Choices.BoolChrome -eq $true) {
    write-host "Backing up Chrome Bookmarks" -BackgroundColor DarkCyan

    $null = New-Item "$backupLocation\Google Chrome" -ItemType Directory -ErrorAction SilentlyContinue
    Write-Host "Backing up default chrome bookmarks. If user has multiple profiles, be sure to export those." -ForegroundColor Yellow
    Copy-Item "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks" -Destination "$backupLocation\Google Chrome"
    Copy-Item "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks.bak" -Destination "$backupLocation\Google Chrome"

    $null = $log.Warnings.Add("Be sure to backup Google saved passwords if not synced to a google profile.")
}

if ($log.Choices.BoolFirefox -eq $true){
    write-host "Backing up Firefox Bookmarks" -BackgroundColor DarkCyan

    $null = New-Item "$backupLocation\FireFox" -ItemType Directory -ErrorAction SilentlyContinue
    $FirefoxDefault = (get-item "$env:APPDATA\Mozilla\Firefox\Profiles\*.default").FullName
    Copy-Item "$FirefoxDefault\places.sqlite" -Destination "$backupLocation\Firefox"
    Copy-Item "$FirefoxDefault\Bookmarkbackups\" -Destination "$backupLocation\Firefox" -Recurse

}

if ($log.Choices.BoolEdge -eq $true){
    write-host "Backing up Edge Bookmarks" -BackgroundColor DarkCyan

    $null = New-Item "$backupLocation\Edge" -ItemType Directory -ErrorAction SilentlyContinue
    $EdgeDefault = "$env:LOCALAPPDATA\microsoft\edge\user data\default"
    Copy-Item "$EdgeDefault\Bookmarks" -Destination "$backupLocation\Edge"

    $null = $log.Warnings.Add("Be sure to backup Edge saved passwords.")
}

if ($log.Choices.BoolWifi -eq $true){
    Write-Host "Backing up Wifi Profiles" -BackgroundColor DarkCyan

    $null = New-Item "$backupLocation\Wifi Profiles" -ItemType Directory -ErrorAction SilentlyContinue
    $null = netsh wlan export profile key=clear folder="$backupLocation\Wifi Profiles"
    $wifiProfileNames = (netsh wlan show profiles) | Select-String "\:(.+)$"
    $null = New-Item -Path "$backupLocation\Wifi Profiles\ManualRestore.txt" -ItemType File
    foreach ($profile in $wifiProfileNames) {
        $name = $profile.Matches.Groups[1].Value.Trim()
        "netsh wlan add profile filename=`"$backupLocation\Wifi Profiles\Wi-Fi-$name.xml`"" | Add-Content "$backupLocation\Wifi Profiles\ManualRestore.txt" 
    }
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
if ($log.Choices.BoolChrome -eq $true) {
    $ChromePass = Read-host "Would you like to back-up Chrome passwords? y/n"
    if ($ChromePass -eq "y") {
        Write-host "Please wait for chrome to finish loading. chrome://settings/passwords"
        Start-process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList  'google.com --profile-directory="Default"'
        Start-Sleep -Seconds 1.5
        [System.Windows.Forms.SendKeys]::SendWait("^l")
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("chrome://settings/passwords")
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")

        $null = $log.Passwords.add("Chrome")
        Pause
    }  
}
if ($log.Choices.BoolEdge -eq $true) {
    $EdgePass = Read-host "Would you like to back-up Edge passwords? y/n"
    if ($EdgePass -eq "y") {
        Write-host "Please wait for Edge to finish loading: edge://settings/passwords"
        Start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList  "bing.com"
        Start-Sleep -Seconds 1.5
        [System.Windows.Forms.SendKeys]::SendWait("^l")
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("edge://settings/passwords")
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")

        $null = $log.Passwords.add("Edge")
        Pause
    }
    
}

if ($log.Choices.BoolPrinters -eq $true){
    #printbrm.exe does not handle spaces in file paths correctly. creating backup in profile, then copying to backup folder.
    $null = New-Item "$backupLocation\PrinterExport" -ItemType Directory -Force -ErrorAction SilentlyContinue
    $null = New-Item "$env:USERPROFILE\Printerexport" -ItemType Directory -Force -ErrorAction SilentlyContinue
    $printerBackupName = "printers_$(get-date -Format "yyyy.MM.dd-HH.mm.ss").printerExport"

    Write-Host "Working on backing up printers." -BackgroundColor DarkCyan
    write-host "This can take a some time depending on how many printers and drivers are installed." -ForegroundColor Yellow
    Write-host "If you would like to delete old or unused printers. Please do so now via PrintManagement.msc before hitting enter." -ForegroundColor Yellow
    Pause
    Write-Host "Starting to back up printers" -ForegroundColor Green
    [void](C:\Windows\System32\Spool\Tools\printbrm.exe -B -f "$env:USERPROFILE\printerexport\$($printerBackupName)")
    "To restore, use printbrmui.exe located in: C:\Windows\System32\PrintBrmUi.exe" | Set-Content $backupLocation\PrinterExport\ReadMe.txt
    Copy-Item "$env:USERPROFILE\printerexport\$($printerBackupName)" -Destination "$backupLocation\PrinterExport" -Force

    $log | Add-Member -MemberType NoteProperty -Name "printerFile" -Value $printerBackupName
}

Write-Host "=======================================" -ForegroundColor Black
Write-host "Warnings:" -ForegroundColor Yellow
$log.Warnings | ForEach-Object {write-host -ForegroundColor Yellow $_}

Write-Host "=======================================" -ForegroundColor Black
Write-Host "Errors:" -ForegroundColor Red
$log.Errors | ForEach-Object {write-host -BackgroundColor Red $_}
Write-Host "=======================================" -ForegroundColor Black

$log | ConvertTo-Json -Depth 4| Set-Content $backupLocation\backup.json
invoke-item $backupLocation
pause