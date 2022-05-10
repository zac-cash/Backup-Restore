[void][System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
Add-Type -AssemblyName System.Windows.Forms
Clear-Host
Write-host "This Script will overwrite current files with those in the backup folder.`nOverwritten files will not be saved or recoverable. Would you like to continue?" -ForegroundColor DarkYellow
do {$exit = read-host "y/n"} while (($exit -ne "y") -and ($exit -ne  "n"))
if ($exit -eq "n"){exit}

Write-Host "Please select the backup.json in onedrive, created by the backup script." -ForegroundColor Green
Pause

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('UserProfile') }
$null = $FileBrowser.ShowDialog()

$backupjson = Get-Content -Path ($FileBrowser.FileName) | ConvertFrom-Json
if (Test-Path $backupjson.backupLocation){
    Set-Location $backupjson.backupLocation
    $backupLocation = $backupjson.backupLocation
}else{
    Write-host "Unable to find backup location listed in JSON file." -ForegroundColor Red
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        RootFolder            = "UserProfile"
        Description           = "$Env:ComputerName - Select a folder"
    }
    $null = $FolderBrowser.ShowDialog()
    if ($FolderBrowser.SelectedPath -eq ""){
        Write-host "Location not found. Exiting" -ForegroundColor Red
        exit
    }
}

Write-Host "Attempting to restore Appdata" -ForegroundColor Cyan
if (test-path $backupLocation\Microsoft\){
    try {
        Copy-Item $backupLocation\Microsoft\* -Destination $env:APPDATA\Microsoft\ -Recurse -Force -ErrorAction Stop
    } catch {
        Write-host "unable to copy all files due to Office365 apps being open. Please close all but OneDrive and hit enter." -BackgroundColor Red
        Pause
        Copy-Item $backupLocation\Microsoft\* -Destination $env:APPDATA\Microsoft\ -Recurse -Force
    }
}

$ChromeDefault = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\"
$ChromeInstalled = test-path $ChromeDefault
if ($backupjson.Choices.BoolChrome -eq $true -and $ChromeInstalled -eq $true) {
    Write-Host "Restoring Chrome." -ForegroundColor Cyan
    Copy-Item "$backupLocation\Google Chrome\Bookmarks" -Destination  $ChromeDefault -Force
}

$FirefoxDefault = (get-item "$env:APPDATA\Mozilla\Firefox\Profiles\*.default" -ErrorAction SilentlyContinue).FullName
$firefoxInstalled = test-path "$env:APPDATA\Mozilla\Firefox\Profiles\" -ErrorAction SilentlyContinue
if (!$firefoxInstalled){
    Write-Host "Firefox does not appear to be installed." -ForegroundColor Cyan
}
if ($backupjson.Choices.BoolFirefox -eq $true -and $firefoxInstalled -eq $true) {
    Write-Host "Restoring Firefox"
    Copy-Item "$backupLocation\FireFox\places.sqlite" -Destination "$FirefoxDefault"
    Copy-Item "$backupLocation\FireFox\Bookmarkbackups\*" -Destination "$FirefoxDefault\Bookmarkbackups\" -Recurse -Force
}

if ($backupjson.Choices.BoolEdge -eq $true){
    write-host "Restoring Edge" -ForegroundColor Cyan
    Copy-Item "$backupLocation\Edge\Bookmarks" -Destination "$env:LOCALAPPDATA\microsoft\edge\user data\default"
}

if ($backupjson.Choices.BoolWifi -eq $true) {
    Write-Host "Restoring Wifi" -ForegroundColor Cyan
    $wifiDecision = write-host "Would you like to delete the wifi profiles after importing? They contain clear text passwords." -ForegroundColor Yellow
    Get-ChildItem -Path "$backupLocation\Wifi Profiles" -Exclude "*.txt" | ForEach-Object {netsh wlan add profile filename="$_"}

    if ($wifiDecision -eq "y") {
        Get-ChildItem -Path "$backupLocation\Wifi Profiles" -Exclude "*.txt" | Remove-Item -Force
    }
}

if ($backupjson.Choices.BoolPrinters -eq $true) {
    Write-Host "Restoring Printers. Please enter administrator credentials for printer installation."
    if (test-path $backupLocation\PrinterExport\*.printerExport){
        $PrinterFile = (get-item $backupLocation\PrinterExport\*.printerExport)[0].Name
        New-item -ItemType Directory "$env:USERPROFILE\printerRestore" -ErrorAction SilentlyContinue
        Copy-Item -Path "$backuplocation\PrinterExport\$Printerfile" -Destination "$env:USERPROFILE\printerRestore\"

        $path = "$env:USERPROFILE\printerRestore\$Printerfile"
        Start-Process cmd -ArgumentList "/k C:\Windows\System32\spool\tools\PrintBrm.exe -r -f $path" -Credential (Get-Credential)
    }
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
if ($backupjson.Choices.BoolChrome -eq $true) {
    $ChromePass = Read-host "Would you like to restore Chrome passwords? Selecting yes will attempt to open Chrome with the import flag checked. y/n"
    if ($ChromePass -eq "y") {
        Write-host "Please wait for chrome to finish loading. chrome://settings/passwords"
        try{
            Start-process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList  'google.com --profile-directory="Default" -enable-features=PasswordImport'
        }catch{
            write-host "unable to find in x86"
            Start-process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList  'google.com --profile-directory="Default" -enable-features=PasswordImport'
        }
        Start-Sleep -Seconds 2
        [System.Windows.Forms.SendKeys]::SendWait("^l")
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("chrome://settings/passwords")
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")

        Pause
    }  
}

if ($backupjson.Choices.BoolEdge -eq $true) {
    $EdgePass = Read-host "Would you like to restore Edge passwords? This will attempt to open Edge directly to the passwords screen. y/n"
    if ($EdgePass -eq "y") {
        Write-host "Please wait for Edge to finish loading: edge://settings/passwords"
        Start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList  "bing.com"

        Start-Sleep -Seconds 2
        [System.Windows.Forms.SendKeys]::SendWait("^l")
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("edge://settings/passwords")
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")
    }

    Write-Host "finished" -ForegroundColor Green
    Pause
}
