# Backup-Restore

## How to use:
Open Powershell under the user's context and enter the following:

### Backup
```
(Invoke-WebRequest https://raw.githubusercontent.com/zac-cash/Backup-Restore/main/Restore-backup.ps1 ).content | Invoke-Expression
```

### Restore
```
(Invoke-WebRequest https://raw.githubusercontent.com/zac-cash/Backup-Restore/main/Restore-backup.ps1 -UseBasicParsing).content | Invoke-Expression
```
**Note admin credentials will be prompted for in the middle of the script for restoring of printers.**
