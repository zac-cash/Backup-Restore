# Backup-Restore

##How to use:
### Backup
* (Invoke-WebRequest https://raw.githubusercontent.com/zac-cash/Backup-Restore/main/Restore-backup.ps1).content | Invoke-Expression

### Restore
* (Invoke-WebRequest https://raw.githubusercontent.com/zac-cash/Backup-Restore/main/Restore-backup.ps1 -UseBasicParsing).content | Invoke-Expression
