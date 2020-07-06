# Name: BackupScript.ps1
Version: 2.0
Forked: 2020-07-01
Forked by: [Ryan Dasso](https://github.com/vvildcard)
Forked from: v1.5
LastModified: 2020-07-02
Creator: [Michael Seidl](https://github.com/Seidlm) aka [Techguy](http://www.techguy.at/)

# Description: 
Copies one or more BackupDirs to the Destination.

# Features
A Progress Bar shows the status of copied MB to the total MB.
Email reports. 
Zipped or unzipped backups. 
Removes oldest backups (see Versions parameter)

# Usage:
    BackupScript.ps1 -BackupDirs "C:\path\to\backup", "C:\another\path\" -Destination "C:\path\to\put\the\backup"
## Required Parameters: 
**BackupDirs** (Default: none)
**Destination** (Default: none)
## Optional Parameters: 
**ExcludeDirs** (Default: "$env:SystemDrive\Users\.*\AppData\Local", "$env:SystemDrive\Users\.*\AppData\LocalLow")
**TempDir** (Default: "$env:TEMP\BackupScript")
**LogPath** (Default: "$TempDir\Logging")
**LogFileName** (Default: "Log")
**LoggingLevel** (Default: "3")
**Zip** (Default: $True)
**Use7ZIP** (Default: $False
**7zPath** (Default: "$env:ProgramFiles\7-Zip\7z.exe")
**Versions** (Default: "2") *Bug*: It will actually keep X+1 backups
**UseStaging** (Default: $True)
**StagingDir** (Default: "$TempDir\Staging")
**ClearStaging** (Default: $True)
**SendEmail** (Default: $False)
**EmailTo** (Default: "test@domain.com")
**EmailFrom** (Default: "from@domain.com")
**EmailSMTP** (Default; "smtp.domain.com")

# Version 2.0 (2020.07.02)
NEW: All configurable variables are parameters. 
FIX: Removed all author/computer-specific paths and messages. 
FIX: Typos
FIX: Correctly handles multiple backups made on the same day (no more randomness)
DIF: Reworked many things for clarity or to be more "PowerShell-like"
DIF: Combined 'staging' and non-staging. "Non-staging" is just staging in the Destination. 
DIF: Added/reworked comments and log output for clarity/consistency. 
 
# Version 1.x
See original: https://github.com/Seidlm/PowerShell-Backup-Script

# Other
PowerShell Self Service Web Portal at www.au2mator.com/PowerShell
