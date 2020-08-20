# Name: BackupScript.ps1
    Version: 2.3
    LastModified: 2020-08-19

# Description
    Copies one or more BackupDirs to the Destination.

# Features
    Zipped or unzipped backups. 
    Optional staging folder for zipping reduces network bandwidth (only copies the final zip to the Destination)
    A Progress Bar shows the status of copied MB to the total MB.
    Email reports. 
    Removes oldest backups (see Versions parameter). 

# Usage
    BackupScript.ps1 -BackupDirs "C:\path\to\backup", "C:\another\path\" -Destination "C:\path\to\put\the\backup"

## Required Parameters
    BackupDirs (Default: none)
    Destination (Default: none)

## Optional Parameters
    ExcludeDirs (Default: "$env:SystemDrive\Users\.*\AppData\Local", "$env:SystemDrive\Users\.*\AppData\LocalLow")
    TempDir (Default: "$env:TEMP\BackupScript")
    LogPath (Default: "$TempDir\Logging")
    LogFileName (Default: "Log")
    LoggingLevel (Default: 2)
    Zip (Default: $True)
    Use7ZIP (Default: $False)
    7zPath (Default: "$env:ProgramFiles\7-Zip\7z.exe")
	$7zCompression (Default: 9, Ultra)
    Versions (Default: "2") *Bug*: It will actually keep X+1 backups
    UseStaging (Default: $True)
    StagingDir (Default: "$TempDir\Staging")
    ClearStaging (Default: $True)
    SendEmail (Default: $False)
    EmailTo (Default: "test@domain.com")
    EmailFrom (Default: "from@domain.com")
    EmailSMTP (Default; "smtp.domain.com")
	
# Version 2.3 (2020-08-19)
	NEW: Implemented 7-Zip encryption option
	FIX: Small e-mail send improvements
	
# Version 2.2 (2020-08-19)
	FIX: Declared $Log variable for the email attachment
	FIX: Don't include parent folder in the archive when using 7-Zip
	CHANGE: Uze .7z extension when using 7-Zip
	NEW: 7-Zip compression level

# Version 2.1 (2020-07-06)
	FIX: ERROR, WARNING and INFO log levels work for console output (the log is always DEBUG level)
	DIF: Adjusted some logging output levels. 
	DIF: Set LogLevel default to 2 (INFO). 

# Version 2.0 (2020-07-02)
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
    Forked: 2020-07-01
    Forked by: [Ryan Dasso](https://github.com/vvildcard)
    Forked from: v1.5
    Creator: [Michael Seidl](https://github.com/Seidlm) aka [Techguy](http://www.techguy.at/)
    PowerShell Self Service Web Portal at www.au2mator.com/PowerShell
