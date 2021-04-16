########################################################
# Name: BackupScript.ps1                              
# Version: 2.4.2
# LastModified: 2021-04-15
# GitHub: https://github.com/vvildcard/PowerShell-Backup-Script
# 
# 
# Description: 
# Copies one or more BackupDirs to the Destination.
# A Progress Bar shows the status of copied MB to the total MB.
# https://raw.githubusercontent.com/vvildcard/PowerShell-Backup-Script/master/README.md
# 
# 
# Usage:
# BackupScript.ps1 -BackupDirs "C:\path\to\backup", "C:\another\path\" -Destination "C:\path\to\put\the\backup"
# Change variables with parameters (ex: Add -NoStaging to turn off staging)
# Change LoggingLevel to 3 an get more output in Powershell Windows.
# 
# 
# Forked: 2020-08-01
# Forked by: [Ryan Dasso](https://github.com/vvildcard)
# Forked from: v1.5
# Creator: [Michael Seidl](https://github.com/Seidlm) aka [Techguy](http://www.techguy.at/)
# PowerShell Self-Service Web Portal at www.au2mator.com/PowerShell
# 
# Feature Requests/Ideas
#   Create txt file index of files in a zip backup.
#   ZIP directly into the archive file. 

### Variables

Param(
    [string]$Versions = "2", # Default number of backups you want to keep. 

    # Source/Dest
    [Parameter(Mandatory=$True)][string[]][Alias("b","s")]$BackupDirs, # Folders you want to backup. Comma-delimited. 
        # To hard-code the source paths, set the above line to: $BackupDirs = @("C:\path\to\backup", "C:\another\path\")
    [Parameter(Mandatory=$True)][string][Alias("d")]$Destination, # Backup to this path. Can be a UNC path (\\server\share)
        # To hard-code the destination, set the above line to: $Destination = "C:\path\to\put\the\backup"
    [string[]]$DefaultExcludedDirs=@("$env:SystemDrive\Users\*\AppData\Local", "$env:SystemDrive\Users\*\AppData\LocalLow", "*\.git\*"), # This list of Directories will not be copied. Comma-delimited. 
    [string[]]$ExcludeDirs=@(""), # Second set of excluded directories for the user to define.

    # Logging
    [string]$TempDir = "$env:TEMP\BackupScript", # Temporary location for logging and zipping. 
    [string]$LogPath = "$TempDir\Logging",
    [string]$LogFileName = "Log", #  Name
    [int]$LoggingLevel = 2, # LoggingLevel only for Output in Powershell Window, 1=smart, 3=Heavy

    # Zip
    [switch]$NoZip, # Don't Zip the backup.
    [switch]$Use7ZIP, # Make sure 7-Zip is installed. (https://7-zip.org)
    [string]$7zPath = "$env:ProgramFiles\7-Zip\7z.exe",

    # Staging -- Ignored when $NoZip is specified. Staging moves all files to a temporary directory before zipping them
    # for performance improvement, especially when copying over a slow link (it's faster to copy a single large zip file
    # than many small files across the network or to a USB device). Staging keeps all the heavy lifting on the local
    # disk, but can also use a lot of space. You need enough space to hold the temporary copy of the backed-up files and
    # the ZIP of those files. 
    [switch]$NoStaging, # Sets $UseStaging. $Destination will be used for Staging.
    [string]$StagingDir = "$TempDir\Staging", # Set your own location for staging. 
    [switch]$NoDeleteStaging, # Don't delete StagingDir after backup. 

    # Email
    [switch]$SendEmail, # $True will send report via email (SMTP send)
    [string]$EmailTo = 'test@domain.com', # List of recipients. For multiple users, use "User01 &lt;user01@example.com&gt;" ,"User02 &lt;user02@example.com&gt;"
    [string]$EmailFrom = 'from@domain.com', # Sender/ReplyTo
    [string]$EmailSMTP = 'smtp.domain.com' # SMTP server address
)


# Parse Excluded directories
function ExcludeCleanUp($Dirs) {  # Clean-up and Convert each Directory to regex
	foreach ($Entry in ($Dirs)) {
        # Remove trailing backslashes and wildcards
        while (($Entry.SubString($Entry.length-1) -eq ("*")) -or ($Entry.SubString($Entry.length-1) -eq ("\"))) {
            $Entry = $Entry.Substring(0,$Entry.Length-1)
        }

        $Entry = $Entry.Replace("*", ".*")  # Convert wildcards to regex style
		$Temp = $Entry.Replace("\", "\\") + "\\.*"  # Convert \ to regex style and add trailing \\.*
		$ExcludeString += $Temp + "|"  # Add to the exclude list regex OR
	}
	$ExcludeString = $ExcludeString.Substring(0, $ExcludeString.Length - 1) # Remove the trailing |
	Return [RegEx]$ExcludeString
}

$AllExcludedDirs = $DefaultExcludedDirs
if ($ExcludeDirs.length -gt 2) {
    $AllExcludedDirs += $ExcludeDirs
}
$Exclude = ExcludeCleanUp($AllExcludedDirs)

# Set the staging directory and name the backup file or folder. 
$BackupName = "Backup-$(Get-Date -format yyyy-MM-dd-hhmmss)"
if ($NoZip) { $Zip = $False } else { $Zip = $True; $ZipFileName = "$($BackupName).zip" }
if ($NoStaging) { $UseStaging = $False } else { $UseStaging = $True }
if ($NoDeleteStaging) { $ClearStaging = $False } else { $ClearStaging = $True }
if ($UseStaging -and $Zip) {
    $DestinationBackupDir = "$StagingDir"
} else {
    $DestinationBackupDir = "$Destination\$BackupName"
}

# Counters
$Items = 0
$Count = 0
$ErrorCount = 0
$BackupStartDate = Get-Date

### FUNCTIONS

# Logging
function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')][string]$Type,
        [string]$Text
    )
	if ($Type -eq 'DEBUG') {
		$TypeLevel = 3
	} elseif ($Type -eq 'INFO') {
		$TypeLevel = 2
	} elseif ($Type -eq 'WARNING') {
		$TypeLevel = 1
	} else {
		$TypeLevel = 0
	}
       
    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogFileName
    $logEntry = '{0}: <{1}> <{2}> {3}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $PID, $Text
    
    try { Add-Content -Path $logFile -Value $logEntry }
    catch {
        Start-sleep -Milliseconds 50
        Add-Content -Path $logFile -Value $logEntry
    }
    if ($LoggingLevel -ge $TypeLevel) { Write-Host $Text }
}


# Create the DestinationBackupDir
Function NewBackupDir {
    if (Test-Path $DestinationBackupDir) {
        Write-au2matorLog -Type DEBUG -Text "Backup/Staging directory already exists: $DestinationBackupDir"
    } else {
        New-Item -Path $DestinationBackupDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Write-au2matorLog -Type INFO -Text "Created backup/staging directory: $DestinationBackupDir"
    }
}

# Delete the oldest Directory backup
Function RemoveDirBackup {
    # Count the previous directory backups and remove the oldest (if needed)
    Write-au2matorLog -Type DEBUG -Text "Checking if there are more than $Versions backup directories in the Destination"
    $BackupDirCount = (Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Directory" -and ($_.Name -like "Backup-????-??-??-??????") }).count

    if ($BackupDirCount -gt $Versions) { 
        Write-au2matorLog -Type INFO -Text "Found $count Directory Backups"
        $RemoveFolder = Get-ChildItem $Destination | Where-Object { ($_.Attributes -eq "Directory") -and ($_.Name -like "Backup-????-??-??-??????") } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
        if ($RemoveFolder) {
            $RemoveFolder.FullName | Remove-Item -Recurse -Force 
            Write-au2matorLog -Type INFO -Text "Deleted oldest directory backup: $RemoveFolder"
        }
    } else {
        Write-au2matorLog -Type DEBUG -Text "Not enough previous backup directories found. Skipping backup deletion."
    }
}

# Delete the oldest Zip backup
Function RemoveZipBackup {
    Write-au2matorLog -Type DEBUG -Text "Checking if there are more than $Versions Zip in the Destination"
    $ZipBackupCount = (Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Archive" -and $_.Name -like "Backup-????-??-??-??????.zip" }).count
    if ($ZipBackupCount -gt $Versions) {
        Write-au2matorLog -Type INFO -Text "Found $ZipBackupCount Zip backups"
        $RemoveZip = Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Archive" -and $_.Name -like "Backup-????-??-??-??????.zip" } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
        if ($RemoveZip) {
            $RemoveZip.FullName | Remove-Item -Recurse -Force 
            Write-au2matorLog -Type INFO -Text "Deleted oldest zip backup: $RemoveZip"
        }
    } else {
        Write-au2matorLog -Type DEBUG -Text "Not enough previous backup zips found. Skipping backup deletion."
    }
}

# Check if DestinationBackupDir and Destination is available
function CheckDir {
    Write-au2matorLog -Type INFO -Text "Checking if DestinationBackupDir and Destination exists"
    if (!(Test-Path $DestinationBackupDir)) {
        Write-au2matorLog -Type ERROR -Text "$DestinationBackupDir does not exist"
        return $False
    }
    if (!(Test-Path $Destination)) {
        Write-au2matorLog -Type ERROR -Text "$Destination does not exist"
        return $False
    } 
    Write-au2matorLog -Type DEBUG -Text "Backup Dirs: $($DestinationBackupDir)"
    Write-au2matorLog -Type DEBUG -Text "Destination: $($Destination)"
    return $True
}


# Save all the Files
Function MakeBackup {
    Write-au2matorLog -Type INFO -Text "Starting the Backup"
    $BackupDirFiles = @{ } # Hash of BackupDir & Files
    $Files = @()
    $SumMB = 0
    $SumItems = 0
    $SumCount = 0
    $colItems = 0
    Write-au2matorLog -Type DEBUG -Text "Counting all files and calculating size"

    # Count the number of files to backup. 
    foreach ($Backup in $BackupDirs) {
        # Get recursive list of files for each Backup Dir once and save in $BackupDirFiles to use later.
        # Optimize performance by getting included folders first, and then only recursing files for those.
        # Use -LiteralPath option to work around known issue with PowerShell FileSystemProvider wildcards.
        # See: https://github.com/PowerShell/PowerShell/issues/6733
		
		write-host "Backup: $Backup"

        if ($Backup) {
            $Files = Get-ChildItem -LiteralPath $Backup -Recurse -Attributes !D+!ReparsePoint -ErrorVariable +errItems -ErrorAction SilentlyContinue | `
				ForEach-Object -Process { `
					Add-Member -InputObject $_ -NotePropertyName "ParentFullName" `
					-NotePropertyValue ($_.FullName.Substring(0, $_.FullName.LastIndexOf("\" + $_.Name))) `
					-PassThru -ErrorAction SilentlyContinue } | `
				Where-Object { $_.FullName -notmatch $Exclude.ToString() } | ` # -and $_.ParentFullName -notmatch $Exclude } | `
				Get-ChildItem -Attributes !D -ErrorVariable +errItems -ErrorAction SilentlyContinue | `
				Where-Object { $_.DirectoryName -notmatch $Exclude.ToString() }
            #$Files = $Files | Where-Object { $_.FullName -notmatch $Exclude.ToString() }
            $BackupDirFiles.Add($Backup, $Files)
            
            $colItems = ($Files | Measure-Object -property length -sum) 
			#write-host "colItems: $colItems"
            #Write-au2matorLog -Type DEBUG -Text "DEBUG"
            #Copy-Item -LiteralPath $Backup -Destination $BackupDir -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch $exclude }
            if ($colItems) {
				$SumMB += $colItems.Sum.ToString()
				$SumItems += $colItems.Count
			}
        }
    }

    # Calculate total size of the backup. 
    $TotalMB = "{0:N2}" -f ($SumMB / 1MB) + " MB"
    Write-au2matorLog -Type INFO -Text "There are $SumItems files ($TotalMB) to copy"

    # Log any errors from above from building the list of files to backup.
    [System.Management.Automation.ErrorRecord]$errItem = $null
    foreach ($errItem in $errItems) {
        Write-au2matorLog -Type ERROR -Text "Skipping '$($errItem.TargetObject)' - Error: $($errItem.CategoryInfo)"
    }
    Remove-Variable errItem
    Remove-Variable errItems

    # Before copy files, remove the oldest previous backup (if needed)
    RemoveDirBackup

    # Copy files to the backup location. 
    Write-au2matorLog -Type WARNING -Text "Copying files to $DestinationBackupDir"
	Write-au2matorLog -Type DEBUG -Text "BackupDirs = $($BackupDirs)"
    foreach ($Backup in $BackupDirs) {
		Write-au2matorLog -Type DEBUG -Text "Backup = $($Backup)"
        $BackupAsObject = Get-Item $Backup  # Convert the backup string path to an object to fix case sensitivity
		Write-au2matorLog -Type DEBUG -Text "BackupAsObject = $($BackupAsObject)"
        $Files = $BackupDirFiles[$Backup]
        Write-au2matorLog -Type DEBUG -Text "Files = $($Files)"

        foreach ($File in $Files) {
            Write-au2matorLog -Type DEBUG -Text "File.FullName = $($File.FullName)"
			if ($BackupAsObject.Parent.FullName -like "?:\") { # This IF handles directories on the root of the tree
				$RelativePath = $File.FullName.Replace("$($BackupAsObject.Parent.Name)", "")
				Write-au2matorLog -Type DEBUG -Text "BackupAsObject Parent = $($BackupAsObject.Parent.FullName)"
			} else {
				$RelativePath = $File.FullName.Replace("$($BackupAsObject.Parent.FullName)\", "")
				Write-au2matorLog -Type DEBUG -Text "BackupAsObject Parent = $($BackupAsObject.Parent.FullName)\"
			}  # RelativePath has the parent folder(s) and file name, starting from the directory being backed up. 
               # Example: "Desktop\myfile.txt"
            Write-au2matorLog -Type DEBUG -Text "RelativePath = $($RelativePath)"
            try {
                # Use New-Item to create the destination directory if it doesn't yet exist. Then copy the file.
                Write-au2matorLog -Type DEBUG -Text "'$($File.FullName)' copied to '$DestinationBackupDir\$RelativePath'"
                New-Item -Path (Split-Path -Path "$DestinationBackupDir\$RelativePath" -Parent) -ItemType "directory" -Force -ErrorAction SilentlyContinue | Out-Null
                Copy-Item -LiteralPath $File.FullName -Destination "$DestinationBackupDir\$RelativePath" -Force -ErrorAction Continue | Out-Null
            }
            catch {
                $ErrorCount++
                Write-au2matorLog -Type ERROR -Text "'$($File.FullName)' returned an error and was not copied"
            }
            $Items += (Get-item -LiteralPath $File.FullName).Length
            $status = "Copy file {0} of {1} and copied {3} MB of {4} MB: {2}" -f $count, $SumItems, $File.Name, ("{0:N2}" -f ($Items / 1MB)).ToString(), ("{0:N2}" -f ($SumMB / 1MB)).ToString()
            $Index = [array]::IndexOf($BackupDirs, $Backup) + 1
            $Text = "Copy data Location {0} of {1}" -f $Index , $BackupDirs.Count
            Write-Progress -Activity $Text $status -PercentComplete ($Items / $SumMB * 100)  
            if ($File.Attributes -ne "Directory") { $count++ }
        }
    }
    $SumCount += $Count
    $SumTotalMB = "{0:N2}" -f ($Items / 1MB) + " MB"
    Write-au2matorLog -Type DEBUG -Text "----------------------"
    Write-au2matorLog -Type DEBUG -Text "Copied $SumCount files. Size of $SumTotalMB"
    if ($ErrorCount) {
		Write-au2matorLog -Type WARNING -Text "$ErrorCount files could not be copied"
	}


    # Send e-mail with reports as attachments
    if ($SendEmail -eq $True) {
        $EmailSubject = "$env:COMPUTERNAME - Backup on $(get-date -format yyyy.MM.dd)"
        $EmailBody = "Backup Script $(get-date -format yyyy.MM.dd) (last Month).`n
                      Computer: $env:COMPUTERNAME`n
                      Source(s): $BackupDirs`n
                      Backup Location: $Destination"
        Write-au2matorLog -Type INFO -Text "Sending e-mail to $EmailTo from $EmailFrom (SMTPServer = $EmailSMTP) "
        # The attachment is $log 
        Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTP -attachment $Log 
    }
}

# Zip Backup
Function ZipBackup {
    $ZipStartDate = Get-Date
    # Count the previous zip backups and remove the oldest (if needed)
    RemoveZipBackup

    Write-au2matorLog -Type INFO -Text "Compressing the Backup Destination"
    if ($Use7ZIP) {
        7ZipCompression -ZipPath "$DestinationBackupDir\*" -ZipDest "$DestinationBackupDir\$ZipFileName" 
    } else {  # Use powershell-native compression
        PowershellCompression -ZipPath "$DestinationBackupDir\*" -ZipDest "$DestinationBackupDir\$ZipFileName" 
    }

    $ZipEnddate = Get-Date
    $ZipSpan = $ZipEndDate - $ZipStartDate
    $ZipDuration = "Zip duration $($ZipSpan.Hours) hours $($ZipSpan.Minutes) minutes $($ZipSpan.Seconds - $SleepTime) seconds"
    Write-au2matorLog -Type INFO -Text "$ZipDuration"

    # This would be a good place to put compression stats. 
    # $TotalMB

}

# 7-Zip Compression
Function 7ZipCompression {
    param(
        [string]$ZipPath,
        [string]$ZipDest
    )
    Write-au2matorLog -Type DEBUG -Text "Use 7-Zip"
    if (test-path $7zPath) { 
        Write-au2matorLog -Type DEBUG -Text "7-Zip found: $($7zPath)!" 
        set-alias sz "$7zPath"
        #sz a -t7z "$directory\$zipfile" "$directory\$name"    
    } else {
        Write-au2matorLog -Type DEBUG -Text "Looking for 7-Zip here: $($7zPath)" 
        Write-au2matorLog -Type ERROR -Text "7-Zip not found: Reverting to Powershell compression" 
        $Use7ZIP = $FALSE
    }
    if ($Use7ZIP -and $UseStaging) { # Zip to the staging directory, then move to the destination.
        # Usage: sz a -t7z <archive.zip> <source1> <source2> <sourceX>
        sz a -t7z $ZipDest $ZipPath
        Write-au2matorLog -Type INFO -Text "Moving Zip to $Destination"
        Move-Item -Path $ZipDest -Destination $Destination

    } else { # Zip straight to the BackupDir. 
        sz a -t7z "$DestinationBackupDir\$ZipFileName" $DestinationBackupDir
    }
}

# Powershell Compression
Function PowershellCompression {
    param(
        [string]$ZipPath,
        [string]$ZipDest
    )
    Write-au2matorLog -Type DEBUG -Text "Using Powershell Compress-Archive"
    #Write-au2matorLog -Type DEBUG -Text "DestinationBackupDir = $($DestinationBackupDir)"
    #Write-au2matorLog -Type DEBUG -Text "StagingDir = $($StagingDir)"
    #Write-au2matorLog -Type DEBUG -Text "ZipFileName = $($ZipFileName)"
    #Write-au2matorLog -Type DEBUG -Text "Destination = $($Destination)"

    $SleepTime = 2 # Seconds
    Write-au2matorLog -Type DEBUG -Text "Pausing for $SleepTime seconds to let things settle"
    Start-sleep -Seconds ($SleepTime) 

    Write-au2matorLog -Type DEBUG -Text "Starting compression"
    Compress-Archive -Path $ZipPath -DestinationPath $ZipDest  -CompressionLevel Optimal -Force
    Write-au2matorLog -Type INFO -Text "Moving Zip to $Destination"
    Move-Item -Path $ZipDest -Destination $Destination
}

### End of Functions

# Create Backup Dir
NewBackupDir
Write-au2matorLog -Type INFO -Text "----------------------"
Write-au2matorLog -Type DEBUG -Text "Start the Script"

# Start the Backup
$CheckDir = CheckDir
if (-not $CheckDir) {
    Write-au2matorLog -Type ERROR -Text "One of the Directories are not available, Script has stopped"
} else {
    MakeBackup

    $CopyEndDate = Get-Date
    $CopySpan = $CopyEnddate - $BackupStartDate
    Write-au2matorLog -Type INFO -Text "`nCopy duration $($CopySpan.Hours) hours $($CopySpan.Minutes) minutes $($CopySpan.Seconds) seconds"
    Write-au2matorLog -Type INFO -Text "----------------------"

    if ($Zip) {
        ZipBackup
        # Clean-up Staging
        if ($ClearStaging) {
            Write-au2matorLog -Type INFO -Text "Clearing Staging"
            Get-ChildItem -Path $DestinationBackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse -Force
        }

        # Clean-up DestinationBackupDir --- REMOVING: Combine this concept with the Staging concept. 
        Get-ChildItem -Path $DestinationBackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse
        Get-Item -Path $DestinationBackupDir | Remove-Item -Confirm:$False -Recurse
    }

    $TotalEndDate = Get-Date
    $TotalSpan = $TotalEndDate - $BackupStartDate
    $TotalDuration = "Total duration $($TotalSpan.Hours) hours $($TotalSpan.Minutes) minutes $($TotalSpan.Seconds) seconds"
    Write-au2matorLog -Type INFO -Text "$TotalDuration"

}

Write-au2matorLog -Type WARNING -Text "Backup $BackupName Finished"

#Write-Host "Press any key to close ..."
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# SIG # Begin signature block
# MIIPVAYJKoZIhvcNAQcCoIIPRTCCD0ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUFZeYE8LG9MAlad1zNXIDOhOl
# tx6gggzFMIIFwDCCA6igAwIBAgITFgAAAAR84b1HddGLUAAAAAAABDANBgkqhkiG
# 9w0BAQsFADAdMRswGQYDVQQDExJFVFNNTU9NTlBLSU9SMDItQ0EwHhcNMTUwOTIz
# MTYxNTA1WhcNMzEwOTIxMjAzMDIzWjBJMRMwEQYKCZImiZPyLGQBGRYDY29tMRcw
# FQYKCZImiZPyLGQBGRYHamhhY29ycDEZMBcGA1UEAxMQRVRTTU1PUEtJQ0EwMi1D
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALdVbFlHVe/fmyEbMyuT
# CdP0Q5SUwQEjvzy49EV+ld6Z2qCQkVdUErozBr/ZEEIGu9HzQAgQOHeJ1nTb23Hz
# v4XtrLHKyDGXAoTfPfCEVmsvmf7TfC7SkEJku3xNfIJu5GADmj7+lxunbNC/24xc
# kc4IF2RApxir5ewm6obANVrysRhCPqORSIm3mIVNSfZyo7OUnmPH1oauf6uCVx/+
# Rn+ft5iGvECgW9RE7OvKiH/UIpkJgaodD5XjsFlNnx8QWwmcL6bplgqcRJG/uYvi
# t5HlfAX8f7zQLp2BMOr/jR3+GMxqdcYEvYHIJ3CTCLKd03ITKKztTovZkP4OHCec
# rXECAwEAAaOCAcswggHHMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBSVJt/U
# u9LtOgzcbod0rZmQ05457jAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTPANSGpoGz6YLn
# F9x4VRV+T0k5ZDB/BgNVHR8EeDB2MHSgcqBwhjZodHRwOi8vTU1PTU5QS0lDUkwu
# amtoeS5jb20vQ0RQL0VUU01NT01OUEtJT1IwMi1DQS5jcmyGNmh0dHA6Ly9CTU9N
# TlBLSUNSTC5qa2h5LmNvbS9DRFAvRVRTTU1PTU5QS0lPUjAyLUNBLmNybDCBuAYI
# KwYBBQUHAQEEgaswgagwUgYIKwYBBQUHMAGGRmh0dHA6Ly9NTU9NTlBLSUNSTC5q
# a2h5LmNvbS9DRFAvRVRTTU1PTU5QS0lPUjAyX0VUU01NT01OUEtJT1IwMi1DQS5j
# cnQwUgYIKwYBBQUHMAGGRmh0dHA6Ly9CTU9NTlBLSUNSTC5qa2h5LmNvbS9DRFAv
# RVRTTU1PTU5QS0lPUjAyX0VUU01NT01OUEtJT1IwMi1DQS5jcnQwDQYJKoZIhvcN
# AQELBQADggIBADRUPg0di2PCj2+gbaYUWBKuPkE1BAW9vvvwzin1EjJIRO3WwjlD
# n0HiHRrttaJtay/4DiJIFw/6l5fq+/HXZsk6K7eOegdhPxQ9XfAd+Yy8Tz1uibD2
# vKoU5s8xA9d/GYUQs0ON/HwzpamMELj5mmHPz/Wyl2V0rdLf6DlOJQ5Ac0j5TKV1
# 4ZwXCInEYekg6RYj15l7KyhEaIxMh/5TLJ9rkM19N26ssECxAoZH4/HnVe5WmXkJ
# xjE3+VpSwVSkHHOtY5GHnRiQuSHVDWm3HLS9Yv6Puu0WbpO4URK352+jzpufF8cG
# zkTj9gMJdUZjqjcxNCMQdUiATfXsvG7Mq4LQwWA46fGPN1cp/pNDl9fyHCTsiJYd
# DqcpKkfA6XEFNfX/c4cl9nCdVwSal479IATNeMor+vhla228jv/Qd913xS1DCsVk
# sHmXj2i1pO4IXyZwr94rax9EHaRcSMqU9LlU5j0des09MFEFFUkbaKxy54N9vpBp
# infAHHgryzuBxb9FnK0SNNLO200FcihRK9qKL+7lJc3k0B06d/CxWVyh/KdjIbMn
# xRa7SmPyyPHQ1UYELqIdqt7Il+erSSJKUFyrboDc4hA95kx8n5+/mBMZ2TfSk8wl
# nPFzM9HVt+pNmH5rPhXOpgSv5fzq/j+F4w5Z8znTtdBjZIYHrqhCzgBaMIIG/TCC
# BeWgAwIBAgITUAAGmGLjUbe2rIz1RwAAAAaYYjANBgkqhkiG9w0BAQsFADBJMRMw
# EQYKCZImiZPyLGQBGRYDY29tMRcwFQYKCZImiZPyLGQBGRYHamhhY29ycDEZMBcG
# A1UEAxMQRVRTTU1PUEtJQ0EwMi1DQTAeFw0yMDA1MDEyMTM2MTNaFw0yMzA1MDEy
# MTM2MTNaMB8xHTAbBgNVBAMMFFJEYXNzb0BqYWNraGVucnkuY29tMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzsMFGE6bDu/k/+NXyrQS4oXCQ/XlCZ0/
# YE7JAP0lATbYs1vkLygrnZ8cfY+rXi8Z+gGAOog89KChqPDSdN/IjZ2156oFK+LP
# gHyEBJ/ucjzv1+reezbevjqbfy0ibRfI/v4+HGPl1RkXsjWQfnCFhhC7oyIyqJpu
# YwvHtxPWIeb7UNLvlLGLpy/Xbz3HkPurhWhL/OiRiwTkZmxGUoibDRB0fPoDBcjf
# 5mKNw3wB/rRMlRvkx6QQxGnG95FFYIaIjRwzih0QDI/y4tyrkn2YeqiIqZcn9pGO
# 2EAKEcvG5xXXOXO7SHM6jxn2INxj3QOzgO1k9II/jhQDea1+hRo0FQIDAQABo4IE
# BjCCBAIwPQYJKwYBBAGCNxUHBDAwLgYmKwYBBAGCNxUIh9rwZYKHuiuD0Y0/hfq8
# JYf34nplgrmlMoPh+lMCAWQCAQkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwCwYDVR0P
# BAQDAgeAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFAix
# wFYolUgH+0is7RdsctliJpO/MB8GA1UdIwQYMBaAFJUm39S70u06DNxuh3StmZDT
# njnuMIH2BgNVHR8Ege4wgeswgeiggeWggeKGOWh0dHA6Ly9FVFNNTU9QS0lDQTAy
# LkpIQUNPUlAuQ09NL0NEUC9FVFNNTU9QS0lDQTAyLUNBLmNybIY5aHR0cDovL0VU
# U0JNT1BLSUNBMDIuamhhY29ycC5jb20vQ0RQL0VUU01NT1BLSUNBMDItQ0EuY3Js
# hjRodHRwOi8vTU1PTU5QS0lDUkwuamtoeS5jb20vQ0RQL0VUU01NT1BLSUNBMDIt
# Q0EuY3JshjRodHRwOi8vQk1PTU5QS0lDUkwuamtoeS5jb20vQ0RQL0VUU01NT1BL
# SUNBMDItQ0EuY3JsMIICRwYIKwYBBQUHAQEEggI5MIICNTBnBggrBgEFBQcwAoZb
# ZmlsZTovL2M6L3dpbmRvd3Mvc3lzdGVtMzIvQ2VydFNydi9DZXJ0RW5yb2wvRVRT
# TU1PUEtJQ0EwMi5qaGFjb3JwLmNvbV9FVFNNTU9QS0lDQTAyLUNBLmNydDBQBggr
# BgEFBQcwAoZEZmlsZTovL0M6L2luZXRwdWIvQ0RQL0VUU01NT1BLSUNBMDIuamhh
# Y29ycC5jb21fRVRTTU1PUEtJQ0EwMi1DQS5jcnQwXwYIKwYBBQUHMAKGU2h0dHA6
# Ly9FVFNNTU9QS0lDQTAyLkpIQUNPUlAuQ09NL0NEUC9FVFNNTU9QS0lDQTAyLmpo
# YWNvcnAuY29tX0VUU01NT1BLSUNBMDItQ0EuY3J0MF8GCCsGAQUFBzAChlNodHRw
# Oi8vRVRTQk1PUEtJQ0EwMi5qaGFjb3JwLmNvbS9DRFAvRVRTTU1PUEtJQ0EwMi5q
# aGFjb3JwLmNvbV9FVFNNTU9QS0lDQTAyLUNBLmNydDBaBggrBgEFBQcwAoZOaHR0
# cDovL01NT01OUEtJQ1JMLmpraHkuY29tL0NEUC9FVFNNTU9QS0lDQTAyLmpoYWNv
# cnAuY29tX0VUU01NT1BLSUNBMDItQ0EuY3J0MFoGCCsGAQUFBzAChk5odHRwOi8v
# Qk1PTU5QS0lDUkwuamtoeS5jb20vQ0RQL0VUU01NT1BLSUNBMDIuamhhY29ycC5j
# b21fRVRTTU1PUEtJQ0EwMi1DQS5jcnQwDQYJKoZIhvcNAQELBQADggEBAAWzlkE9
# 4vy2bXxiWM0ln0pcNThh6Y4mX+vKcCKMMIT1nT+wXRU/LCLaMKPXs61IN+RvnOzu
# /R9nilgh/oRIkeu3QxgU/EyidkWnnInEwRf2eNK3eUc+4s8I4mEjn+/5wBvRnH6O
# /D9DkVUlEgehlFNxVr3WhTePiKrLfA/zaDVtwNUZgAb8Ze77MJDF5JaKv7hAWiAN
# 8Rj13FZj6oJzCzl4w1lDR9tbWs+sLdhwKHHNpMcxrqSMzFQbnVv//GTELbHx+I6R
# rKwwzeMqfu8kSNfiV86T9lTanv0N6BHQ1U9Hh+yoqlx0cwxHNdOYmhkMhDqX1tq1
# XOEJ9o9PX7v31kYxggH5MIIB9QIBATBgMEkxEzARBgoJkiaJk/IsZAEZFgNjb20x
# FzAVBgoJkiaJk/IsZAEZFgdqaGFjb3JwMRkwFwYDVQQDExBFVFNNTU9QS0lDQTAy
# LUNBAhNQAAaYYuNRt7asjPVHAAAABphiMAkGBSsOAwIaBQCgcDAQBgorBgEEAYI3
# AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgEL
# MQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUVLq4cmSSr+sONgWfYY4Q
# A1+sjmEwDQYJKoZIhvcNAQEBBQAEggEAY5T4rdkCB1Ab9mLqAymVKX4DSWQDK52G
# uAKVi6nEn0xjA7eT/HeZKlUDZV/kvBl3ap7bsvzASGYmHb06s/zjy5bFPSVF4E1B
# BAS7A+0L+B0oAptJG1Vq5FZM5A9ZKhpDXNRUHzcFIaFTimyWn4Ok6QKge03H3UWv
# 86gYaEA/OoxPjOQ7m0QvsG4w2MJ0tEXXoSd/hY9Pn6JPaSzZvubt0alforF+Pvjv
# zB3iUJ+SzU3rm9GJ2mN/amx0fP9FaUixm+kYgDxngocyU9c/ljdj9BM3u6bLW/w5
# TSV+C5UITuvcwmjjoNtT1oE3qEQq23YRoEr1w5xRpDv4T32qZFRBEw==
# SIG # End signature block
