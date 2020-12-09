########################################################
# Name: BackupScript.ps1                              
# Version: 2.3
# LastModified: 2020-12-09
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
# Change variables with parameters (ex: Add "-UseStaging $False" to turn off staging)
# Change LoggingLevel to 3 an get more output in Powershell Windows.
# 
# 
# Forked: 2020-08-01
# Forked by: [Ryan Dasso](https://github.com/vvildcard)
# Forked from: v1.5
# Creator: [Michael Seidl](https://github.com/Seidlm) aka [Techguy](http://www.techguy.at/)
# PowerShell Self-Service Web Portal at www.au2mator.com/PowerShell


### Variables

Param(
    # Source/Dest
    [Parameter(Mandatory=$True)][string[]][Alias("b","s")]$BackupDirs, # Folders you want to backup. Comma-delimited. 
        # To hard-code the source paths, set the above line to: $BackupDirs = @("C:\path\to\backup", "C:\another\path\")
    [Parameter(Mandatory=$True)][string][Alias("d")]$Destination, # Backup to this path. Can be a UNC path (\\server\share)
        # To hard-code the destination, set the above line to: $Destination = "C:\path\to\put\the\backup"
    [string[]]$ExcludeDirs=@("$env:SystemDrive\Users\.*\AppData\Local", "$env:SystemDrive\Users\.*\AppData\LocalLow"), # This list of Directories will not be copied. Comma-delimited. 

    # Logging
    [string]$TempDir = "$env:TEMP\BackupScript", # Temporary location for logging and zipping. 
    [string]$LogPath = "$TempDir\Logging",
    [string]$LogFileName = "Log", #  Name
    [string]$LoggingLevel = 2, # LoggingLevel only for Output in Powershell Window, 1=smart, 3=Heavy

    # Zip
    [switch]$Zip, # Zip the backup. 
    [switch]$Use7ZIP = $False, # Make sure 7-Zip is installed. (https://7-zip.org)
    [string]$7zPath = "$env:ProgramFiles\7-Zip\7z.exe",
    [string]$Versions = "2", # Number of backups you want to keep. 

    # Staging -- Only used if Zip = $True.
    [switch]$UseStaging, # If set, $Destination will be used for Staging.
    [string]$StagingDir = "$TempDir\Staging", # Temporary location zipping. 
    [switch]$ClearStaging, # If $True: Delete StagingDir after backup. 

    # Email
    [switch]$SendEmail = $False, # $True will send report via email (SMTP send)
    [string]$EmailTo = 'test@domain.com', # List of recipients. For multiple users, use "User01 &lt;user01@example.com&gt;" ,"User02 &lt;user02@example.com&gt;"
    [string]$EmailFrom = 'from@domain.com', # Sender/ReplyTo
    [string]$EmailSMTP = 'smtp.domain.com' # SMTP server address
)


# Parse Excluded directories
$ExcludeString = ""
foreach ($Entry in $ExcludeDirs) {
    # Exclude the directory itself
    $Temp = "^" + $Entry.Replace("\", "\\") + "$"
    $ExcludeString += $Temp + "|"

    # Exclude the directory's children
    $Temp = "^" + $Entry.Replace("\", "\\") + "\\.*"
    $ExcludeString += $Temp + "|"
}
$ExcludeString = $ExcludeString.Substring(0, $ExcludeString.Length - 1)
[RegEx]$exclude = $ExcludeString

# Set the staging directory and name the backup file or folder. 
$BackupName = "Backup-$(Get-Date -format yyyy-MM-dd-hhmmss)"
$ZipFileName = "$($BackupName).zip"
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

# Delete the BackupDir
Function RemoveBackupDir {
    $RemoveFolder = Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Directory" } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
    $RemoveFolder.FullName | Remove-Item -Recurse -Force 
    Write-au2matorLog -Type INFO -Text "Deleted oldest directory backup: $RemoveFolder"
}


# Delete Zip
Function RemoveZip {
    $RemoveZip = Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
    $RemoveZip.FullName | Remove-Item -Recurse -Force 
    Write-au2matorLog -Type INFO -Text "Deleted oldest zip backup: $RemoveZip"
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
				Where-Object { $_.FullName -notmatch $exclude -and $_.ParentFullName -notmatch $exclude } | `
				Get-ChildItem -Attributes !D -ErrorVariable +errItems -ErrorAction SilentlyContinue | `
				Where-Object { $_.DirectoryName -notmatch $exclude }
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

    # Copy files to the backup location. 
    Write-au2matorLog -Type WARNING -Text "Copying files to $DestinationBackupDir"
    foreach ($Backup in $BackupDirs) {
        $BackupAsObject = Get-Item $Backup  # Convert the backup string path to an object to fix case sensitivity
        $Files = $BackupDirFiles[$Backup]
        # Write-au2matorLog -Type DEBUG -Text "Files = $($Files)"

        foreach ($File in $Files) {
            $RelativePath = $File.FullName.Replace("$($BackupAsObject.Parent)\", "") # RelativePath has the parent folder(s) and file name, starting from the directory being backed up. 
            # Example: "Desktop\myfile.txt"
            Write-au2matorLog -Type DEBUG -Text "RelativePath = $($RelativePath)"
            try {
                # Use New-Item to create the destination directory if it doesn't yet exist. Then copy the file.
                New-Item -Path (Split-Path -Path "$DestinationBackupDir\$RelativePath" -Parent) -ItemType "directory" -Force -ErrorAction SilentlyContinue | Out-Null
                Copy-Item -LiteralPath $File.FullName -Destination "$DestinationBackupDir\$RelativePath" -Force -ErrorAction Continue | Out-Null
                Write-au2matorLog -Type DEBUG -Text "'$($File.FullName)' copied to '$DestinationBackupDir\$RelativePath'"
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

### End of Functions

# Create Backup Dir
NewBackupDir
Write-au2matorLog -Type INFO -Text "----------------------"
Write-au2matorLog -Type DEBUG -Text "Start the Script"

# Check if BackupDir needs to be cleaned
$Count = (Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Directory" }).count
Write-au2matorLog -Type DEBUG -Text "Checking if there are more than $Versions Directories in the Destination"

if ($count -gt $Versions) { 
    Write-au2matorLog -Type INFO -Text "Found $count Directory Backups"
    RemoveBackupDir
}


# Count the previous zip backups and remove the oldest (if needed)
if ($Zip) {
    $CountZip = (Get-ChildItem $Destination | Where-Object { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" }).count
    Write-au2matorLog -Type DEBUG -Text "Checking if there are more than $Versions Zip in the Destination"
    if ($CountZip -gt $Versions) {
        Write-au2matorLog -Type INFO -Text "Found $CountZip Zip backups"
        RemoveZip 
    }
}

# Start the Backup
$CheckDir = CheckDir
if (-not $CheckDir) {
    Write-au2matorLog -Type ERROR -Text "One of the Directories are not available, Script has stopped"
} else {
    MakeBackup

    $CopyEndDate = Get-Date
    $span = $CopyEnddate - $BackupStartDate
    Write-au2matorLog -Type INFO -Text "Copy duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds) seconds"
    Write-au2matorLog -Type INFO -Text "----------------------"

    if ($Zip) {
        $ZipStartDate = Get-Date
        Write-au2matorLog -Type INFO -Text "Compressing the Backup Destination"
        if ($Use7ZIP) {
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
                sz a -t7z "$DestinationBackupDir\$ZipFileName" $DestinationBackupDir
                Write-au2matorLog -Type INFO -Text "Moving Zip to $Destination"
                Move-Item -Path "$DestinationBackupDir\$ZipFileName" -Destination $Destination

            } else { # Zip straight to the BackupDir. 
                sz a -t7z "$DestinationBackupDir\$ZipFileName" $DestinationBackupDir
            }
            
        }
		if (-not $Use7ZIP) {  # Use powershell-native compression
			Write-au2matorLog -Type DEBUG -Text "Using Powershell Compress-Archive"
			#Write-au2matorLog -Type DEBUG -Text "DestinationBackupDir = $($DestinationBackupDir)"
			#Write-au2matorLog -Type DEBUG -Text "StagingDir = $($StagingDir)"
			#Write-au2matorLog -Type DEBUG -Text "ZipFileName = $($ZipFileName)"
			#Write-au2matorLog -Type DEBUG -Text "Destination = $($Destination)"

			$SleepTime = 2 # Seconds
			Write-au2matorLog -Type DEBUG -Text "Pausing for $SleepTime seconds to let things settle"
			Start-sleep -Seconds ($SleepTime) 

            Write-au2matorLog -Type DEBUG -Text "Starting compression"
			Compress-Archive -Path "$DestinationBackupDir\*" -DestinationPath "$DestinationBackupDir\$ZipFileName"  -CompressionLevel Optimal -Force
			Write-au2matorLog -Type INFO -Text "Moving Zip to $Destination"
			Move-Item -Path "$DestinationBackupDir\$ZipFileName" -Destination $Destination
		}

        $ZipEnddate = Get-Date
        $span = $ZipEndDate - $ZipStartDate
        $ZipDuration = "Zip duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds - $SleepTime) seconds"
        Write-au2matorLog -Type INFO -Text "$ZipDuration"

        # This would be a good place to put compression stats. 
        # $TotalMB

        # Clean-up Staging
        if ($ClearStaging) {
            Write-au2matorLog -Type INFO -Text "Clearing Staging"
            Get-ChildItem -Path $DestinationBackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse -Force
        }

        # Clean-up DestinationBackupDir --- REMOVING: Combine this concept with the Staging concept. 
        Get-ChildItem -Path $DestinationBackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse
        Get-Item -Path $DestinationBackupDir | Remove-Item -Confirm:$False -Recurse
    }

    $span = $ZipEndDate - $BackupStartDate
    $TotalDuration = "Total duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds) seconds"
    Write-au2matorLog -Type INFO -Text "$TotalDuration"

}

Write-au2matorLog -Type WARNING -Text "Backup $BackupName Finished"

#Write-Host "Press any key to close ..."
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# SIG # Begin signature block
# MIIPVAYJKoZIhvcNAQcCoIIPRTCCD0ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmlcV2g9YXmFXjtedxov2eo3e
# 4cSgggzFMIIFwDCCA6igAwIBAgITFgAAAAR84b1HddGLUAAAAAAABDANBgkqhkiG
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
# MQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUGl0GTyPTQH4P9nCvWOfd
# Nzy0ytcwDQYJKoZIhvcNAQEBBQAEggEAhtZdhJUTVI7hMuV8FyD8kiYJzSTn90ov
# wm9FaipRrx7tG9IQLtzyTSUJvLABEAbktf03lCnTlNoUoDIq3cakpwpvFMQgz/TX
# KL0cQIvVx9sT+t3HlmhiC+kq1ofkeFMS3dLye8HqWrUTYyWEZY9SlMwCuFBlJP3p
# wEgX/21+LIBNHBvgvsl8hd22TNGz0ukCT9ViaB7YjbYBHXYqREhs89G5I3UGiTEO
# C+gXBiYzVq07Y3jPcLps6FK08uCWrWjrJaL3AfbTXtN7kMWsbtpGUd6/RJDNBJt8
# 6iEVxu2DJUFVhgcTdh59h1byGDFdcoM8w7yShpcSkMUUzTFdhNN00g==
# SIG # End signature block
