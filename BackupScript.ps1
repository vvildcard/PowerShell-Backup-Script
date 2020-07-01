########################################################
# Name: BackupScript.ps1                              
# Creator: Michael Seidl aka Techguy                    
# CreationDate: 21.01.2014                              
# LastModified: 26.06.2020                               
# Version: 1.5.1
# Doc: http://www.techguy.at/tag/backupscript/
# GitHub: https://github.com/Seidlm/PowerShell-Backup-Script
# PSVersion tested: 3 and 4
#
# PowerShell Self Service Web Portal at www.au2mator.com/PowerShell
#
#
# Description: 
# Copies one or more BackupDirs to the Destination.
# A Progress Bar shows the status of copied MB to the total MB.
# 
# Usage:
# BackupScript.ps1 -BackupDirs "C:\path\to\backup", "C:\another\path\" -Destination "C:\path\to\put\the\backup"
# Change variables in the Variables section.
# Change LoggingLevel to 3 an get more output in Powershell Windows.
# 
#
########################################################
#
# www.techguy.at                                        
# www.facebook.com/TechguyAT                            
# www.twitter.com/TechguyAT                             
# michael@techguy.at 
#
########################################################

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
    [string]$LoggingLevel = "3", # LoggingLevel only for Output in Powershell Window, 1=smart, 3=Heavy

    # Zip
    [bool]$Zip = $True, # Zip the backup. 
    [bool]$Use7ZIP = $False, # Make sure 7-Zip is installed. (https://7-zip.org)
    [string]$7zPath = "$env:ProgramFiles\7-Zip\7z.exe",
    [string]$Versions = "2", # Number of backups you want to keep. 

    # Staging -- Only used if Zip = $True.
    [bool]$UseStaging = $True, # If $False: $Destination will be used for Staging.
    [string]$StagingDir = "$TempDir\Staging", # Temporary location zipping. 
    [bool]$ClearStaging = $True, # If $True: Delete StagingDir after backup. 

    # Email
    [bool]$SendEmail = $False, # $True will send report via email (SMTP send)
    [string]$EmailTo = 'test@domain.com', # List of recipients. For multiple users, use "User01 &lt;user01@example.com&gt;" ,"User02 &lt;user02@example.com&gt;"
    [string]$EmailFrom = 'from@domain.com', # Sender/ReplyTo
    [string]$EmailSMTP = 'smtp.domain.com' # SMTP server address
)

### STOP - No changes from here
### STOP - No changes from here

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
    $BackupDir = "$StagingDir"
} else {
    $BackupDir = "$Destination\$BackupName"
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
    if ($LoggingLevel -eq "3") { Write-Host $Text }
}


# Create the BackupDir
Function New-BackupDir {
    if (Test-Path $BackupDir) {
        Write-au2matorLog -Type DEBUG -Text "Backup/Staging directory already exists: $BackupDir"
    } else {
        New-Item -Path $BackupDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Write-au2matorLog -Type INFO -Text "Created backup/staging directory: $BackupDir"
    }
}

# Delete the BackupDir
Function Remove-BackupDir {
    $RemoveFolder = Get-ChildItem $Destination | where { $_.Attributes -eq "Directory" } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
    $RemoveFolder.FullName | Remove-Item -Recurse -Force 
    Write-au2matorLog -Type Info -Text "Deleted oldest directory backup: $RemoveFolder"
}


# Delete Zip
Function Remove-Zip {
    $RemoveZip = Get-ChildItem $Destination | where { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" } | Sort-Object -Property CreationTime -Descending:$False | Select-Object -First 1
    $RemoveZip.FullName | Remove-Item -Recurse -Force 
    Write-au2matorLog -Type Info -Text "Deleted oldest zip backup: $RemoveZip"
}

# Check if BackupDirs and Destination is available
function Check-Dir {
    Write-au2matorLog -Type Info -Text "Check if BackupDir and Destination exists"
    if (!(Test-Path $BackupDirs)) {
        Write-au2matorLog -Type Error -Text "$BackupDirs does not exist"
        return $False
    }
    if (!(Test-Path $Destination)) {
        Write-au2matorLog -Type Error -Text "$Destination does not exist"
        return $False
    }
    return $True
}

# Save all the Files
Function Make-Backup {
    Write-au2matorLog -Type Info -Text "Started the Backup"
    $BackupDirFiles = @{ } # Hash of BackupDir & Files
    $Files = @()
    $SumMB = 0
    $SumItems = 0
    $SumCount = 0
    $colItems = 0
    Write-au2matorLog -Type Info -Text "Count all files and create the Top Level Directories"

    # Count the number of files to backup. 
    foreach ($Backup in $BackupDirs) {
        # Get recursive list of files for each Backup Dir once and save in $BackupDirFiles to use later.
        # Optimize performance by getting included folders first, and then only recursing files for those.
        # Use -LiteralPath option to work around known issue with PowerShell FileSystemProvider wildcards.
        # See: https://github.com/PowerShell/PowerShell/issues/6733

        $Files = Get-ChildItem -LiteralPath $Backup -recurse -Attributes D+!ReparsePoint, D+H+!ReparsePoint -ErrorVariable +errItems -ErrorAction SilentlyContinue | 
        ForEach-Object -Process { Add-Member -InputObject $_ -NotePropertyName "ParentFullName" -NotePropertyValue ($_.FullName.Substring(0, $_.FullName.LastIndexOf("\" + $_.Name))) -PassThru -ErrorAction SilentlyContinue } |
        Where-Object { $_.FullName -notmatch $exclude -and $_.ParentFullName -notmatch $exclude } |
        Get-ChildItem -Attributes !D -ErrorVariable +errItems -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -notmatch $exclude }
        $BackupDirFiles.Add($Backup, $Files)

        $colItems = ($Files | Measure-Object -property length -sum) 
        $Items = 0
        #Write-au2matorLog -Type Info -Text "DEBUG"
        #Copy-Item -LiteralPath $Backup -Destination $BackupDir -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch $exclude }
        $SumMB += $colItems.Sum.ToString()
        $SumItems += $colItems.Count
    }

    # Calculate total size of the backup. 
    $TotalMB = "{0:N2}" -f ($SumMB / 1MB) + " MB of Files"
    Write-au2matorLog -Type Info -Text "There are $SumItems Files with  $TotalMB to copy"

    # Log any errors from above from building the list of files to backup.
    [System.Management.Automation.ErrorRecord]$errItem = $null
    foreach ($errItem in $errItems) {
        Write-au2matorLog -Type ERROR -Text "Skipping '$($errItem.TargetObject)' - Error: $($errItem.CategoryInfo)"
    }
    Remove-Variable errItem
    Remove-Variable errItems

    # Copy files to the backup location. 
    Write-au2matorLog -Type Info -Text "Copying files to $BackupDir"
    foreach ($Backup in $BackupDirs) {
        $Index = $Backup.LastIndexOf("\")
        $SplitBackup = $Backup.substring(0, $Index)
        $Files = $BackupDirFiles[$Backup]

        foreach ($File in $Files) {
            $restpath = $file.fullname.replace($SplitBackup, "")
            try {
                # Use New-Item to create the destination directory if it doesn't yet exist. Then copy the file.
                New-Item -Path (Split-Path -Path $($BackupDir + $restpath) -Parent) -ItemType "directory" -Force -ErrorAction SilentlyContinue | Out-Null
                Copy-Item -LiteralPath $file.fullname -Destination $($BackupDir + $restpath) -Force -ErrorAction SilentlyContinue | Out-Null
                Write-au2matorLog -Type Info -Text "'$($File.FullName)' copied to '$BackupDir'"
            }
            catch {
                $ErrorCount++
                Write-au2matorLog -Type Error -Text "'$($File.FullName)' returned an error and was not copied"
            }
            $Items += (Get-item -LiteralPath $file.fullname).Length
            $status = "Copy file {0} of {1} and copied {3} MB of {4} MB: {2}" -f $count, $SumItems, $file.Name, ("{0:N2}" -f ($Items / 1MB)).ToString(), ("{0:N2}" -f ($SumMB / 1MB)).ToString()
            $Index = [array]::IndexOf($BackupDirs, $Backup) + 1
            $Text = "Copy data Location {0} of {1}" -f $Index , $BackupDirs.Count
            Write-Progress -Activity $Text $status -PercentComplete ($Items / $SumMB * 100)  
            if ($File.Attributes -ne "Directory") { $count++ }
        }
    }
    $SumCount += $Count
    $SumTotalMB = "{0:N2}" -f ($Items / 1MB) + " MB of Files"
    Write-au2matorLog -Type Info -Text "----------------------"
    Write-au2matorLog -Type Info -Text "Copied $SumCount files with $SumTotalMB"
    Write-au2matorLog -Type Info -Text "$ErrorCount Files could not be copied"


    # Send e-mail with reports as attachments
    if ($SendEmail -eq $True) {
        $EmailSubject = "$env:COMPUTERNAME - Backup on $(get-date -format yyyy.MM.dd)"
        $EmailBody = "Backup Script $(get-date -format yyyy.MM.dd) (last Month).`n
                      Computer: $env:COMPUTERNAME`n
                      Source(s): $BackupDirs`n
                      Backup Location: $Destination"
        Write-au2matorLog -Type Info -Text "Sending e-mail to $EmailTo from $EmailFrom (SMTPServer = $EmailSMTP) "
        # The attachment is $log 
        Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTP -attachment $Log 
    }
}

### End of Functions

# Create Backup Dir
New-BackupDir
Write-au2matorLog -Type Info -Text "----------------------"
Write-au2matorLog -Type Info -Text "Start the Script"

# Check if BackupDir needs to be cleaned and create BackupDir
$Count = (Get-ChildItem $Destination | where { $_.Attributes -eq "Directory" }).count
Write-au2matorLog -Type Info -Text "Check if there are more than $Versions Directories in the BackupDir"

if ($count -gt $Versions) { 
    Write-au2matorLog -Type Info -Text "Found $count Directory Backups"
    Remove-BackupDir
}


if ($Zip) {
    $CountZip = (Get-ChildItem $Destination | where { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" }).count
    Write-au2matorLog -Type Info -Text "Check if there are more than $Versions Zip in the BackupDir"
    if ($CountZip -gt $Versions) {
        Write-au2matorLog -Type Info -Text "Found $CountZip Zip backups"
        Remove-Zip 
    }
}

if (Check-Dir) {
    Write-au2matorLog -Type Error -Text "One of the Directories are not available, Script has stopped"
} else {
    Make-Backup

    $CopyEndDate = Get-Date
    $span = $CopyEnddate - $BackupStartDate
    $CopyDuration = "Copy duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds) seconds"

    Write-au2matorLog -Type Info -Text "$CopyDuration"
    Write-au2matorLog -Type Info -Text "----------------------"
    Write-au2matorLog -Type Info -Text "----------------------" 

    if ($Zip) {
        $ZipStartDate = Get-Date
        Write-au2matorLog -Type Info -Text "Compress the Backup Destination"
        if ($Use7ZIP) {
            Write-au2matorLog -Type Info -Text "Use 7-Zip"
            if (test-path $7zPath) { 
                Write-au2matorLog -Type Info -Text "7-Zip found: $($7zPath)!" 
                set-alias sz "$7zPath"
                #sz a -t7z "$directory\$zipfile" "$directory\$name"    
            } else {
                Write-au2matorLog -Type Warning -Text "Looking for 7-Zip here: $($7zPath)" 
                Write-au2matorLog -Type Warning -Text "7-Zip not found. Aborting!" 
            }

            if ($UseStaging) { # Zip to the staging directory, then move to the destination.
                # Usage: sz a -t7z <archive.zip> <source1> <source2> <sourceX>
                sz a -t7z "$BackupDir\$ZipFileName" $BackupDir
                Write-au2matorLog -Type Info -Text "Move Zip to Destination"
                Move-Item -Path "$BackupDir\$ZipFileName" -Destination $Destination

            } else { # Zip straight to the BackupDir. 
                sz a -t7z "$BackupDir\$ZipFileName" $BackupDir
            }
            
        } else {  # Use powershell-native compression
            Write-au2matorLog -Type Info -Text "Use Powershell Compress-Archive"
            #Write-au2matorLog -Type Info -Text "BackupDir = $($BackupDir)"
            #Write-au2matorLog -Type Info -Text "StagingDir = $($StagingDir)"
            #Write-au2matorLog -Type Info -Text "ZipFileName = $($ZipFileName)"
            #Write-au2matorLog -Type Info -Text "Destination = $($Destination)"

            Start-sleep -Milliseconds 2000 # Sleep for a second to let things settle. 

            Compress-Archive -Path "$BackupDir\*" -DestinationPath "$BackupDir\$ZipFileName"  -CompressionLevel Optimal -Force
            Move-Item -Path "$BackupDir\$ZipFileName" -Destination $Destination
        }

        $ZipEnddate = Get-Date
        $span = $ZipEndDate - $ZipStartDate
        $ZipDuration = "Zip duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds) seconds"
        Write-au2matorLog -Type Info -Text "$ZipDuration"


        # Clean-up Staging
        if ($ClearStaging) {
            Write-au2matorLog -Type Info -Text "Clear Staging"
            Get-ChildItem -Path $BackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse -Force
        }

        # Clean-up BackupDir --- REMOVING: Combine this concept with the Staging concept. 
        Get-ChildItem -Path $BackupDir -Recurse -Force | Remove-Item -Confirm:$False -Recurse
        Get-Item -Path $BackupDir | Remove-Item -Confirm:$False -Recurse
    }

    $span = $ZipEndDate - $BackupStartDate
    $TotalDuration = "Total duration $($span.Hours) hours $($span.Minutes) minutes $($span.Seconds) seconds"
    Write-au2matorLog -Type Info -Text "$TotalDuration"

}

Write-au2matorLog -Type Info -Text "Finished"

#Write-Host "Press any key to close ..."
#$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
