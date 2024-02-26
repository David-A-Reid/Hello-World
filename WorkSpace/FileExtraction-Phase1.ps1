<#

.DESCRIPTION
    Script for Retrieving User Data and Luminate Database log backups from Site

.PARAMETER ParameterName
    Has a toggle debug parameter to allow verbose reporting of progress within log file.

.NOTES
    File Name      : File Extraction.ps1
    Author         : David Reid
    Prerequisite   : Requires PSVerion 5 or later.
    Creation Date  : 10/12/2023
    Version        : 1.0

.AMENDMENTS
# 10/12/2023 - Phase 1 - Being tested by Ivan - DR

#>

# Check file name exists 

function Check-FileExists {
    param (
        [string]$DirectoryPath,
        [string]$FileName
    )

    if (Test-Path "$DirectoryPath\$FileName") {
        Write-DebugMessage "The file $FileName exists in $DirectoryPath."
    } else {
        throw "The file $FileName does not exist in $DirectoryPath."
    }
}

# Check file count 

function Check-FileCount {
    param (
        [string]$DirectoryPath,
        [int]$ExpectedFileCount
    )

    $fileCount = (Get-ChildItem $DirectoryPath | Measure-Object).Count

    if ($fileCount -eq $ExpectedFileCount) {
        Write-DebugMessage "The directory contains exactly $ExpectedFileCount files."
    } else {
        throw "The directory does not contain $ExpectedFileCount files."
    }
}

# Check file size

function Check-FileSize {
    param (
        [string]$DirectoryPath,
        [string]$FileName,
        [int]$MinimumSize
    )

    $file = Get-Item "$DirectoryPath\$FileName"

    if ($file.Length -gt $MinimumSize) {
        Write-DebugMessage "The file $FileName exists in $DirectoryPath and is larger than $MinimumSize bytes."
    } else {
        throw "The file $FileName does not exist in $DirectoryPath or is not larger than $MinimumSize bytes."
    }
}


# Check date format

function Check-Date {
    param (
        [string]$date,
        [string]$format
    )

    if ($date -match '^\d{8}$' -and $date.Length -eq 8) {
        Write-DebugMessage "Valid date entered"
    } else {
        throw "Invalid date format - $date"
    }
}

# Test path valid

function Test-PathDebug {
    param (
        [Parameter(Mandatory=$true)]
        [string]$folder
    )

    if (Test-Path $folder) {
        Write-DebugMessage "The path $folder exists."
    } else {
        throw "The path $folder does not exist."
    }
}

# Copy files over to local machine

function Transfer-Files {
    param (
        [string]$SourceSiteAccess,
        [string]$TargetSiteAccess
    )

    $SourceSiteAccessRemoveLastChar = $SourceSiteAccess.Substring(0, $SourceSiteAccess.Length - 1)
    $TargetSiteAccessRemoveLastChar = $TargetSiteAccess.Substring(0, $TargetSiteAccess.Length - 1)

    if (Test-Path $SourceSiteAccess ) {

        if ((Get-Item $SourceSiteAccess).Length -gt 100MB) {
            $robocopyOutput = robocopy $SourceSiteAccessRemoveLastChar $TargetSiteAccessRemoveLastChar /S /E /J /MT:64 /R:3 /W:10 # /MOVE delete from source if required
            Write-output $robocopyOutput
        } else {
            Copy-Item $SourceSiteAccessRemoveLastChar $TargetSiteAccessRemoveLastChar -Recurse -Force
        }

        if ($? -and (Test-Path $TargetSiteAccess )) {
            Write-DebugMessage "Successful Copy of $SourceSiteAccess to $TargetSiteAccess"
        } else {
            throw "Error copying from $SourceSiteAccess to $TargetSiteAccess - Please check log file"
        }
    } else {
        throw "Source file $SourceSiteAccess not found."
    }
}

function Write-DebugMessage {
    param (
        [string]$Message,
        [bool]$DebugFlag = $debug
    )

    if ($DebugFlag) {
        $MessageTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $Message = $MessageTime + " " + $Message 
        Write-Host $Message
    }
}

function Remove-OldDirectories {
    param (
        [string]$BaseDirectory
    )

    # Get a list of subdirectories in the base directory
    $subDirectories = Get-ChildItem -Path $BaseDirectory -Directory

    # Sort subdirectories by LastWriteTime (most recent first)
    $subDirectories = $subDirectories | Sort-Object LastWriteTime -Descending

    if ($subDirectories.Count -gt 1) {
        # Exclude the most recent subdirectory
        $excludeDirectory = $subDirectories[0]
        $subDirectories = $subDirectories | Where-Object { $_ -ne $excludeDirectory }

        # Remove all other directories and their contents
        foreach ($directory in $subDirectories) {
            Remove-Item -Path $directory.FullName -Recurse -Force
            Write-DebugMessage "Removing --> $($directory.FullName)"
        }

        Write-DebugMessage "Removed old directories and their contents, excluding the most recent one."
    } else {
        Write-DebugMessage "No old directories to remove."
    } 
}

function Copy-FilesToS3 {
    param (
        [string]$LocalPath,
        [string]$S3BucketName,
        [string]$S3Prefix,
        [string]$AWSAccessKey,
        [string]$AWSSecretKey
    )

    # Import the AWS module
    #Install-Module -Name AWSPowerShell -Force -AllowClobber # PRE REQUISITE Needs installing if it has never been used before
    Import-Module AWSPowerShell

    # Set AWS credentials
    Set-AWSCredentials -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey

    # Create an S3 client
    $s3Client = New-Object Amazon.S3.AmazonS3Client

    # Upload files to S3
    Get-ChildItem -Path $LocalPath | ForEach-Object {
        $s3Key = $S3Prefix + $_.Name
        Write-S3Object -BucketName $S3BucketName -File $_.FullName -Key $s3Key
        Write-Host "Uploaded $s3Key to S3"
    }
}

# main program

function File-ExtractionRoutine {
    param (
        [bool]$debug = $true
    )

    # Stop on errors, debug flag, USB drive
	
    cls
    $ErrorActionPreference = "Stop"
	$USBdriveLetter = (Get-WmiObject Win32_Volume -Filter "DriveType='2'").DriveLetter
    if (!$USBdriveLetter) {throw "No USB removable drive letter found."}
    $SourceSiteAccess = $USBdriveLetter + "\"
    $logFiles = "$($SourceSiteAccess)logFiles"

    # Header / Setup debug mode

    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd")
 
    if ($debug) {
        $Debug = $true
        Test-PathDebug -folder $logFiles

        $FileExtractionLogFile = "$($logFiles)\FileExtractionLog_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".log"       
        Start-Transcript -Path $FileExtractionLogFile
    } else {
        $Debug = $false
    }


    Write-DebugMessage "File Extraction Routine - $CurrentDate"
    Write-DebugMessage "USB Drive Letter $USBdriveLetter" 

    # Check Source database folder structure

    $SourceSiteAccessDatabase = $SourceSiteAccess + "database\"
    $SourceSiteAccessUserdata = $SourceSiteAccess + "userdata\" 

    # Process database directory files

    $DatabaseDir = Get-ChildItem -Path $SourceSiteAccessDatabase | Where-Object { $_.PSIsContainer -eq $true }
    if (!$DatabaseDir) {throw "No Database files found to copy."}

    foreach ($file in $DatabaseDir) {

        Write-DebugMessage "database checks. $file" 

        # Get the file count in the directory
        Check-FileCount -DirectoryPath $file.FullName -ExpectedFileCount 4 

        # Check file Name exists

        Check-FileExists -DirectoryPath $file.FullName -FileName "*DB_EXPORTS_UNENC.tar*" 
        Check-FileExists -DirectoryPath $file.FullName -FileName "*DB_EXPORTS_UNENC_passphrase*" 
        Check-FileExists -DirectoryPath $file.FullName -FileName "*LUMINATE_SCENARIO.sql.gz.enc*" 
        Check-FileExists -DirectoryPath $file.FullName -FileName "*LUMINATE_SCENARIO.sql.gz_passphrase*" 

        # Check File Size
        Check-FileSize   -DirectoryPath $file.FullName -FileName "*DB_EXPORTS*" -MinimumSize 0 
 
    }

     # Process userdata directory files

    $userdataDir = Get-ChildItem -Path $SourceSiteAccessUserdata | Where-Object { $_.PSIsContainer -eq $true }
    if (!$userdataDir) {throw "No UserData files found to copy."}

    foreach ($file in $userdataDir) {

        Write-DebugMessage "userdata checks. $file" 

        # Get the file count in the directory
        Check-FileCount -DirectoryPath $file.FullName -ExpectedFileCount 2 

        # Check file Name exists

        Check-FileExists -DirectoryPath $file.FullName -FileName "*USERDATA_EXPORT*" 
        Check-FileExists -DirectoryPath $file.FullName -FileName "*userdata_passphrase*" 

        # Check File Size
        Check-FileSize   -DirectoryPath $file.FullName -FileName "*USERDATA_EXPORT*" -MinimumSize 0 
 
    }

    # Create file target file path variables
    #$TargetSiteAccess = [Environment]::GetFolderPath("MyDocuments") + "\LUMINATE LOGS\"
    $TargetSiteAccess = "C:\LUMINATE LOGS\"
    New-Item -ItemType Directory -Path "$($TargetSiteAccess)" -Force | Out-Null
    $TargetSiteAccessSiteAccessDatabase = $TargetSiteAccess + "database\" + $CurrentDate + "\" 
    $TargetSiteAccessSiteAccessUserdata = $TargetSiteAccess + "userdata\" + $CurrentDate + "\" 

    Write-DebugMessage "Copying Files...."
    Write-DebugMessage "Copying files from $SourceSiteAccessDatabase to $TargetSiteAccessSiteAccessDatabase" 

    # Validate and Copy database files
    Transfer-Files -SourceSiteAccess $SourceSiteAccessDatabase -TargetSiteAccess $TargetSiteAccessSiteAccessDatabase

    Write-DebugMessage "Copying files from $SourceSiteAccessUserdata to $TargetSiteAccessSiteAccessUserdata" 

    # Validate and Copy userdata files
    Transfer-Files -SourceSiteAccess $SourceSiteAccessUserdata -TargetSiteAccess $TargetSiteAccessSiteAccessUserdata

    # Remove old directories from USB stick leaving the latest one

     Remove-OldDirectories -BaseDirectory $SourceSiteAccessDatabase
     Remove-OldDirectories -BaseDirectory $SourceSiteAccessUserdata

    # Copy files up to S3 environment

    #Copy-FilesToS3 -LocalPath "C:\LocalFiles\" -S3BucketName "my-bucket" -S3Prefix "my-folder/" -AWSAccessKey "your-access-key" -AWSSecretKey "your-secret-key"

    if ($debug) {Stop-Transcript} 

}

File-ExtractionRoutine -debug $true