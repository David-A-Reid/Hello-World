<#

.DESCRIPTION
    Script for Retrieving User Data and Luminate Database log backups from Site

.PARAMETER ParameterName
    Has a toggle debug parameter to allow verbose reporting of progress within log file.
    v35

.NOTES
    File Name      : File Extraction.ps1 
    Author         : David Reid
    Prerequisite   : Requires PSVerion 5 or later.
    Creation Date  : 10/12/2023
    Version        : 1.1 

.AMENDMENTS
 10/12/2023 - V1.0 - DR - Phase 1
 Tested by Ivan - DR
 
 21/02/2023 - V1.1 - DR 
 Added functionality to allow copying from hard disk to s3 
 Added prompted for site details 
 Added prompting to perform copy from USB to Hard disk if required and the same functionality for copying from Harddisk to S3

#>

# Check file name exists 

function Check-FileExists {
    param (
        [string]$DirectoryPath,
        [string]$FileName
    )

    $matchingFiles = Get-ChildItem -Path $DirectoryPath -Filter $FileName

    if ($matchingFiles.Count -gt 0) {
        $firstMatchingFile = $matchingFiles[0]
        Write-DebugMessage "The file $firstMatchingFile exists in $DirectoryPath."
        return $firstMatchingFile.Name
    } 
    else
        {Throw-Error "The file $FileName does not exist in $DirectoryPath."}
}

# Check file count 

function Check-FileCount {
    param (
        [string]$DirectoryPath,
        [int]$ExpectedFileCount
    )

    $fileCount = (Get-ChildItem $DirectoryPath | Measure-Object).Count

    if ($fileCount -eq $ExpectedFileCount) 
        {Write-DebugMessage "The directory contains exactly $ExpectedFileCount files."}
    else 
        {Throw-Error "The directory does not contain $ExpectedFileCount files."}
}

# Check file size

function Check-FileSize {
    param (
        [string]$DirectoryPath,
        [string]$FileName,
        [int]$MinimumSize
    )

    $file = Get-Item "$DirectoryPath\$FileName"

    if ($file.Length -gt $MinimumSize) 
        {Write-DebugMessage "The file $FileName exists in $DirectoryPath and is larger than $MinimumSize bytes."}
    else 
        {Throw-Error "The file $FileName does not exist in $DirectoryPath or is not larger than $MinimumSize bytes."}
}

# Test path and create if doesn't exist

function Test-PathDebug {
    param ([Parameter(Mandatory=$true)][string]$folder)

    if (Test-Path $folder) {
        Write-DebugMessage "The path $folder exists."
    } else {
        try {
            New-Item -Path $folder -ItemType Directory -Force
            Write-DebugMessage "Created the path $folder."
        } 
        catch {Throw-Error "Failed to create the path $folder."}
    }
}

# Copy files over to local machine

function Transfer-Files {
    param (
        [string]$SourceSiteAccess,
        [string]$TargetSiteAccess
    )

    $Directory = Get-ChildItem -Path $SourceSiteAccess 

    foreach ($file in $Directory) {

        $TargetSiteAccess = "$TargetSiteAccess$($file.name)"

        if ((Get-Item $file.Fullname).Length -gt 100MB) {
            $copyOutput = robocopy $file.Fullname $TargetSiteAccess /S /E /J /MT:64 /R:3 /W:10 # /MOVE delete from source if required
            Write-Output $copyOutput
        } else {
            $copyOutput = Copy-Item $file.Fullname $TargetSiteAccess -Recurse -Force
            Write-Output $copyOutput
        }
        
        if  ($? -and (Test-Path $TargetSiteAccess)) 
            {Write-DebugMessage "Successful Copy of $($file.Fullname) to $TargetSiteAccess"}
        else 
            {Throw-Error "Error copying from $($file.Fullname) to $TargetSiteAccess - Please check log file"}
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
    param ([string]$BaseDirectory)

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
     }
     else 
        {Write-DebugMessage "No old directories to remove."} 
}

function Set-AWSEnvironmentVariable {
    param (
        [string]$CredentialPath = "C:\LUMINATE LOGS\credentials\cred.xml",
        [switch]$ForceCredentialCreation
    )

    if ((Test-Path -Path $CredentialPath) -and !$ForceCredentialCreation) {
        Write-DebugMessage "Credential file already exists at: $CredentialPath"
        $credentials = Import-CliXml -Path $CredentialPath

        $AWSAccessKeyPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credentials.AWSAccessKey))
        $AWSSecretKeyPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($credentials.AWSSecretKey))

        $env:AWSAccessKey = $AWSAccessKeyPlainText
        $env:AWSSecretKey = $AWSSecretKeyPlainText

    }
    else {
        $UserInput = Read-Host "Enter a value for your AWSAccessKey environment variable"
        $env:AWSAccessKey = $UserInput
        $SecureAWSAccessKey = ConvertTo-SecureString -String $UserInput -AsPlainText -Force

        $UserInput = Read-Host "Enter a value for your AWSSecretKey environment variable"
        $env:AWSSecretKey = $UserInput
        $SecureAWSSecretKey = ConvertTo-SecureString -String $UserInput -AsPlainText -Force

        # Save the credentials to a file
        $credentials = @{AWSAccessKey = $SecureAWSAccessKey;AWSSecretKey = $SecureAWSSecretKey}
        $credentials | Export-CliXml -Path $CredentialPath

        Write-DebugMessage "Credential saved to file: $CredentialPath"
    }

    # Display the environment variable values
    #Write-Host "Environment variable AWSAccessKey set to: $env:AWSAccessKey"
    #Write-Host "Environment variable AWSSecretKey set to: $env:AWSSecretKey"
}

function Upload-FileToS3 {
    param (
        [Parameter(Mandatory=$true)]
        [string]$LocalFilePath,
        [Parameter(Mandatory=$true)]
        [string]$S3File,
        [Parameter(Mandatory=$true)]
        [string]$BucketName
    )

    if (-not (Get-Module -ListAvailable -Name AWSPowerShell)) {
        Import-Module AWSPowerShell
        Write-DebugMessage "Using AWS Powershell module"
    }

    # Set AWS credentials with the region

    try   {Set-AWSCredentials -AccessKey $env:AWSAccessKey -SecretKey $env:AWSSecretKey}
    Catch {Throw-Error "Error setting credentials for S3 login"}

    try   {Set-DefaultAWSRegion -Region eu-west-1}
    catch {Throw-Error "Error setting region for for S3 Environment"}

    #$S3folderExists = Get-S3Object -BucketName $BucketName -KeyPrefix $S3File
    #if (!(Get-S3Object -BucketName $bucketName -Key "$S3File")) {
    #    write-host "$S3File does not exist"
    #} else {
    #    Write-Host "The file $S3File already exists in the folder ."
    #}

    try {           
            #Write-S3Object -BucketName $BucketName -File $LocalFilePath -Key $S3File -ErrorAction Stop
            Write-DebugMessage "File uploaded to s3://$BucketName/$S3File"
    }
    catch {Throw-Error "Error uploading file to S3: This could be AWS access issue or $_"}
}

function Validate-Prompt {
    param (
        [int]$attempts = 3,
        [string[]]$validValues ,
        [string]$promptMessage 
    )

    $attemptsRemaining = $attempts

    while ($attemptsRemaining -gt 0) {
        $inputSite = Read-Host $promptMessage

        if ($validValues -contains $inputSite) {
            $capitalizedSite = $inputSite.ToUpper()
            return $capitalizedSite
        }
        elseif ($inputSite -eq '') {
            Write-DebugMessage "Input cannot be blank. Please try again."
        }
        else {
            Write-DebugMessage "Invalid input. Please enter one of the following: $($validValues -join ', ')."
        }

        $attemptsRemaining--
    }

    # Throw-Error exception if validation fails the specified number of times
    Throw-Error "Exceeded maximum attempts. Program terminated."
}

function Upload-FileToS3v2 {
    param ([Parameter(Mandatory=$true)][array]$TransferDetails)

    if (-not (Get-Module -ListAvailable -Name AWSPowerShell)) {
        Import-Module AWSPowerShell
        Write-DebugMessage "Using AWS Powershell module"
    }

    # Set AWS credentials with the region

    try   {Set-AWSCredentials -AccessKey $env:AWSAccessKey -SecretKey $env:AWSSecretKey}        
    Catch {Throw-Error "Error setting credentials for S3 login"}

    try   {
            Set-DefaultAWSRegion -Region eu-west-1
            get-DefaultAWSRegion
          }
    catch
          {Throw-Error "Error setting region for for S3 Environment"}

    try {             
         foreach ($entry in $TransferDetails) {
             $awsCopyCommand = "Upload-FileToS3 -LocalFilePath `"$($entry.Name)`" -S3File `"$($entry.CopyPath)`" -BucketName `"tech.resonate.logs`""
             Write-DebugMessage $awsCopyCommand
             #Write-S3Object -BucketName $BucketName -File $LocalFilePath -Key $S3File -ErrorAction Stop
             Write-DebugMessage "File uploaded to s3://$BucketName/$S3File"
         }  
    }
    catch 
        {Throw-Error "Error uploading file to S3: This could be AWS access issue or $_"}
}

function Throw-Error {
    param ([Parameter(Mandatory=$true)][string]$ErrorDetails)

    Throw $ErrorDetails
    if ($debug) {Stop-Transcript}
}

# main program

function File-ExtractionRoutine {
    param ([bool]$debug = $true)

    # Stop on errors, debug flag, USB drive
	
    cls
    $ErrorActionPreference = "Stop"
	$USBdriveLetter = (Get-WmiObject Win32_Volume -Filter "DriveType='2'").DriveLetter
    if (!$USBdriveLetter) {Throw-Error "No USB removable drive letter found."}
    $SourceSiteAccess = $USBdriveLetter + "\"
    $logFiles = "$($SourceSiteAccess)logFiles"

    # Pre Housekeeping before starting routine
    
    if (Test-Path -Path "$($logFiles)\Please Ignore required for Copy to Zip file purposes.txt") {remove-item "$logFiles\Please Ignore required for Copy to Zip file purposes.txt"}
    if (Test-Path -Path "$($SourceSiteAccess)userdata\Please Ignore required for Copy to Zip file purposes.txt") {remove-item "$($SourceSiteAccess)userdata\Please Ignore required for Copy to Zip file purposes.txt"}
    if (Test-Path -Path "$($SourceSiteAccess)database\Please Ignore required for Copy to Zip file purposes.txt") {remove-item "$($SourceSiteAccess)database\Please Ignore required for Copy to Zip file purposes.txt"}
    if (Test-Path -Path "$($SourceSiteAccess)File Extraction.zip") {remove-item "$($SourceSiteAccess)File Extraction.zip"}

    # Header / Setup debug mode

    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd")

    # Check debug location for logs

    if ($debug) {
        $Debug = $true
        Test-PathDebug -folder $logFiles

        $FileExtractionLogFile = "$($logFiles)\FileExtractionLog_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".log"       
        Start-Transcript -Path $FileExtractionLogFile
    } else {
        $Debug = $false
    }

    # Ensure main directories are setup on Laptop

    Test-PathDebug -folder "C:\LUMINATE LOGS\"
    Test-PathDebug -folder "C:\LUMINATE LOGS\userdata"
    Test-PathDebug -folder "C:\LUMINATE LOGS\database"
    Test-PathDebug -folder "C:\LUMINATE LOGS\credentials"

    $BannerHead = @("
    #######                    #######                                                            ######                                      
    #       # #      ######    #       #    # ##### #####    ##    ####  ##### #  ####  #    #    #     #  ####  #    # ##### # #    # ###### 
    #       # #      #         #        #  #    #   #    #  #  #  #    #   #   # #    # ##   #    #     # #    # #    #   #   # ##   # #      
    #####   # #      #####     #####     ##     #   #    # #    # #        #   # #    # # #  #    ######  #    # #    #   #   # # #  # #####  
    #       # #      #         #         ##     #   #####  ###### #        #   # #    # #  # #    #   #   #    # #    #   #   # #  # # #      
    #       # #      #         #        #  #    #   #   #  #    # #    #   #   # #    # #   ##    #    #  #    # #    #   #   # #   ## #      
    #       # ###### ######    ####### #    #   #   #    # #    #  ####    #   #  ####  #    #    #     #  ####   ####    #   # #    # ######
    "
    )

    Write-DebugMessage $BannerHead
    Write-DebugMessage "Date: $CurrentDate"
    Write-DebugMessage "USB Drive Letter $USBdriveLetter" 

    $ProceedWithUpload = Validate-Prompt -promptMessage "Do you want to proceed with the copy from USB to your PC ? - Type (Yes or No)" -attempts 6 -validValues @('Yes', 'No')
    if ($ProceedWithUpload -eq "YES") {

        # Get Site information to setup S3 target PATH variables

        try {Write-DebugMessage "Please enter Site details :-"
            Write-DebugMessage "Ang - Anglia"
            Write-DebugMessage "Wes - Western"
            Write-DebugMessage "Sco - Scotland"
            Write-DebugMessage "Mar - Marylebone"

            $inputSite = Validate-Prompt -promptMessage "Enter a three-character Site Code" -attempts 6 -validValues @('Ang', 'Wes' ,'Sco', 'Mar')
            Write-DebugMessage "Site received: $inputSite"
        }
        catch 
            {Write-DebugMessage "Error: $_"}       

        $s3CopyDetails = @() # Array for storing all source files to be copied and their destinations
        
        switch ($inputSite) {
            'ANG' { 
                $s3DatabaseCopyPath = "romford/tms/raw/full-databases/$($CurrentDate)/"
                $s3UserdataCopyPath = "romford/tms/raw/userdata/$($CurrentDate)/"
                $s3ScenarioCopyPath = "romford/tms/raw/scenario-databases/$($CurrentDate)/"
            }
            'WES' { 
                #$s3DatabaseCopyPath = "tvsc/tms/raw/full-databases/$($CurrentDate)/"
                #$s3UserdataCopyPath = "tvsc/tms/raw/userdata/$($CurrentDate)/"
                #$s3ScenarioCopyPath = "tvsc/tms/raw/scenario-databases/$($CurrentDate)/"
                $s3DatabaseCopyPath = "tvsc/tms/raw/full-databases/2024-02-13/"
                $s3UserdataCopyPath = "tvsc/tms/raw/userdata/2024-02-13/"
                $s3ScenarioCopyPath = "tvsc/tms/raw/scenario-databases/2024-02-13/"
            }
            # Need Marylebone site details
            'MAR' { 
                $s3DatabaseCopyPath = "marylebone/tms/raw/full-databases/$($CurrentDate)/"
                $s3UserdataCopyPath = "marylebone/tms/raw/userdata/$($CurrentDate)/"
                $s3ScenarioCopyPath = "marylebone/tms/raw/scenario-databases/$($CurrentDate)/"
            }
            # Need Scottish site details
            'SCO' { 
                $s3DatabaseCopyPath = "scotland/tms/raw/full-databases/$($CurrentDate)/"
                $s3UserdataCopyPath = "scotland/tms/raw/userdata/$($CurrentDate)/"
                $s3ScenarioCopyPath = "scotland/tms/raw/scenario-databases/$($CurrentDate)/"
            }
            '' { 
                Throw-Error "Site Input cannot be blank. Program terminated."
            }
        }

        # Check Source database folder structure

        $SourceSiteAccessDatabase = $SourceSiteAccess + "database\"
        $SourceSiteAccessUserdata = $SourceSiteAccess + "userdata\" 

        # Create target file path variables

        $TargetSiteAccess = "C:\LUMINATE LOGS\"
        New-Item -ItemType Directory -Path "$($TargetSiteAccess)" -Force | Out-Null
        $TargetSiteAccessSiteAccessDatabase = $TargetSiteAccess + "database\" + $CurrentDate + "\" 
        $TargetSiteAccessSiteAccessUserdata = $TargetSiteAccess + "userdata\" + $CurrentDate + "\" 

        # Process database directory files

        $DatabaseDir = Get-ChildItem -Path $SourceSiteAccessDatabase | Where-Object { $_.PSIsContainer -eq $true }
        if (!$DatabaseDir) {Throw-Error "No Database files found to copy."}

        foreach ($file in $DatabaseDir) {

            Write-DebugMessage "database checks. $file" 

            # Get the file count in the directory

            Check-FileCount  -DirectoryPath $file.FullName -ExpectedFileCount 4 

            # Check file Name exists

            $file1 = Check-FileExists -DirectoryPath $file.FullName -FileName "*DB_EXPORTS_UNENC.tar*" 
            $file1Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessDatabase)$file1";CopyPath = "$s3DatabaseCopyPath$($file1)"}

            $file2 = Check-FileExists -DirectoryPath $file.FullName -FileName "*DB_EXPORTS_UNENC_passphrase*" 
            $file2Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessDatabase)$file2";CopyPath = "$s3DatabaseCopyPath$($file2)"}
                    
            $file3 = Check-FileExists -DirectoryPath $file.FullName -FileName "*LUMINATE_SCENARIO.sql.gz.enc*" 
            $file3Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessDatabase)$file3";CopyPath = "$s3ScenarioCopyPath$($file3)"}
            
            $file4 = Check-FileExists -DirectoryPath $file.FullName -FileName "*LUMINATE_SCENARIO.sql.gz_passphrase*" 
            $file4Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessDatabase)$file4";CopyPath = "$s3ScenarioCopyPath$($file4)"}

            # Check File Size

            Check-FileSize   -DirectoryPath $file.FullName -FileName "*DB_EXPORTS_UNENC.tar*" -MinimumSize 0 # 2097152 # 2 Megabyte

            $s3CopyDetails += $file1Info
            $s3CopyDetails += $file2Info
            $s3CopyDetails += $file3Info
            $s3CopyDetails += $file4Info
        }

        # Process userdata directory files

        $userdataDir = Get-ChildItem -Path $SourceSiteAccessUserdata | Where-Object { $_.PSIsContainer -eq $true }
        if (!$userdataDir) {Throw-Error "No UserData files found to copy."}

        foreach ($file in $userdataDir) {

            Write-DebugMessage "userdata checks. $file" 

            # Get the file count in the directory
 
            Check-FileCount -DirectoryPath $file.FullName -ExpectedFileCount 2 

            # Check file Name exists

            $file5 = Check-FileExists -DirectoryPath $file.FullName -FileName "*USERDATA_EXPORT*" 
            $file5Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessUserdata)$file5"; CopyPath = "$s3UserdataCopyPath$($file5)"}

            $file6 = Check-FileExists -DirectoryPath $file.FullName -FileName "*userdata_passphrase*" 
            $file6Info = [PSCustomObject]@{Name = "$($TargetSiteAccessSiteAccessUserdata)$file6"; CopyPath = "$s3UserdataCopyPath$($file6)"}

            # Check File Size

            Check-FileSize   -DirectoryPath $file.FullName -FileName "*USERDATA_EXPORT*" -MinimumSize 0  # 2097152 # 2 Megabyte
 
            $s3CopyDetails += $file5Info
            $s3CopyDetails += $file6Info
        }

        Write-DebugMessage "Copying Files...."

        # Validate and Copy database & scenario files

        Write-DebugMessage "Copying files from $SourceSiteAccessDatabase to $TargetSiteAccessSiteAccessDatabase" 
        Transfer-Files -SourceSiteAccess $SourceSiteAccessDatabase -TargetSiteAccess $TargetSiteAccessSiteAccessDatabase

        # Validate and Copy userdata files

        Write-DebugMessage "Copying files from $SourceSiteAccessUserdata to $TargetSiteAccessSiteAccessUserdata" 
        Transfer-Files -SourceSiteAccess $SourceSiteAccessUserdata -TargetSiteAccess $TargetSiteAccessSiteAccessUserdata

        # Remove old directories from USB stick leaving the latest one

        Remove-OldDirectories -BaseDirectory $SourceSiteAccessDatabase
        Remove-OldDirectories -BaseDirectory $SourceSiteAccessUserdata

        Write-DebugMessage "********************************************************"
        Write-DebugMessage "****** Copying Files to LapTop has been completed ******"
        Write-DebugMessage "********************************************************"
    }

    # Copy files up to S3 environment

    $ProceedWithS3Copy = Validate-Prompt -promptMessage "Do you want to proceed with the copy to S3? - Type (Yes or No)" -attempts 6 -validValues @('Yes', 'No')
    if ($ProceedWithS3Copy -eq "YES") {
         $S3CopyHeader = @("
          #####                         #######                                            #####   #####  
        #     #  ####  #####  #   #    #       # #      ######  ####     #####  ####     #     # #     # 
        #       #    # #    #  # #     #       # #      #      #           #   #    #    #             # 
        #       #    # #    #   #      #####   # #      #####   ####       #   #    #     #####   #####  
        #       #    # #####    #      #       # #      #           #      #   #    #          #       # 
        #     # #    # #        #      #       # #      #      #    #      #   #    #    #     # #     # 
         #####   ####  #        #      #       # ###### ######  ####       #    ####      #####   ##### 
         "    )

         Write-DebugMessage $S3CopyHeader

         # Get AWS Credentials   
         # Prompt for passord change for AWS s3 access if required

         Write-DebugMessage "If you have errors in the process when copying to S3 you may want to consider"
         Write-DebugMessage "resetting the AWS password, otherwise please ignore this next prompt and"
         Write-DebugMessage "just press return"

         $ProceedWithAWSPasswordreset = Validate-Prompt -promptMessage "Do you want to proceed with AWS password reset for accessing S3? - Type (Yes or No)" -attempts 6 -validValues @('Yes', 'No', '')
         if ($ProceedWithAWSPasswordreset -eq "YES") 
            {Set-AWSEnvironmentVariable -ForceCredentialCreation}
         else
            {Set-AWSEnvironmentVariable}
         

         foreach ($entry in $s3CopyDetails) {
             $awsCopyCommand = "Upload-FileToS3 -LocalFilePath `"$($entry.Name)`" -S3File `"$($entry.CopyPath)`" -BucketName `"tech.resonate.logs`""
             Write-DebugMessage $awsCopyCommand
             Upload-FileToS3 -LocalFilePath "$($entry.Name)" -S3File "$($entry.CopyPath)" -BucketName "tech.resonate.logs"
         }

    } else
    {
        Write-DebugMessage "Please follow the manual process of copying the files to S3 when you are in the appropriate environment"
        Write-DebugMessage "Follow the instructions as detailed on the following link - https://resonate.atlassian.net/wiki/spaces/DA/pages/866746440/Raw+Data+Upload+Instructions"
    }

    if ($debug) {Stop-Transcript} 

}

File-ExtractionRoutine -debug $true