<#
.SYNOPSIS
    This script is used to organize media files by copying them from source directories to target directories based on their file extensions.

.DESCRIPTION
    The script takes multiple parameters to customize the behavior:
    - $sourceDirs: An array of source directories where the media files are located.
    - $imageTargetDir: The target directory for image files.
    - $videoTargetDir: The target directory for video files.
    - $officeTargetDir: The target directory for office files.
    - $imageExtensions: An array of image file extensions to be copied.
    - $videoExtensions: An array of video file extensions to be copied.
    - $officeExtensions: An array of office file extensions to be copied.
    - $excludeDirs: An array of directories to be excluded from the copy process.

.PARAMETER sourceDirs
    An array of source directories where the media files are located. Multiple values can be provided.

.PARAMETER imageTargetDir
    The target directory for image files.

.PARAMETER videoTargetDir
    The target directory for video files.

.PARAMETER officeTargetDir
    The target directory for office files.

.PARAMETER imageExtensions
    An array of image file extensions to be copied.

.PARAMETER videoExtensions
    An array of video file extensions to be copied.

.PARAMETER officeExtensions
    An array of office file extensions to be copied.

.PARAMETER excludeDirs
    An array of directories to be excluded from the copy process.

.EXAMPLE
    .\media-organizer-tool.ps1 -sourceDirs "C:\Users", "D:\Media" -imageTargetDir "D:\Images" -videoTargetDir "D:\Videos" -officeTargetDir "D:\OfficeFiles" -imageExtensions "*.jpg", "*.jpeg", "*.png" -videoExtensions "*.mp4", "*.avi", "*.mov" -officeExtensions "*.doc", "*.docx", "*.xls", "*.xlsx", "*.ppt", "*.pptx" -excludeDirs "C:\Windows", "C:\Program Files"
    This example runs the script with custom parameters to organize media files from the source directories to the target directories.

.NOTES
    - The script creates the target directories if they don't exist.
    - It excludes the specified directories from the copy process.
    - If a file already exists in the target directory, it checks for file content equality using MD5 hash.
    - If the file is identical, it skips the copy process.
    - If the file is different, it appends a unique identifier to the file name and copies it to the target directory.
#>

param (
    [string[]]$sourceDirs = @("C:\Users"),
    [string]$imageTargetDir = "D:\Images",
    [string]$videoTargetDir = "D:\Videos",
    [string]$officeTargetDir = "D:\OfficeFiles",
    [string[]]$imageExtensions = @("*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp"),
    [string[]]$videoExtensions = @("*.mp4", "*.avi", "*.mov", "*.wmv", "*.flv"),
    [string[]]$officeExtensions = @("*.doc", "*.docx", "*.xls", "*.xlsx", "*.ppt", "*.pptx"),
    [string[]]$excludeDirs = @("C:\Windows", "C:\Program Files")
)

# Helper function to create target directories if they don't exist
function Test-TargetDirectory {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Force -Path $Path
    }
}

# Ensure all target directories exist
Ensure-TargetDirectory -Path $imageTargetDir
Ensure-TargetDirectory -Path $videoTargetDir
Ensure-TargetDirectory -Path $officeTargetDir

# Function to copy files while excluding certain directories and handling duplicates
function Copy-Files {
    param (
        [string[]]$extensions,
        [string]$targetDir
    )
    foreach ($dir in $sourceDirs) {
        foreach ($ext in $extensions) {
            Get-ChildItem -Path $dir -Include $ext -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { -not ($excludeDirs -contains $_.DirectoryName) } |
            ForEach-Object {
                $destinationFile = Join-Path -Path $targetDir -ChildPath $_.Name
                if (-not (Test-Path -Path $destinationFile)) {
                    try {
                        $_ | Copy-Item -Destination $targetDir -Force
                        Write-Host "Copied $($_.FullName) to $targetDir"
                    } catch {
                        Write-Warning "Failed to copy $($_.FullName) to $targetDir"
                    }
                } else {
                    $sourceHash = Get-FileHash -Path $_.FullName -Algorithm MD5
                    $destinationHash = Get-FileHash -Path $destinationFile -Algorithm MD5
                    if ($sourceHash.Hash -eq $destinationHash.Hash) {
                        Write-Host "File $($_.FullName) is identical to $destinationFile, skipping"
                    } else {
                        $newDestinationFile = Join-Path -Path $targetDir -ChildPath ("duplicate_{0}_{1}" -f [System.Guid]::NewGuid(), $_.Name)
                        try {
                            $_ | Copy-Item -Destination $newDestinationFile -Force
                            Write-Host "Copied $($_.FullName) to $newDestinationFile"
                        } catch {
                            Write-Warning "Failed to copy $($_.FullName) to $newDestinationFile"
                        }
                    }
                }
            }
        }
    }
}

# Execute the copy process for each file type
Copy-Files -extensions $imageExtensions -targetDir $imageTargetDir
Copy-Files -extensions $videoExtensions -targetDir $
Copy-Files -extensions $officeExtensions -targetDir $officeTargetDir