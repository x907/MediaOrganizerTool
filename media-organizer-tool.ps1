<#
.SYNOPSIS
    This script is used to organize media files (images, videos, and office files) from source directories to target directories.

.DESCRIPTION
    The script takes the following parameters:
    - sourceDirs: The source directories where the media files are located. Default is "C:\Users".
    - imageTargetDir: The target directory where image files will be copied. Default is "D:\Images".
    - videoTargetDir: The target directory where video files will be copied. Default is "D:\Videos".
    - officeTargetDir: The target directory where office files will be copied. Default is "D:\OfficeFiles".
    - imageExtensions: The file extensions of image files to be copied. Default is "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp".
    - videoExtensions: The file extensions of video files to be copied. Default is "*.mp4", "*.avi", "*.mov", "*.wmv", "*.flv".
    - officeExtensions: The file extensions of office files to be copied. Default is "*.doc", "*.docx", "*.xls", "*.xlsx", "*.ppt", "*.pptx".
    - excludeDirs: The directories to be excluded from the copy process. Default is "C:\Windows", "C:\Program Files".

    The script creates the target directories if they don't exist and then copies the files from the source directories to the respective target directories.
    If a file already exists in the target directory, it checks if the file is identical by comparing the MD5 hash. If it is identical, the file is skipped.
    If it is not identical, a new file with a "duplicate_" prefix and a random number is created in the target directory.

.PARAMETER sourceDirs
    The source directories where the media files are located.

.PARAMETER imageTargetDir
    The target directory where image files will be copied.

.PARAMETER videoTargetDir
    The target directory where video files will be copied.

.PARAMETER officeTargetDir
    The target directory where office files will be copied.

.PARAMETER imageExtensions
    The file extensions of image files to be copied.

.PARAMETER videoExtensions
    The file extensions of video files to be copied.

.PARAMETER officeExtensions
    The file extensions of office files to be copied.

.PARAMETER excludeDirs
    The directories to be excluded from the copy process.

.EXAMPLE
    .\media-organizer-tool.ps1 -sourceDirs "C:\Users\JohnDoe\Documents" -imageTargetDir "D:\Images" -videoTargetDir "D:\Videos" -officeTargetDir "D:\OfficeFiles"

    This example runs the script with custom source directories and target directories.

.NOTES
    - This script requires PowerShell version 3.0 or above.
    - The script may take some time to complete depending on the number of files and the size of the files being copied.
#>
param (
    [string]$sourceDirs = "C:\Users",
    [string]$imageTargetDir = "D:\Images",
    [string]$videoTargetDir = "D:\Videos",
    [string]$officeTargetDir = "D:\OfficeFiles",
    [string[]]$imageExtensions = "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp",
    [string[]]$videoExtensions = "*.mp4", "*.avi", "*.mov", "*.wmv", "*.flv",
    [string[]]$officeExtensions = "*.doc", "*.docx", "*.xls", "*.xlsx", "*.ppt", "*.pptx",
    [string[]]$excludeDirs = "C:\Windows", "C:\Program Files"
)

# Create target directories if they don't exist
New-Item -ItemType Directory -Force -Path $imageTargetDir
New-Item -ItemType Directory -Force -Path $videoTargetDir
New-Item -ItemType Directory -Force -Path $officeTargetDir

# Function to copy files while excluding certain directories
function Copy-Files {
    param (
        [string[]]$extensions,
        [string]$targetDir
    )
    foreach ($ext in $extensions) {
        Get-ChildItem -Path $sourceDirs -Include $ext -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $excludeDirs -notcontains $_.DirectoryName } |
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
                    $newDestinationFile = Join-Path -Path $targetDir -ChildPath ("duplicate_{0}_{1}" -f (Get-Random -Maximum 9999), $_.Name)
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

# Execute the copy for images, videos, and office files
Copy-Files -extensions $imageExtensions -targetDir $imageTargetDir
Copy-Files -extensions $videoExtensions -targetDir $videoTargetDir
Copy-Files -extensions $officeExtensions -targetDir $officeTargetDir

Write-Host "Images, videos, and office files have been copied to $imageTargetDir, $videoTargetDir, and $officeTargetDir."