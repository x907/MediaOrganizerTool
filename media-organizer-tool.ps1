# Define the source and target directories
$sourceDirs = "C:\Users" # Adjust this path to where you want to search
$imageTargetDir = "D:\Images" # Change D:\ with your new drive letter
$videoTargetDir = "D:\Videos" # Change D:\ with your new drive letter
$officeTargetDir = "D:\OfficeFiles" # Change D:\ with your new drive letter

# Define extensions to search for
$imageExtensions = "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp"
$videoExtensions = "*.mp4", "*.avi", "*.mov", "*.wmv", "*.flv"
$officeExtensions = "*.doc", "*.docx", "*.xls", "*.xlsx", "*.ppt", "*.pptx" # Add more extensions if needed

# Define directories to exclude (stock images/videos locations)
$excludeDirs = "C:\Windows", "C:\Program Files" # Adjust these paths

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