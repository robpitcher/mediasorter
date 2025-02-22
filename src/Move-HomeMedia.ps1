<#
.SYNOPSIS
    Organizes media files into a year/month folder structure based on their creation date.

.DESCRIPTION
    The Move-HomeMedia function moves media files from a source directory to a destination directory,
    organizing them into a folder structure based on the date the media was created (YYYY/MM).
    It uses Windows Shell COM object to access file metadata and determine the creation date.

.PARAMETER SourcePath
    The root path containing the media files to be organized.

.PARAMETER DestinationPath
    The target path where the organized folder structure will be created.

.PARAMETER FileExtensions
    An array of file extensions to process. By default, includes common media formats:
    .jpg, .jpeg, .png, .mp4, .mov, .avi, .mts, .heic, .tif, .bmp, .mpg, .3gp, .wmv, .gif

.EXAMPLE
    Move-HomeMedia -SourcePath "C:\Photos" -DestinationPath "D:\Organized"
    Moves all supported media files from C:\Photos to D:\Organized, organizing them by year and month.

.EXAMPLE
    Move-HomeMedia -SourcePath "C:\Media" -DestinationPath "D:\Sorted" -FileExtensions @('.jpg', '.mp4') -Verbose
    Moves only .jpg and .mp4 files from C:\Media to D:\Sorted with verbose output.

.NOTES
    Author: Rob Pitcher
    Version: 1.0
    Requires: Windows PowerShell 5.1 or later
    The function uses Windows Shell COM object to read file metadata.

.LINK
    https://github.com/robpitcher/mediasorter
#>

function Move-HomeMedia {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'FileExtensions')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [string[]]$FileExtensions = @(
            '.jpg', '.jpeg', '.png', '.mp4', '.mov',
            '.avi', '.mts', '.heic', '.tif', '.bmp',
            '.mpg', '.3gp', '.wmv', '.gif', '.dng'
        )
    )

    begin {
        # Load Windows Shell COM object for accessing file metadata
        $shell = New-Object -ComObject Shell.Application

        # Get all files recursively from source
        $allFiles = Get-ChildItem -Path $SourcePath -File -Recurse
        $filteredFiles = $allFiles.where({ $FileExtensions -contains $_.Extension.ToLower() })
        Write-Output "Found $($filteredFiles.Count) files to process"
    }

    process {
        $filteredFiles | ForEach-Object {
            $currentFile = $_
            Write-Verbose "Processing file: $($currentFile.FullName)"

            try {
                # Get the shell folder and item for metadata access
                $folder = $shell.Namespace($currentFile.Directory.FullName)
                $item = $folder.ParseName($currentFile.Name)

                # Get date taken (index 12 is usually Date Taken)
                $dateTaken = $folder.GetDetailsOf($item, 12)

                if ([string]::IsNullOrEmpty($dateTaken)) {
                    Write-Warning "No date taken found for: $($currentFile.FullName)"
                    return
                }

                # Parse the date
                try {
                    $cleanDateTaken = $dateTaken -replace '[^\x20-\x7E]', ''
                    $dateObj = [DateTime]::Parse($cleanDateTaken)
                    $yearFolder = $dateObj.ToString('yyyy')
                    $monthFolder = $dateObj.Month.ToString('00')

                    # Create destination path
                    $newPath = Join-Path -Path $DestinationPath -ChildPath $yearFolder
                    $newPath = Join-Path -Path $newPath -ChildPath $monthFolder

                    # Create folders if they don't exist
                    if (-not (Test-Path $newPath)) {
                        New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                    }

                    # Generate new file path
                    $newFilePath = Join-Path -Path $newPath -ChildPath $currentFile.Name

                    # Move the file
                    Write-Verbose "Moving to: $newFilePath"
                    Move-Item -Path $currentFile.FullName -Destination $newFilePath -Force
                    Write-Verbose "Successfully moved file"
                }
                catch {
                    Write-Error "Error processing date for file: $($currentFile.FullName). Error: $_"
                }
            }
            catch {
                Write-Error "Error moving file: $($currentFile.FullName). Error: $_"
            }
        }
    }

    end {
        # Clean up COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
}