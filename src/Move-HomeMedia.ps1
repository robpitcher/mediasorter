function Move-HomeMedia {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [string[]]$FileExtensions = @('.jpg', '.jpeg', '.png', '.mp4', '.mov', '.avi', '.mts', '.heic', '.tif', '.bmp', '.mpg', '.3gp', '.wmv', '.gif')
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
                    $monthFolder = $dateObj.ToString('MM')

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