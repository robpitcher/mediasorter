# MediaSorter

A PowerShell module to organize your media files (photos and videos) into a structured folder hierarchy based on their creation dates.

⚠️ **WARNING**
> Before using this tool, please ensure you have a backup of all your media files. While this tool is designed to safely handle your files, unexpected issues can occur during file operations. The authors are not responsible for any data loss that may occur while using this software. Always verify your backups are current before proceeding.

## Features

- Automatically organizes media files into Year/Month folders
- Supports multiple media formats including jpg, jpeg, png, mp4, mov, and more
- Uses file metadata to determine creation dates
- Preserves original filenames
- Provides verbose logging options
- Supports WhatIf parameter for safe testing

## Requirements

- Windows PowerShell 5.1 or later
- Windows OS (uses Windows Shell COM object for metadata access)

## Installation

1. Clone this repository:
```powershell
git clone https://github.com/robpitcher/mediasorter.git
```

2. Import the module:
```powershell
Import-Module .\src\Move-HomeMedia.ps1
```

## Usage

### Basic Usage

```powershell
Move-HomeMedia -SourcePath "C:\UnsortedPhotos" -DestinationPath "D:\OrganizedPhotos"
```

### With Specific File Types

```powershell
Move-HomeMedia -SourcePath "C:\Media" -DestinationPath "D:\Sorted" -FileExtensions @('.jpg', '.mp4')
```

### Verbose Output

```powershell
Move-HomeMedia -SourcePath "C:\Photos" -DestinationPath "D:\Organized" -Verbose
```

## Supported File Extensions

- Images: .jpg, .jpeg, .png, .heic, .tif, .bmp, .gif
- Videos: .mp4, .mov, .avi, .mts, .mpg, .3gp, .wmv

## Output Structure

```
DestinationPath/
├── 2024/
│   ├── 01/
│   │   ├── photo1.jpg
│   │   └── video1.mp4
│   └── 02/
│       └── photo2.jpg
└── 2023/
    └── 12/
        └── photo3.jpg
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
