# LibraryLint

A linter for your media library. LibraryLint is a PowerShell tool for media library organization and cleanup. Designed for managing movie and TV show collections with support for Kodi/Plex/Jellyfin-compatible naming and metadata.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-10+-green.svg)
![Version](https://img.shields.io/badge/Version-5.0-orange.svg)

## Features

### Core Functionality
- **Dry-run mode** - Preview all changes before applying them
- **Comprehensive logging** - All operations logged with timestamps
- **Progress tracking** - Visual progress indicators for long operations
- **Automatic 7-Zip installation** - Installs 7-Zip if not present

### Movie Processing
- Extract archives (.rar, .zip, .7z, .tar, .gz, .bz2)
- Remove unnecessary files (samples, proofs, screenshots)
- Process trailers (move to `_Trailers` folder or delete)
- Process subtitles (keep preferred language, delete others)
- Create individual folders for loose video files
- Clean folder names by removing quality/codec/release tags
- Format movie years with parentheses (`Movie 2024` → `Movie (2024)`)
- Generate Kodi-compatible NFO files

### TV Show Processing
- Extract all archives
- Parse episode info (S01E01, 1x01, multi-episode S01E01-E03)
- Organize episodes into Season folders
- Rename episodes to standard format
- Detect missing episodes (gap detection)
- Remove empty folders

### Advanced Features
- **Duplicate Detection** - Find duplicates using file hashing and quality scoring
- **TMDB Integration** - Fetch metadata from The Movie Database
- **Codec Analysis** - Analyze video codecs and generate FFmpeg transcode scripts
- **Health Check** - Validate library for issues (empty folders, missing files, etc.)
- **MediaInfo Integration** - Accurate codec detection from file headers
- **Export Reports** - Generate CSV, HTML, and JSON library reports
- **Undo/Rollback** - Manifest-based rollback of changes
- **Configuration Files** - Save/load settings to JSON

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later
- 7-Zip (automatically installed if not present)
- [MediaInfo CLI](https://mediaarea.net/en/MediaInfo) (optional, for accurate codec detection)
- [TMDB API Key](https://www.themoviedb.org/settings/api) (optional, for metadata fetching)

## Installation

1. Clone the repository:
   ```powershell
   git clone https://github.com/kliatsko/librarylint.git
   ```

2. Run the script:
   ```powershell
   .\LibraryLint.ps1
   ```

## Usage

### Interactive Mode
Simply run the script and follow the prompts:
```powershell
.\LibraryLint.ps1
```

### With Verbose Output
```powershell
.\LibraryLint.ps1 -Verbose
```

### With Custom Config File
```powershell
.\LibraryLint.ps1 -ConfigFile "C:\path\to\config.json"
```

## Main Menu Options

| Option | Description |
|--------|-------------|
| 1 | **Process Movies** - Full movie library cleanup and organization |
| 2 | **Process TV Shows** - Organize episodes into season folders |
| 3 | **Health Check** - Scan library for issues |
| 4 | **Codec Analysis** - Analyze codecs and generate transcode queue |
| 5 | **TMDB Metadata Fetch** - Download metadata from TMDB |
| 6 | **Export Library Report** - Generate CSV/HTML/JSON reports |
| 7 | **Enhanced Duplicate Scan** - Find duplicates using file hashing |
| 8 | **Undo Previous Session** - Rollback changes from a previous run |
| 9 | **Configuration Management** - Save/load/reset settings |

## Supported Formats

### Video
`.mp4`, `.mkv`, `.avi`, `.mov`, `.wmv`, `.flv`, `.m4v`

### Subtitles
`.srt`, `.sub`, `.idx`, `.ass`, `.ssa`, `.vtt`

### Archives
`.rar`, `.zip`, `.7z`, `.tar`, `.gz`, `.bz2`

## Configuration

Settings are stored in `%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json`

Key configuration options:
- `DryRun` - Preview mode (no changes made)
- `KeepSubtitles` - Keep subtitle files
- `KeepTrailers` - Move trailers to `_Trailers` folder
- `PreferredSubtitleLanguages` - Languages to keep (default: English)
- `GenerateNFO` - Auto-generate Kodi NFO files
- `TMDBApiKey` - Your TMDB API key
- `RetryCount` - Number of retries for failed operations
- `EnableUndo` - Enable undo manifest creation

## File Locations

| File Type | Location |
|-----------|----------|
| Logs | `%LOCALAPPDATA%\LibraryLint\Logs\` |
| Config | `%LOCALAPPDATA%\LibraryLint\LibraryLint.config.json` |
| Undo Manifests | `%LOCALAPPDATA%\LibraryLint\Undo\` |

## Quality Scoring

When detecting duplicates, files are scored based on:
- **Resolution**: 2160p (100) > 1080p (80) > 720p (60) > 480p (40)
- **Source**: BluRay (50) > WEB-DL (40) > WEBRip (35) > HDRip (30)
- **Codec**: x265/HEVC (30) > x264 (20) > XviD (10)
- **Audio**: Atmos (25) > DTS-HD (20) > TrueHD (20) > DTS (15) > AC3 (10)
- **HDR**: +20 bonus points

## Screenshots

### Processing Summary
```
╔══════════════════════════════════════════════════════════════╗
║                    PROCESSING SUMMARY                        ║
╠══════════════════════════════════════════════════════════════╣
║  Duration:              00:02:34                             ║
║  Files Deleted:         47                                   ║
║  Space Reclaimed:       2.3 GB                               ║
║  Archives Extracted:    12                                   ║
║  Folders Created:       8                                    ║
║  Folders Renamed:       23                                   ║
╚══════════════════════════════════════════════════════════════╝
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Nick Kliatsko**

## Acknowledgments

- [7-Zip](https://www.7-zip.org/) for archive extraction
- [MediaInfo](https://mediaarea.net/en/MediaInfo) for codec detection
- [The Movie Database (TMDB)](https://www.themoviedb.org/) for metadata
