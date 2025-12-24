<#
.SYNOPSIS
    MediaCleaner - Automated media library organization and cleanup tool

.DESCRIPTION
    This PowerShell script automates the cleanup and organization of downloaded media files (movies and TV shows).
    It performs the following operations:

    FOR MOVIES:
    - Extracts archives (.rar, .zip, .7z, .tar, .gz, .bz2)
    - Removes unnecessary files (samples, proofs, screenshots)
    - Processes trailers: Move to _Trailers folder or delete
    - Processes subtitles: Keep preferred language (English), delete others
    - Creates individual folders for loose video files
    - Cleans folder names by removing quality/codec/release tags
    - Formats movie years with parentheses (e.g., "Movie 2024" -> "Movie (2024)")
    - Replaces dots with spaces in folder names
    - Generates Kodi-compatible NFO files (optional)
    - Parses existing NFO files for metadata display

    FOR TV SHOWS:
    - Extracts all archives
    - Removes unnecessary files (samples, proofs, screenshots)
    - Processes subtitles: Keep preferred language
    - Parses episode info (S01E01, 1x01, multi-episode S01E01-E03)
    - Organizes episodes into Season folders
    - Renames episodes to standard format (optional)
    - Shows episode summary with gap detection
    - Removes empty folders

    FEATURES:
    - Dry-run mode: Preview all changes before applying them
    - Comprehensive logging: All operations logged with timestamps
    - Error handling: Graceful handling of locked files and permission issues
    - Progress tracking: Visual progress indicators for long operations
    - Multi-format support: Handles mp4, mkv, avi, mov, wmv, flv, m4v
    - Subtitle handling: Keep English subtitles (.srt, .sub, .idx, .ass, .ssa, .vtt)
    - Trailer management: Move trailers to _Trailers folder instead of deleting
    - NFO support: Read existing and generate new Kodi NFO files
    - Duplicate detection: Find and remove duplicate movies with quality scoring
    - Quality scoring: Score files by resolution, codec, source, audio, HDR
    - Health check: Validate library for issues (empty folders, missing files, etc.)
    - Codec analysis: Analyze video codecs and generate transcoding queue
    - TMDB integration: Fetch movie/show metadata from The Movie Database
    - Automatic 7-Zip installation if not present
    - Statistics summary: Detailed report of all operations performed

    NEW IN v4.1:
    - Configuration file: Save/load settings to JSON file
    - MediaInfo integration: Accurate codec detection from file headers
    - Undo/rollback support: Manifest-based rollback of changes
    - Enhanced duplicate detection: File hashing for exact duplicates
    - Export reports: CSV, HTML, and JSON library exports
    - Retry logic: Automatic retry of failed operations with backoff
    - Verbose mode: Optional debug output with -Verbose flag

.PARAMETER None
    This script is interactive and prompts for all necessary inputs

.EXAMPLE
    .\MediaCleaner.ps1
    Runs the script interactively, prompting for dry-run mode and folder selection

.NOTES
    Created By: Nick Kliatsko
    Last Updated: 12/13/2024
    Version: 4.1

    Requirements:
    - Windows 10 or later
    - PowerShell 5.1 or later
    - 7-Zip (will be installed automatically if not present)

    Supported Video Formats:
    - .mp4, .mkv, .avi, .mov, .wmv, .flv, .m4v

    Supported Archive Formats:
    - .rar, .zip, .7z, .tar, .gz, .bz2

    Log File Location:
    - Same directory as the script, named MediaCleaner_YYYYMMDD_HHMMSS.log

.LINK
    https://github.com/yourusername/MediaCleaner
#>

#============================================
# INITIALIZATION & CONFIGURATION
#============================================

# Script parameters for command-line usage
param(
    [string]$ConfigFile = $null,
    [switch]$Verbose
)

# Verbose mode flag
$script:VerboseMode = $Verbose.IsPresent

function Write-Verbose-Message {
    param([string]$Message, [string]$Color = "Magenta")
    if ($script:VerboseMode) {
        Write-Host "[DEBUG] $Message" -ForegroundColor $Color
    }
}

Write-Verbose-Message "Script starting..."
Write-Verbose-Message "Loading Windows Forms assembly..."
Add-Type -AssemblyName System.Windows.Forms
Write-Verbose-Message "Windows Forms loaded successfully"

# AppData folder for logs, config, and undo manifests (won't be uploaded to GitHub)
$script:AppDataFolder = Join-Path $env:LOCALAPPDATA "MediaCleaner"
$script:LogsFolder = Join-Path $script:AppDataFolder "Logs"
$script:UndoFolder = Join-Path $script:AppDataFolder "Undo"

# Create folders if they don't exist
@($script:AppDataFolder, $script:LogsFolder, $script:UndoFolder) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -Path $_ -ItemType Directory -Force | Out-Null
    }
}

# Configuration file path
$script:ConfigFilePath = if ($ConfigFile) { $ConfigFile } else { Join-Path $script:AppDataFolder "MediaCleaner.config.json" }

# Default configuration
$script:DefaultConfig = @{
    LogFile = Join-Path $script:LogsFolder "MediaCleaner_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    DryRun = $false
    SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
    MediaInfoPath = "C:\Program Files\MediaInfo\MediaInfo.exe"
    FFmpegPath = "C:\Program Files\ffmpeg\bin\ffmpeg.exe"
    VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v')
    SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    ArchiveExtensions = @('*.rar', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
    ArchiveCleanupPatterns = @('*.r??', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
    UnnecessaryPatterns = @('*Sample*', '*Proof*', '*Screens*')
    TrailerPatterns = @('*Trailer*', '*trailer*', '*TRAILER*', '*Teaser*', '*teaser*')
    PreferredSubtitleLanguages = @('eng', 'en', 'english')
    KeepSubtitles = $true
    KeepTrailers = $true
    GenerateNFO = $false
    OrganizeSeasons = $true
    RenameEpisodes = $false
    CheckDuplicates = $false
    TMDBApiKey = $null
    EnableParallelProcessing = $true
    MaxParallelJobs = 4
    EnableUndo = $true
    RetryCount = 3
    RetryDelaySeconds = 2
    Tags = @(
        # Resolution tags
        '1080p', '2160p', '720p', '480p', '4K', '2K', 'UHD',
        # Video quality/source tags
        'HDRip', 'DVDRip', 'BRRip', 'BR-Rip', 'BDRip', 'BD-Rip', 'WEB-DL', 'WEBRip', 'BluRay', 'DVDR', 'DVDScr',
        # Codec tags
        'x264', 'x265', 'X265', 'H264', 'H265', 'HEVC', 'hevc-d3g', 'XviD', 'DivX', 'AVC',
        # Audio tags
        'AAC', 'AC3', 'DTS', 'Atmos', 'TrueHD', 'DD5.1', '5.1', '7.1', 'DTS-HD',
        # HDR/Color tags
        'HDR', 'HDR10', 'HDR10+', 'DolbyVision', 'SDR', '10bit', '8bit',
        # Release type tags
        'Extended', 'Unrated', 'Remastered', 'REPACK', 'PROPER', 'iNTERNAL', 'LiMiTED', 'REAL', 'HC', 'ExtCut',
        'Anniversary Edition', 'Restored', "Director's Cut", 'Theatrical', 'DUBBED', 'SUBBED',
        # Language/Region tags
        'MULTi', 'DUAL', 'ENG', 'MULTI.VFF', 'TRUEFRENCH', 'FRENCH',
        # Release group examples
        'YIFY', 'YTS', 'RARBG', 'SPARKS', 'AMRAP', 'CMRG', 'FGT', 'EVO', 'STUTTERSHIT', 'FLEET', 'ION10',
        # Misc tags
        '1080-hd4u', 'NF', 'AMZN', 'HULU', 'WEB'
    )
}

# Initialize Config from defaults
$script:Config = $script:DefaultConfig.Clone()

#============================================
# CONFIGURATION FILE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Loads configuration from JSON file
.PARAMETER Path
    Path to the configuration file
.OUTPUTS
    Boolean indicating success
#>
function Import-Configuration {
    param(
        [string]$Path = $script:ConfigFilePath
    )

    if (Test-Path $Path) {
        try {
            Write-Verbose-Message "Loading configuration from: $Path"
            $jsonConfig = Get-Content -Path $Path -Raw | ConvertFrom-Json

            # Merge loaded config with defaults (loaded values override defaults)
            foreach ($property in $jsonConfig.PSObject.Properties) {
                if ($script:Config.ContainsKey($property.Name)) {
                    # Handle arrays specially
                    if ($property.Value -is [System.Array] -or $property.Value -is [System.Collections.ArrayList]) {
                        $script:Config[$property.Name] = @($property.Value)
                    } else {
                        $script:Config[$property.Name] = $property.Value
                    }
                }
            }

            # Ensure LogFile is unique for each session
            $script:Config.LogFile = Join-Path $PSScriptRoot "MediaCleaner_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

            Write-Host "Configuration loaded from: $Path" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Warning: Could not load config file: $_" -ForegroundColor Yellow
            Write-Host "Using default configuration" -ForegroundColor Cyan
            return $false
        }
    } else {
        Write-Verbose-Message "No configuration file found at: $Path"
        return $false
    }
}

<#
.SYNOPSIS
    Saves current configuration to JSON file
.PARAMETER Path
    Path to save the configuration file
#>
function Export-Configuration {
    param(
        [string]$Path = $script:ConfigFilePath
    )

    try {
        # Create a clean config object for export (exclude session-specific values)
        $exportConfig = @{}
        foreach ($key in $script:Config.Keys) {
            if ($key -ne 'LogFile') {  # Don't save session-specific log file
                $exportConfig[$key] = $script:Config[$key]
            }
        }

        $exportConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Configuration saved to: $Path" -ForegroundColor Green
        Write-Log "Configuration saved to: $Path" "INFO"
        return $true
    }
    catch {
        Write-Host "Error saving configuration: $_" -ForegroundColor Red
        Write-Log "Error saving configuration: $_" "ERROR"
        return $false
    }
}

<#
.SYNOPSIS
    Prompts user to save current configuration
#>
function Invoke-ConfigurationSavePrompt {
    $saveConfig = Read-Host "`nSave current settings as defaults? (Y/N) [N]"
    if ($saveConfig -eq 'Y' -or $saveConfig -eq 'y') {
        Export-Configuration
    }
}

# Load configuration file if it exists
Import-Configuration | Out-Null

# Global statistics tracking
$script:Stats = @{
    StartTime = $null
    EndTime = $null
    FilesDeleted = 0
    BytesDeleted = 0
    ArchivesExtracted = 0
    ArchivesFailed = 0
    FoldersCreated = 0
    FoldersRenamed = 0
    FilesMoved = 0
    EmptyFoldersRemoved = 0
    SubtitlesProcessed = 0
    SubtitlesDeleted = 0
    TrailersMoved = 0
    NFOFilesCreated = 0
    NFOFilesRead = 0
    Errors = 0
    Warnings = 0
    OperationsRetried = 0
}

# Undo manifest for rollback support
$script:UndoManifest = @{
    SessionId = [guid]::NewGuid().ToString()
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Operations = @()
}

# Failed operations queue for retry
$script:FailedOperations = @()

#============================================
# HELPER FUNCTIONS
#============================================

<#
.SYNOPSIS
    Writes a message to the log file with timestamp and severity level
.PARAMETER Message
    The message to log
.PARAMETER Level
    The severity level (INFO, WARNING, ERROR, DRY-RUN)
#>
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:Config.LogFile -Value $logMessage

    # Track errors and warnings in stats
    if ($Level -eq "ERROR") { $script:Stats.Errors++ }
    if ($Level -eq "WARNING") { $script:Stats.Warnings++ }
}

#============================================
# UNDO/ROLLBACK FUNCTIONS
#============================================

<#
.SYNOPSIS
    Records an operation to the undo manifest for potential rollback
.PARAMETER OperationType
    Type of operation (Move, Rename, Delete, Create)
.PARAMETER SourcePath
    Original path/name
.PARAMETER DestinationPath
    New path/name (if applicable)
.PARAMETER FileSize
    Size of the file (for deleted files tracking)
#>
function Add-UndoOperation {
    param(
        [string]$OperationType,
        [string]$SourcePath,
        [string]$DestinationPath = $null,
        [long]$FileSize = 0
    )

    if (-not $script:Config.EnableUndo) { return }

    $operation = @{
        Type = $OperationType
        Source = $SourcePath
        Destination = $DestinationPath
        Size = $FileSize
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    $script:UndoManifest.Operations += $operation
    Write-Verbose-Message "Recorded undo operation: $OperationType - $SourcePath"
}

<#
.SYNOPSIS
    Saves the undo manifest to a JSON file
.PARAMETER Path
    Optional path for the manifest file
#>
function Save-UndoManifest {
    param(
        [string]$Path = $null
    )

    if (-not $script:Config.EnableUndo) { return }
    if ($script:UndoManifest.Operations.Count -eq 0) {
        Write-Verbose-Message "No operations to save in undo manifest"
        return
    }

    if (-not $Path) {
        $Path = Join-Path $script:UndoFolder "MediaCleaner_Undo_$($script:UndoManifest.SessionId.Substring(0,8)).json"
    }

    try {
        $script:UndoManifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Undo manifest saved to: $Path" -ForegroundColor Cyan
        Write-Log "Undo manifest saved with $($script:UndoManifest.Operations.Count) operations" "INFO"
    }
    catch {
        Write-Host "Warning: Could not save undo manifest: $_" -ForegroundColor Yellow
        Write-Log "Error saving undo manifest: $_" "WARNING"
    }
}

<#
.SYNOPSIS
    Loads and executes an undo manifest to rollback changes
.PARAMETER ManifestPath
    Path to the undo manifest JSON file
#>
function Invoke-UndoOperations {
    param(
        [string]$ManifestPath
    )

    if (-not (Test-Path $ManifestPath)) {
        Write-Host "Undo manifest not found: $ManifestPath" -ForegroundColor Red
        return
    }

    try {
        $manifest = Get-Content -Path $ManifestPath -Raw | ConvertFrom-Json
        Write-Host "`nUndo Manifest: $($manifest.SessionId)" -ForegroundColor Cyan
        Write-Host "Created: $($manifest.Timestamp)" -ForegroundColor Gray
        Write-Host "Operations: $($manifest.Operations.Count)" -ForegroundColor Gray

        $confirm = Read-Host "`nUndo all operations? (Y/N) [N]"
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            Write-Host "Undo cancelled" -ForegroundColor Yellow
            return
        }

        # Process operations in reverse order
        $reversed = $manifest.Operations | Sort-Object { $_.Timestamp } -Descending
        $undone = 0
        $failed = 0

        foreach ($op in $reversed) {
            try {
                switch ($op.Type) {
                    "Move" {
                        if (Test-Path $op.Destination) {
                            Move-Item -Path $op.Destination -Destination $op.Source -Force -ErrorAction Stop
                            Write-Host "Restored: $($op.Source)" -ForegroundColor Green
                            $undone++
                        } else {
                            Write-Host "Cannot restore (destination missing): $($op.Destination)" -ForegroundColor Yellow
                            $failed++
                        }
                    }
                    "Rename" {
                        if (Test-Path $op.Destination) {
                            Rename-Item -Path $op.Destination -NewName (Split-Path $op.Source -Leaf) -Force -ErrorAction Stop
                            Write-Host "Renamed back: $(Split-Path $op.Source -Leaf)" -ForegroundColor Green
                            $undone++
                        } else {
                            Write-Host "Cannot restore (file missing): $($op.Destination)" -ForegroundColor Yellow
                            $failed++
                        }
                    }
                    "Create" {
                        if (Test-Path $op.Source) {
                            Remove-Item -Path $op.Source -Recurse -Force -ErrorAction Stop
                            Write-Host "Removed created item: $($op.Source)" -ForegroundColor Green
                            $undone++
                        }
                    }
                    "Delete" {
                        Write-Host "Cannot restore deleted file: $($op.Source)" -ForegroundColor Yellow
                        $failed++
                    }
                }
            }
            catch {
                Write-Host "Failed to undo: $($op.Source) - $_" -ForegroundColor Red
                $failed++
            }
        }

        Write-Host "`nUndo complete: $undone restored, $failed failed" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error processing undo manifest: $_" -ForegroundColor Red
    }
}

#============================================
# RETRY LOGIC FUNCTIONS
#============================================

<#
.SYNOPSIS
    Adds a failed operation to the retry queue
.PARAMETER OperationType
    Type of operation that failed
.PARAMETER Parameters
    Hashtable of parameters for the operation
.PARAMETER ErrorMessage
    The error message from the failure
#>
function Add-FailedOperation {
    param(
        [string]$OperationType,
        [hashtable]$Parameters,
        [string]$ErrorMessage
    )

    $script:FailedOperations += @{
        Type = $OperationType
        Parameters = $Parameters
        Error = $ErrorMessage
        RetryCount = 0
        Timestamp = Get-Date
    }
}

<#
.SYNOPSIS
    Retries all failed operations
.DESCRIPTION
    Attempts to retry failed operations with exponential backoff
#>
function Invoke-RetryFailedOperations {
    if ($script:FailedOperations.Count -eq 0) {
        return
    }

    Write-Host "`nRetrying $($script:FailedOperations.Count) failed operation(s)..." -ForegroundColor Yellow
    Write-Log "Starting retry of $($script:FailedOperations.Count) failed operations" "INFO"

    $stillFailed = @()

    foreach ($op in $script:FailedOperations) {
        if ($op.RetryCount -ge $script:Config.RetryCount) {
            Write-Host "Max retries reached for: $($op.Type) - $($op.Parameters.Path)" -ForegroundColor Red
            $stillFailed += $op
            continue
        }

        $op.RetryCount++
        $delay = $script:Config.RetryDelaySeconds * [math]::Pow(2, $op.RetryCount - 1)
        Write-Host "Retry $($op.RetryCount)/$($script:Config.RetryCount) (waiting ${delay}s): $($op.Type)" -ForegroundColor Cyan

        Start-Sleep -Seconds $delay

        try {
            switch ($op.Type) {
                "Move" {
                    Move-Item -Path $op.Parameters.Source -Destination $op.Parameters.Destination -Force -ErrorAction Stop
                    Write-Host "Retry successful: Move $($op.Parameters.Source)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Delete" {
                    Remove-Item -Path $op.Parameters.Path -Recurse -Force -ErrorAction Stop
                    Write-Host "Retry successful: Delete $($op.Parameters.Path)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Rename" {
                    Rename-Item -Path $op.Parameters.Path -NewName $op.Parameters.NewName -Force -ErrorAction Stop
                    Write-Host "Retry successful: Rename $($op.Parameters.Path)" -ForegroundColor Green
                    $script:Stats.OperationsRetried++
                }
                "Extract" {
                    # Re-attempt extraction
                    $process = Start-Process -FilePath $script:Config.SevenZipPath `
                        -ArgumentList "x", "-o`"$($op.Parameters.ExtractPath)`"", "`"$($op.Parameters.ArchivePath)`"", "-r", "-y" `
                        -NoNewWindow -PassThru -Wait
                    if ($process.ExitCode -eq 0) {
                        Write-Host "Retry successful: Extract $($op.Parameters.ArchivePath)" -ForegroundColor Green
                        $script:Stats.OperationsRetried++
                    } else {
                        throw "Extraction failed with exit code: $($process.ExitCode)"
                    }
                }
            }
        }
        catch {
            Write-Host "Retry failed: $_" -ForegroundColor Yellow
            $op.Error = $_.ToString()
            $stillFailed += $op
        }
    }

    $script:FailedOperations = $stillFailed

    if ($stillFailed.Count -gt 0) {
        Write-Host "$($stillFailed.Count) operation(s) still failed after retries" -ForegroundColor Yellow
        Write-Log "$($stillFailed.Count) operations failed after all retries" "WARNING"
    } else {
        Write-Host "All retry operations completed successfully" -ForegroundColor Green
    }
}

#============================================
# MEDIAINFO INTEGRATION
#============================================

<#
.SYNOPSIS
    Checks if MediaInfo CLI is installed
.OUTPUTS
    Boolean indicating if MediaInfo is available
#>
function Test-MediaInfoInstallation {
    if (Test-Path $script:Config.MediaInfoPath) {
        Write-Verbose-Message "MediaInfo found at: $($script:Config.MediaInfoPath)"
        return $true
    }

    # Try to find in PATH
    $mediaInfoInPath = Get-Command "mediainfo" -ErrorAction SilentlyContinue
    if ($mediaInfoInPath) {
        $script:Config.MediaInfoPath = $mediaInfoInPath.Source
        Write-Verbose-Message "MediaInfo found in PATH: $($script:Config.MediaInfoPath)"
        return $true
    }

    Write-Verbose-Message "MediaInfo not found - using filename parsing for codec detection"
    return $false
}

<#
.SYNOPSIS
    Checks if FFmpeg is installed and available
.OUTPUTS
    Boolean indicating if FFmpeg is available
#>
function Test-FFmpegInstallation {
    if (Test-Path $script:Config.FFmpegPath) {
        Write-Verbose-Message "FFmpeg found at: $($script:Config.FFmpegPath)"
        return $true
    }

    # Try to find in PATH
    $ffmpegInPath = Get-Command "ffmpeg" -ErrorAction SilentlyContinue
    if ($ffmpegInPath) {
        $script:Config.FFmpegPath = $ffmpegInPath.Source
        Write-Verbose-Message "FFmpeg found in PATH: $($script:Config.FFmpegPath)"
        return $true
    }

    # Check common installation locations
    $commonPaths = @(
        "C:\ffmpeg\bin\ffmpeg.exe",
        "C:\Program Files\ffmpeg\bin\ffmpeg.exe",
        "C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe",
        "$env:USERPROFILE\ffmpeg\bin\ffmpeg.exe",
        "$env:LOCALAPPDATA\ffmpeg\bin\ffmpeg.exe"
    )

    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            $script:Config.FFmpegPath = $path
            Write-Verbose-Message "FFmpeg found at: $path"
            return $true
        }
    }

    Write-Verbose-Message "FFmpeg not found"
    return $false
}

<#
.SYNOPSIS
    Gets detailed video information using MediaInfo CLI
.PARAMETER FilePath
    Path to the video file
.OUTPUTS
    Hashtable with detailed video properties
#>
function Get-MediaInfoDetails {
    param(
        [string]$FilePath
    )

    $info = @{
        VideoCodec = "Unknown"
        AudioCodec = "Unknown"
        Resolution = "Unknown"
        Width = 0
        Height = 0
        Duration = 0
        Bitrate = 0
        FrameRate = 0
        HDR = $false
        AudioChannels = 0
        Container = "Unknown"
    }

    if (-not (Test-MediaInfoInstallation)) {
        return $null
    }

    try {
        # Get video stream info
        $videoCodec = & $script:Config.MediaInfoPath --Inform="Video;%Format%" "$FilePath" 2>$null
        $width = & $script:Config.MediaInfoPath --Inform="Video;%Width%" "$FilePath" 2>$null
        $height = & $script:Config.MediaInfoPath --Inform="Video;%Height%" "$FilePath" 2>$null
        $duration = & $script:Config.MediaInfoPath --Inform="General;%Duration%" "$FilePath" 2>$null
        $bitrate = & $script:Config.MediaInfoPath --Inform="General;%OverallBitRate%" "$FilePath" 2>$null
        $frameRate = & $script:Config.MediaInfoPath --Inform="Video;%FrameRate%" "$FilePath" 2>$null
        $hdrFormat = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format%" "$FilePath" 2>$null
        $colorSpace = & $script:Config.MediaInfoPath --Inform="Video;%colour_primaries%" "$FilePath" 2>$null

        # Get audio stream info
        $audioCodec = & $script:Config.MediaInfoPath --Inform="Audio;%Format%" "$FilePath" 2>$null
        $audioChannels = & $script:Config.MediaInfoPath --Inform="Audio;%Channels%" "$FilePath" 2>$null

        # Get container info
        $container = & $script:Config.MediaInfoPath --Inform="General;%Format%" "$FilePath" 2>$null

        # Populate info
        if ($videoCodec) { $info.VideoCodec = $videoCodec.Trim() }
        if ($width) { $info.Width = [int]$width }
        if ($height) { $info.Height = [int]$height }
        if ($duration) { $info.Duration = [long]$duration }
        if ($bitrate) { $info.Bitrate = [long]$bitrate }
        if ($frameRate) { $info.FrameRate = [double]$frameRate }
        if ($audioCodec) { $info.AudioCodec = $audioCodec.Trim() }
        if ($audioChannels) { $info.AudioChannels = [int]$audioChannels }
        if ($container) { $info.Container = $container.Trim() }

        # Determine resolution label
        if ($info.Height -ge 2160) { $info.Resolution = "2160p" }
        elseif ($info.Height -ge 1080) { $info.Resolution = "1080p" }
        elseif ($info.Height -ge 720) { $info.Resolution = "720p" }
        elseif ($info.Height -ge 480) { $info.Resolution = "480p" }
        elseif ($info.Height -gt 0) { $info.Resolution = "$($info.Height)p" }

        # Check for HDR
        if ($hdrFormat -or $colorSpace -match 'BT.2020') {
            $info.HDR = $true
        }

        Write-Verbose-Message "MediaInfo: $($info.Resolution) $($info.VideoCodec) $($info.AudioCodec)"
        return $info
    }
    catch {
        Write-Verbose-Message "Error getting MediaInfo: $_" "Yellow"
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed HDR format information using MediaInfo CLI
.PARAMETER FilePath
    Path to the video file
.OUTPUTS
    Hashtable with HDR format details (Format, Profile, Compatibility)
.DESCRIPTION
    Extracts detailed HDR metadata including Dolby Vision profile,
    HDR10+ dynamic metadata, HLG, and compatibility layers
#>
function Get-MediaInfoHDRFormat {
    param(
        [string]$FilePath
    )

    if (-not (Test-MediaInfoInstallation)) {
        return $null
    }

    try {
        $hdrFormat = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format%" "$FilePath" 2>$null
        $hdrFormatCompat = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format_Compatibility%" "$FilePath" 2>$null
        $hdrFormatProfile = & $script:Config.MediaInfoPath --Inform="Video;%HDR_Format_Profile%" "$FilePath" 2>$null
        $colorPrimaries = & $script:Config.MediaInfoPath --Inform="Video;%colour_primaries%" "$FilePath" 2>$null
        $transferChar = & $script:Config.MediaInfoPath --Inform="Video;%transfer_characteristics%" "$FilePath" 2>$null

        $result = @{
            Format = $null
            Profile = $null
            Compatibility = $null
            ColorPrimaries = $null
            TransferCharacteristics = $null
        }

        if ($hdrFormat) { $result.Format = $hdrFormat.Trim() }
        if ($hdrFormatProfile) { $result.Profile = $hdrFormatProfile.Trim() }
        if ($hdrFormatCompat) { $result.Compatibility = $hdrFormatCompat.Trim() }
        if ($colorPrimaries) { $result.ColorPrimaries = $colorPrimaries.Trim() }
        if ($transferChar) { $result.TransferCharacteristics = $transferChar.Trim() }

        # Normalize HDR format names
        if ($result.Format) {
            if ($result.Format -match 'Dolby Vision|DOVI') {
                $result.Format = "Dolby Vision"
            }
            elseif ($result.Format -match 'HDR10\+|SMPTE ST 2094') {
                $result.Format = "HDR10+"
            }
            elseif ($result.Format -match 'SMPTE ST 2086|HDR10') {
                $result.Format = "HDR10"
            }
            elseif ($result.Format -match 'HLG|ARIB STD-B67') {
                $result.Format = "HLG"
            }
        }
        # Detect HDR from transfer characteristics if HDR_Format not present
        elseif ($transferChar -match 'PQ|SMPTE ST 2084') {
            $result.Format = "HDR10"
        }
        elseif ($transferChar -match 'HLG') {
            $result.Format = "HLG"
        }
        # Check for BT.2020 color space (wide color gamut, often indicates HDR)
        elseif ($colorPrimaries -match 'BT\.2020') {
            $result.Format = "HDR"
        }

        if ($result.Format) {
            return $result
        }
        return $null
    }
    catch {
        Write-Verbose-Message "Error getting HDR format: $_" "Yellow"
        return $null
    }
}

#============================================
# IMPROVED DUPLICATE DETECTION WITH HASHING
#============================================

<#
.SYNOPSIS
    Calculates a partial hash of a file for duplicate detection
.PARAMETER FilePath
    Path to the file
.PARAMETER SampleSize
    Number of bytes to sample from start and end (default 1MB each)
.OUTPUTS
    String hash value
.DESCRIPTION
    Instead of hashing the entire file (slow for large videos),
    this samples the first and last portions for a quick fingerprint
#>
function Get-FilePartialHash {
    param(
        [string]$FilePath,
        [int]$SampleSize = 1MB
    )

    try {
        $file = [System.IO.File]::OpenRead($FilePath)
        $fileLength = $file.Length

        # For small files, hash the entire file
        if ($fileLength -le ($SampleSize * 2)) {
            $file.Close()
            return (Get-FileHash -Path $FilePath -Algorithm MD5).Hash
        }

        $hasher = [System.Security.Cryptography.MD5]::Create()

        # Read first chunk
        $buffer = New-Object byte[] $SampleSize
        $file.Read($buffer, 0, $SampleSize) | Out-Null

        # Seek to end and read last chunk
        $file.Seek(-$SampleSize, [System.IO.SeekOrigin]::End) | Out-Null
        $endBuffer = New-Object byte[] $SampleSize
        $file.Read($endBuffer, 0, $SampleSize) | Out-Null

        $file.Close()

        # Combine buffers and hash
        $combined = $buffer + $endBuffer + [System.BitConverter]::GetBytes($fileLength)
        $hashBytes = $hasher.ComputeHash($combined)
        $hash = [System.BitConverter]::ToString($hashBytes) -replace '-', ''

        return $hash
    }
    catch {
        Write-Verbose-Message "Error calculating hash for $FilePath : $_" "Yellow"
        return $null
    }
}

<#
.SYNOPSIS
    Enhanced duplicate detection using file hashing and size comparison
.PARAMETER Path
    Root path of the media library
.OUTPUTS
    Array of duplicate groups with hash information
#>
function Find-DuplicateMoviesEnhanced {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for duplicate movies (enhanced)..." -ForegroundColor Yellow
    Write-Log "Starting enhanced duplicate scan in: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return @()
        }

        # Build lookup by title AND by file hash
        $titleLookup = @{}
        $hashLookup = @{}
        $totalFolders = $movieFolders.Count
        $currentIndex = 0

        foreach ($folder in $movieFolders) {
            $currentIndex++
            $percentComplete = [math]::Round(($currentIndex / $totalFolders) * 100)
            Write-Progress -Activity "Scanning for duplicates (enhanced)" -Status "Processing $currentIndex of $totalFolders - $($folder.Name)" -PercentComplete $percentComplete

            # Find the main video file
            $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Sort-Object Length -Descending |
                Select-Object -First 1

            if (-not $videoFile) { continue }

            $titleInfo = Get-NormalizedTitle -Name $folder.Name
            $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

            # Calculate partial hash for exact duplicate detection
            $fileHash = Get-FilePartialHash -FilePath $videoFile.FullName

            $entry = @{
                Path = $folder.FullName
                OriginalName = $folder.Name
                Year = $titleInfo.Year
                Quality = $quality
                FileSize = $videoFile.Length
                FilePath = $videoFile.FullName
                FileHash = $fileHash
                MatchType = @()  # Will track how this was matched
            }

            # Add to title lookup
            $titleKey = $titleInfo.NormalizedTitle
            if ($titleInfo.Year) {
                $titleKey = "$($titleInfo.NormalizedTitle)|$($titleInfo.Year)"
            }

            if (-not $titleLookup.ContainsKey($titleKey)) {
                $titleLookup[$titleKey] = @()
            }
            $titleLookup[$titleKey] += $entry

            # Add to hash lookup (if hash was calculated)
            if ($fileHash) {
                if (-not $hashLookup.ContainsKey($fileHash)) {
                    $hashLookup[$fileHash] = @()
                }
                $hashLookup[$fileHash] += $entry
            }
        }

        Write-Progress -Activity "Scanning for duplicates (enhanced)" -Completed

        # Find duplicates
        $duplicates = @()
        $processedPaths = @{}

        # First, find exact duplicates by hash
        foreach ($hash in $hashLookup.Keys) {
            if ($hashLookup[$hash].Count -gt 1) {
                foreach ($entry in $hashLookup[$hash]) {
                    $entry.MatchType += "ExactHash"
                    $processedPaths[$entry.Path] = $true
                }
                $duplicates += ,@($hashLookup[$hash])
            }
        }

        # Then find title-based duplicates (excluding already-found exact matches)
        foreach ($key in $titleLookup.Keys) {
            $entries = $titleLookup[$key] | Where-Object { -not $processedPaths.ContainsKey($_.Path) }
            if ($entries.Count -gt 1) {
                foreach ($entry in $entries) {
                    $entry.MatchType += "TitleMatch"

                    # Check if file sizes are similar (within 10%)
                    $avgSize = ($entries | Measure-Object -Property FileSize -Average).Average
                    $sizeDiff = [math]::Abs($entry.FileSize - $avgSize) / $avgSize
                    if ($sizeDiff -lt 0.1) {
                        $entry.MatchType += "SimilarSize"
                    }
                }
                $duplicates += ,@($entries)
            }
        }

        Write-Log "Enhanced duplicate scan found $($duplicates.Count) duplicate groups" "INFO"
        return $duplicates
    }
    catch {
        Write-Host "Error scanning for duplicates: $_" -ForegroundColor Red
        Write-Log "Error in enhanced duplicate scan: $_" "ERROR"
        return @()
    }
}

<#
.SYNOPSIS
    Displays enhanced duplicate report with match type information
.PARAMETER Path
    Root path of the media library
#>
function Show-EnhancedDuplicateReport {
    param(
        [string]$Path
    )

    $duplicates = Find-DuplicateMoviesEnhanced -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "No duplicate movies found!" -ForegroundColor Green
        Write-Log "No duplicates found" "INFO"
        return
    }

    Write-Host "`n" -NoNewline
    Write-Host "+" + ("=" * 64) + "+" -ForegroundColor Yellow
    Write-Host "|" + " ENHANCED DUPLICATE REPORT ".PadLeft(44).PadRight(64) + "|" -ForegroundColor Yellow
    Write-Host "+" + ("=" * 64) + "+" -ForegroundColor Yellow
    Write-Host "|  Found $($duplicates.Count) potential duplicate group(s)".PadRight(64) + "|" -ForegroundColor Yellow
    Write-Host "+" + ("=" * 64) + "+" -ForegroundColor Yellow

    $groupNum = 1
    foreach ($group in $duplicates) {
        $matchTypes = ($group | ForEach-Object { $_.MatchType } | Sort-Object -Unique) -join ", "
        Write-Host "`n[$groupNum] Duplicates (Match: $matchTypes):" -ForegroundColor Cyan

        # Sort by quality score descending
        $sorted = $group | Sort-Object { $_.Quality.Score } -Descending

        $first = $true
        foreach ($movie in $sorted) {
            $sizeStr = Format-FileSize $movie.FileSize
            $scoreStr = "Score: $($movie.Quality.Score)"
            $hashStr = if ($movie.FileHash) { $movie.FileHash.Substring(0, 8) + "..." } else { "N/A" }

            if ($first) {
                Write-Host "  [KEEP] " -ForegroundColor Green -NoNewline
                $first = $false
            } else {
                Write-Host "  [DEL?] " -ForegroundColor Red -NoNewline
            }

            Write-Host "$($movie.OriginalName)" -ForegroundColor White
            Write-Host "         $scoreStr | $sizeStr | Hash: $hashStr" -ForegroundColor Gray
            Write-Host "         $($movie.Quality.Resolution) $($movie.Quality.Source) $($movie.Quality.Codec)" -ForegroundColor DarkGray
        }

        $groupNum++
    }

    Write-Host "`n" -NoNewline
    Write-Log "Enhanced duplicate report: $($duplicates.Count) groups found" "INFO"
}

#============================================
# EXPORT/REPORT FUNCTIONS
#============================================

<#
.SYNOPSIS
    Exports library data to CSV format
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output CSV file
.PARAMETER MediaType
    Type of media (Movies or TVShows)
#>
function Export-LibraryToCSV {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $PSScriptRoot "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }

    Write-Host "`nExporting library to CSV..." -ForegroundColor Yellow
    Write-Log "Starting CSV export for: $Path" "INFO"

    try {
        $items = @()

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            foreach ($folder in $folders) {
                $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $items += [PSCustomObject]@{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    FolderName = $folder.Name
                    FileName = $videoFile.Name
                    FileSizeMB = [math]::Round($videoFile.Length / 1MB, 2)
                    Resolution = $quality.Resolution
                    VideoCodec = $quality.Codec
                    AudioCodec = $quality.Audio
                    Source = $quality.Source
                    QualityScore = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                    Bitrate = if ($quality.Bitrate -gt 0) { [math]::Round($quality.Bitrate / 1000000, 1) } else { $null }
                    Container = $videoFile.Extension.TrimStart('.')
                    Path = $folder.FullName
                    DataSource = $quality.DataSource
                }
            }
        }
        else {
            # TV Shows
            $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

            foreach ($file in $videoFiles) {
                $epInfo = Get-EpisodeInfo -FileName $file.Name
                $quality = Get-QualityScore -FileName $file.Name -FilePath $file.FullName

                $items += [PSCustomObject]@{
                    ShowTitle = $epInfo.ShowTitle
                    Season = $epInfo.Season
                    Episode = $epInfo.Episode
                    EpisodeTitle = $epInfo.EpisodeTitle
                    FileName = $file.Name
                    FileSizeMB = [math]::Round($file.Length / 1MB, 2)
                    Resolution = $quality.Resolution
                    VideoCodec = $quality.Codec
                    AudioCodec = $quality.Audio
                    QualityScore = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                    Bitrate = if ($quality.Bitrate -gt 0) { [math]::Round($quality.Bitrate / 1000000, 1) } else { $null }
                    Container = $file.Extension.TrimStart('.')
                    Path = $file.FullName
                    DataSource = $quality.DataSource
                }
            }
        }

        if ($items.Count -gt 0) {
            $items | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported $($items.Count) items to: $OutputPath" -ForegroundColor Green
            Write-Log "CSV export completed: $($items.Count) items to $OutputPath" "INFO"
        } else {
            Write-Host "No items found to export" -ForegroundColor Yellow
        }

        return $OutputPath
    }
    catch {
        Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
        Write-Log "Error exporting to CSV: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates an HTML report of the media library
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output HTML file
.PARAMETER MediaType
    Type of media (Movies or TVShows)
#>
function Export-LibraryToHTML {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $PSScriptRoot "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }

    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    Write-Log "Starting HTML report for: $Path" "INFO"

    try {
        $items = @()
        $stats = @{
            TotalItems = 0
            TotalSizeGB = 0
            ByResolution = @{}
            ByCodec = @{}
            ByYear = @{}
        }

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            foreach ($folder in $folders) {
                $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $items += @{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    Resolution = $quality.Resolution
                    Codec = $quality.Codec
                    Size = $videoFile.Length
                    Score = $quality.Score
                    HDR = $quality.HDR
                    HDRFormat = $quality.HDRFormat
                }

                # Update stats
                $stats.TotalItems++
                $stats.TotalSizeGB += $videoFile.Length / 1GB

                if (-not $stats.ByResolution.ContainsKey($quality.Resolution)) { $stats.ByResolution[$quality.Resolution] = 0 }
                $stats.ByResolution[$quality.Resolution]++

                if (-not $stats.ByCodec.ContainsKey($quality.Codec)) { $stats.ByCodec[$quality.Codec] = 0 }
                $stats.ByCodec[$quality.Codec]++

                if ($titleInfo.Year) {
                    if (-not $stats.ByYear.ContainsKey($titleInfo.Year)) { $stats.ByYear[$titleInfo.Year] = 0 }
                    $stats.ByYear[$titleInfo.Year]++
                }
            }
        }

        # Generate HTML
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MediaCleaner Library Report</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #1a1a2e; color: #eee; padding: 20px; }
        h1 { color: #00d4ff; margin-bottom: 20px; text-align: center; }
        h2 { color: #00d4ff; margin: 20px 0 10px; border-bottom: 2px solid #00d4ff; padding-bottom: 5px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: #16213e; padding: 20px; border-radius: 10px; text-align: center; }
        .stat-value { font-size: 2.5em; color: #00d4ff; font-weight: bold; }
        .stat-label { color: #888; margin-top: 5px; }
        .chart-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .chart { background: #16213e; padding: 20px; border-radius: 10px; }
        .bar-chart { margin-top: 10px; }
        .bar-item { display: flex; align-items: center; margin: 8px 0; }
        .bar-label { width: 100px; font-size: 0.9em; }
        .bar-container { flex: 1; background: #0f3460; height: 25px; border-radius: 5px; overflow: hidden; }
        .bar { height: 100%; background: linear-gradient(90deg, #00d4ff, #0097b2); transition: width 0.5s; }
        .bar-value { margin-left: 10px; font-size: 0.9em; color: #888; }
        table { width: 100%; border-collapse: collapse; background: #16213e; border-radius: 10px; overflow: hidden; }
        th { background: #0f3460; color: #00d4ff; padding: 15px; text-align: left; }
        td { padding: 12px 15px; border-bottom: 1px solid #0f3460; }
        tr:hover { background: #1f4068; }
        .quality-high { color: #00ff88; }
        .quality-medium { color: #ffaa00; }
        .quality-low { color: #ff4444; }
        .search-box { width: 100%; padding: 12px; margin-bottom: 20px; background: #16213e; border: 2px solid #0f3460; border-radius: 8px; color: #eee; font-size: 1em; }
        .search-box:focus { outline: none; border-color: #00d4ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>MediaCleaner Library Report</h1>
    <p style="text-align: center; color: #888; margin-bottom: 30px;">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>

    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-value">$($stats.TotalItems)</div>
            <div class="stat-label">Total $MediaType</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$([math]::Round($stats.TotalSizeGB, 1)) GB</div>
            <div class="stat-label">Total Size</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($stats.ByResolution.Keys.Count)</div>
            <div class="stat-label">Resolution Types</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">$($stats.ByCodec.Keys.Count)</div>
            <div class="stat-label">Codec Types</div>
        </div>
    </div>

    <div class="chart-container">
        <div class="chart">
            <h2>By Resolution</h2>
            <div class="bar-chart">
"@

        # Add resolution bars
        $maxRes = ($stats.ByResolution.Values | Measure-Object -Maximum).Maximum
        foreach ($res in ($stats.ByResolution.GetEnumerator() | Sort-Object Value -Descending)) {
            $pct = if ($maxRes -gt 0) { [math]::Round(($res.Value / $maxRes) * 100) } else { 0 }
            $html += @"
                <div class="bar-item">
                    <span class="bar-label">$($res.Key)</span>
                    <div class="bar-container"><div class="bar" style="width: $pct%"></div></div>
                    <span class="bar-value">$($res.Value)</span>
                </div>
"@
        }

        $html += @"
            </div>
        </div>
        <div class="chart">
            <h2>By Codec</h2>
            <div class="bar-chart">
"@

        # Add codec bars
        $maxCodec = ($stats.ByCodec.Values | Measure-Object -Maximum).Maximum
        foreach ($codec in ($stats.ByCodec.GetEnumerator() | Sort-Object Value -Descending)) {
            $pct = if ($maxCodec -gt 0) { [math]::Round(($codec.Value / $maxCodec) * 100) } else { 0 }
            $html += @"
                <div class="bar-item">
                    <span class="bar-label">$($codec.Key)</span>
                    <div class="bar-container"><div class="bar" style="width: $pct%"></div></div>
                    <span class="bar-value">$($codec.Value)</span>
                </div>
"@
        }

        $html += @"
            </div>
        </div>
    </div>

    <h2>Library Contents</h2>
    <input type="text" class="search-box" placeholder="Search movies..." onkeyup="filterTable(this.value)">
    <table id="libraryTable">
        <thead>
            <tr>
                <th>Title</th>
                <th>Year</th>
                <th>Resolution</th>
                <th>Codec</th>
                <th>Size</th>
                <th>Score</th>
            </tr>
        </thead>
        <tbody>
"@

        # Add table rows
        foreach ($item in ($items | Sort-Object { $_.Title })) {
            $sizeStr = Format-FileSize $item.Size
            $scoreClass = if ($item.Score -ge 100) { "quality-high" } elseif ($item.Score -ge 50) { "quality-medium" } else { "quality-low" }
            $html += @"
            <tr>
                <td>$([System.Web.HttpUtility]::HtmlEncode($item.Title))</td>
                <td>$($item.Year)</td>
                <td>$($item.Resolution)</td>
                <td>$($item.Codec)</td>
                <td>$sizeStr</td>
                <td class="$scoreClass">$($item.Score)</td>
            </tr>
"@
        }

        $html += @"
        </tbody>
    </table>

    <div class="footer">
        <p>Generated by MediaCleaner v4.1</p>
    </div>

    <script>
        function filterTable(query) {
            const table = document.getElementById('libraryTable');
            const rows = table.getElementsByTagName('tr');
            query = query.toLowerCase();
            for (let i = 1; i < rows.length; i++) {
                const cells = rows[i].getElementsByTagName('td');
                let found = false;
                for (let j = 0; j < cells.length; j++) {
                    if (cells[j].textContent.toLowerCase().includes(query)) {
                        found = true;
                        break;
                    }
                }
                rows[i].style.display = found ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"@

        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "HTML report generated: $OutputPath" -ForegroundColor Green
        Write-Log "HTML report generated: $OutputPath" "INFO"

        # Optionally open in browser
        $openReport = Read-Host "Open report in browser? (Y/N) [Y]"
        if ($openReport -ne 'N' -and $openReport -ne 'n') {
            Start-Process $OutputPath
        }

        return $OutputPath
    }
    catch {
        Write-Host "Error generating HTML report: $_" -ForegroundColor Red
        Write-Log "Error generating HTML report: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Exports library data to JSON format for backup or integration
.PARAMETER Path
    Root path of the media library
.PARAMETER OutputPath
    Path for the output JSON file
#>
function Export-LibraryToJSON {
    param(
        [string]$Path,
        [string]$OutputPath = $null,
        [string]$MediaType = "Movies"
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path $PSScriptRoot "MediaLibrary_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    }

    Write-Host "`nExporting library to JSON..." -ForegroundColor Yellow

    try {
        $library = @{
            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            MediaType = $MediaType
            SourcePath = $Path
            Items = @()
        }

        if ($MediaType -eq "Movies") {
            $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne '_Trailers' }

            foreach ($folder in $folders) {
                $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Length -Descending |
                    Select-Object -First 1

                if (-not $videoFile) { continue }

                $titleInfo = Get-NormalizedTitle -Name $folder.Name
                $quality = Get-QualityScore -FileName $folder.Name -FilePath $videoFile.FullName

                $library.Items += @{
                    Title = $titleInfo.NormalizedTitle
                    Year = $titleInfo.Year
                    FolderName = $folder.Name
                    FileName = $videoFile.Name
                    FileSize = $videoFile.Length
                    Quality = $quality
                    Path = $folder.FullName
                }
            }
        }

        $library | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "Exported $($library.Items.Count) items to: $OutputPath" -ForegroundColor Green
        Write-Log "JSON export completed: $($library.Items.Count) items" "INFO"

        return $OutputPath
    }
    catch {
        Write-Host "Error exporting to JSON: $_" -ForegroundColor Red
        Write-Log "Error exporting to JSON: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Interactive export menu
.PARAMETER Path
    Root path of the media library
#>
function Invoke-ExportMenu {
    param(
        [string]$Path,
        [string]$MediaType = "Movies"
    )

    Write-Host "`n=== Export Library ===" -ForegroundColor Cyan
    Write-Host "1. Export to CSV (spreadsheet-compatible)"
    Write-Host "2. Export to HTML (visual report)"
    Write-Host "3. Export to JSON (backup/integration)"
    Write-Host "4. Export All Formats"
    Write-Host "5. Cancel"

    $choice = Read-Host "`nSelect export format"

    switch ($choice) {
        "1" { Export-LibraryToCSV -Path $Path -MediaType $MediaType }
        "2" { Export-LibraryToHTML -Path $Path -MediaType $MediaType }
        "3" { Export-LibraryToJSON -Path $Path -MediaType $MediaType }
        "4" {
            Export-LibraryToCSV -Path $Path -MediaType $MediaType
            Export-LibraryToHTML -Path $Path -MediaType $MediaType
            Export-LibraryToJSON -Path $Path -MediaType $MediaType
        }
        default { Write-Host "Export cancelled" -ForegroundColor Yellow }
    }
}

<#
.SYNOPSIS
    Displays a folder browser dialog for user to select a directory
.OUTPUTS
    String - The full path of the selected folder
#>
function Select-FolderDialog {
    param(
        [string]$Description = "Select the folder to clean up"
    )
    Write-Host "Opening folder dialog... (check taskbar if it doesn't appear)" -ForegroundColor Yellow
    Write-Verbose-Message "Opening folder browser dialog..."
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowNewFolderButton = $true

    $result = $folderBrowser.ShowDialog()
    Write-Verbose-Message "Dialog closed with result: $result"

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Verbose-Message "Selected path: $($folderBrowser.SelectedPath)"
        return $folderBrowser.SelectedPath
    } else {
        Write-Host "No folder selected" -ForegroundColor Red
        Write-Log "No folder selected by user" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Formats bytes into human-readable format
.PARAMETER Bytes
    The number of bytes to format
#>
function Format-FileSize {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

<#
.SYNOPSIS
    Displays the statistics summary at the end of processing
#>
function Show-Statistics {
    $script:Stats.EndTime = Get-Date
    $duration = $script:Stats.EndTime - $script:Stats.StartTime

    Write-Host "`n" -NoNewline
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    PROCESSING SUMMARY                        ║" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # Time stats
    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Duration:              " -NoNewline
    Write-Host ("{0:hh\:mm\:ss}" -f $duration).PadRight(39) -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # File operations
    if ($script:Stats.FilesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Files Deleted:         " -NoNewline
        Write-Host "$($script:Stats.FilesDeleted)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.BytesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Space Reclaimed:       " -NoNewline
        Write-Host (Format-FileSize $script:Stats.BytesDeleted).PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.ArchivesExtracted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Archives Extracted:    " -NoNewline
        Write-Host "$($script:Stats.ArchivesExtracted)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.ArchivesFailed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Archives Failed:       " -NoNewline
        Write-Host "$($script:Stats.ArchivesFailed)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FoldersCreated -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Folders Created:       " -NoNewline
        Write-Host "$($script:Stats.FoldersCreated)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FoldersRenamed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Folders Renamed:       " -NoNewline
        Write-Host "$($script:Stats.FoldersRenamed)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.FilesMoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Files Moved:           " -NoNewline
        Write-Host "$($script:Stats.FilesMoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.EmptyFoldersRemoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Empty Folders Removed: " -NoNewline
        Write-Host "$($script:Stats.EmptyFoldersRemoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.SubtitlesProcessed -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Subtitles Processed:   " -NoNewline
        Write-Host "$($script:Stats.SubtitlesProcessed)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.SubtitlesDeleted -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Subtitles Deleted:     " -NoNewline
        Write-Host "$($script:Stats.SubtitlesDeleted)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.TrailersMoved -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Trailers Moved:        " -NoNewline
        Write-Host "$($script:Stats.TrailersMoved)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.OperationsRetried -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "Operations Retried:    " -NoNewline
        Write-Host "$($script:Stats.OperationsRetried)".PadRight(39) -ForegroundColor Yellow -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.NFOFilesCreated -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "NFO Files Created:     " -NoNewline
        Write-Host "$($script:Stats.NFOFilesCreated)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    if ($script:Stats.NFOFilesRead -gt 0) {
        Write-Host "║  " -ForegroundColor Cyan -NoNewline
        Write-Host "NFO Files Parsed:      " -NoNewline
        Write-Host "$($script:Stats.NFOFilesRead)".PadRight(39) -ForegroundColor Green -NoNewline
        Write-Host "║" -ForegroundColor Cyan
    }

    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    # Errors and warnings
    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Errors:                " -NoNewline
    $errorColor = if ($script:Stats.Errors -gt 0) { "Red" } else { "Green" }
    Write-Host "$($script:Stats.Errors)".PadRight(39) -ForegroundColor $errorColor -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "║  " -ForegroundColor Cyan -NoNewline
    Write-Host "Warnings:              " -NoNewline
    $warningColor = if ($script:Stats.Warnings -gt 0) { "Yellow" } else { "Green" }
    Write-Host "$($script:Stats.Warnings)".PadRight(39) -ForegroundColor $warningColor -NoNewline
    Write-Host "║" -ForegroundColor Cyan

    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Log the summary
    Write-Log "Processing completed - Duration: $($duration.ToString('hh\:mm\:ss')), Files Deleted: $($script:Stats.FilesDeleted), Space Reclaimed: $(Format-FileSize $script:Stats.BytesDeleted), Archives Extracted: $($script:Stats.ArchivesExtracted), Errors: $($script:Stats.Errors)" "INFO"
}

#============================================
# 7-ZIP FUNCTIONS
#============================================

<#
.SYNOPSIS
    Verifies 7-Zip is installed, installs it if not present
.OUTPUTS
    Boolean - True if 7-Zip is available, False otherwise
#>
function Test-SevenZipInstallation {
    Write-Host "Checking 7-Zip installation..." -ForegroundColor Yellow

    if (Test-Path -Path $script:Config.SevenZipPath) {
        Write-Host "7-Zip installed" -ForegroundColor Green
        Write-Log "7-Zip found at $($script:Config.SevenZipPath)" "INFO"
        return $true
    }

    Write-Host "7-Zip not installed. Installing now..." -ForegroundColor Yellow
    Write-Log "7-Zip not found, attempting installation" "INFO"

    $installerPath = "$env:TEMP\7z-installer.exe"
    $7zipUrl = "https://www.7-zip.org/a/7z2408-x64.exe"

    try {
        Write-Host "Downloading 7-Zip..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $7zipUrl -OutFile $installerPath -UseBasicParsing

        Write-Host "Installing 7-Zip..." -ForegroundColor Yellow
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -NoNewWindow

        Remove-Item $installerPath -ErrorAction SilentlyContinue

        if (Test-Path $script:Config.SevenZipPath) {
            Write-Host "7-Zip installed successfully" -ForegroundColor Green
            Write-Log "7-Zip installed successfully" "INFO"
            return $true
        } else {
            Write-Host "7-Zip installation failed" -ForegroundColor Red
            Write-Log "7-Zip installation failed" "ERROR"
            return $false
        }
    }
    catch {
        Write-Host "Error installing 7-Zip: $_" -ForegroundColor Red
        Write-Host "Please install 7-Zip manually from https://www.7-zip.org/" -ForegroundColor Yellow
        Write-Log "Error installing 7-Zip: $_" "ERROR"
        return $false
    }
}

#============================================
# FILE CLEANUP FUNCTIONS
#============================================

<#
.SYNOPSIS
    Removes unnecessary files matching specified patterns
.PARAMETER Path
    The root path to search for unnecessary files
#>
function Remove-UnnecessaryFiles {
    param(
        [string]$Path
    )

    Write-Host "Cleaning unnecessary files..." -ForegroundColor Yellow
    Write-Log "Starting unnecessary file cleanup in: $Path" "INFO"

    try {
        $filesToDelete = @()

        foreach ($pattern in $script:Config.UnnecessaryPatterns) {
            $files = Get-ChildItem -Path $Path -Filter $pattern -Recurse -ErrorAction SilentlyContinue
            if ($files) {
                $filesToDelete += $files
            }
        }

        if ($filesToDelete.Count -gt 0) {
            Write-Host "Found $($filesToDelete.Count) unnecessary file(s)/folder(s)" -ForegroundColor Cyan

            foreach ($item in $filesToDelete) {
                $itemSize = if ($item.PSIsContainer) {
                    (Get-ChildItem -Path $item.FullName -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                } else {
                    $item.Length
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete: $($item.FullName)" -ForegroundColor Yellow
                    Write-Log "Would delete: $($item.FullName)" "DRY-RUN"
                } else {
                    Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "Deleted: $($item.Name)" -ForegroundColor Gray
                    Write-Log "Deleted: $($item.FullName)" "INFO"
                    $script:Stats.FilesDeleted++
                    $script:Stats.BytesDeleted += $itemSize
                }
            }
            Write-Host "Unnecessary files cleaned" -ForegroundColor Green
        } else {
            Write-Host "No unnecessary files found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Warning: Some files could not be deleted: $_" -ForegroundColor Yellow
        Write-Log "Error during file cleanup: $_" "ERROR"
    }
}

#============================================
# SUBTITLE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Processes subtitle files - keeps preferred languages, removes others
.PARAMETER Path
    The root path to search for subtitle files
.DESCRIPTION
    Scans for subtitle files (.srt, .sub, .idx, .ass, .ssa, .vtt) and either:
    - Keeps them if they match preferred languages or if KeepSubtitles is true
    - Deletes non-preferred language subtitles
    - Moves subtitle files to be alongside their video files
#>
function Invoke-SubtitleProcessing {
    param(
        [string]$Path
    )

    Write-Host "Processing subtitle files..." -ForegroundColor Yellow
    Write-Log "Starting subtitle processing in: $Path" "INFO"

    try {
        # Find all subtitle files
        $subtitleFiles = @()
        foreach ($ext in $script:Config.SubtitleExtensions) {
            $files = Get-ChildItem -Path $Path -Filter "*$ext" -Recurse -ErrorAction SilentlyContinue
            if ($files) {
                $subtitleFiles += $files
            }
        }

        if ($subtitleFiles.Count -eq 0) {
            Write-Host "No subtitle files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($subtitleFiles.Count) subtitle file(s)" -ForegroundColor Cyan

        foreach ($subtitle in $subtitleFiles) {
            $script:Stats.SubtitlesProcessed++

            # Check if subtitle matches preferred language
            $isPreferred = $false
            $subtitleNameLower = $subtitle.BaseName.ToLower()

            foreach ($lang in $script:Config.PreferredSubtitleLanguages) {
                if ($subtitleNameLower -match "\.$lang$" -or $subtitleNameLower -match "\.$lang\." -or $subtitleNameLower -match "_$lang$" -or $subtitleNameLower -match "_$lang[_\.]") {
                    $isPreferred = $true
                    break
                }
            }

            # If no language tag found, assume it's the default/preferred language
            $hasLanguageTag = $subtitleNameLower -match '\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german|ita|it|italian|por|pt|portuguese|rus|ru|russian|jpn|ja|japanese|chi|zh|chinese|kor|ko|korean|ara|ar|arabic|hin|hi|hindi|dut|nl|dutch|pol|pl|polish|swe|sv|swedish|nor|no|norwegian|dan|da|danish|fin|fi|finnish)(\.|$|_)'
            if (-not $hasLanguageTag) {
                $isPreferred = $true  # No language tag = default language, keep it
            }

            if ($script:Config.KeepSubtitles -and $isPreferred) {
                Write-Host "Keeping subtitle: $($subtitle.Name)" -ForegroundColor Green
                Write-Log "Keeping subtitle: $($subtitle.FullName)" "INFO"
            } else {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete subtitle: $($subtitle.Name)" -ForegroundColor Yellow
                    Write-Log "Would delete subtitle: $($subtitle.FullName)" "DRY-RUN"
                } else {
                    $subtitleSize = $subtitle.Length
                    Remove-Item -Path $subtitle.FullName -Force -ErrorAction SilentlyContinue
                    Write-Host "Deleted subtitle: $($subtitle.Name)" -ForegroundColor Gray
                    Write-Log "Deleted subtitle: $($subtitle.FullName)" "INFO"
                    $script:Stats.SubtitlesDeleted++
                    $script:Stats.BytesDeleted += $subtitleSize
                }
            }
        }

        Write-Host "Subtitle processing completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing subtitles: $_" -ForegroundColor Red
        Write-Log "Error processing subtitles: $_" "ERROR"
    }
}

#============================================
# TRAILER FUNCTIONS
#============================================

<#
.SYNOPSIS
    Moves trailer files to a _Trailers subfolder instead of deleting them
.PARAMETER Path
    The root path to search for trailer files
.DESCRIPTION
    Finds trailer/teaser files and moves them to a _Trailers folder,
    preserving them for later viewing while keeping the main library clean
#>
function Move-TrailersToFolder {
    param(
        [string]$Path
    )

    Write-Host "Processing trailer files..." -ForegroundColor Yellow
    Write-Log "Starting trailer processing in: $Path" "INFO"

    try {
        $trailerFiles = @()

        foreach ($pattern in $script:Config.TrailerPatterns) {
            # Search for trailer video files
            foreach ($ext in $script:Config.VideoExtensions) {
                $files = Get-ChildItem -Path $Path -Filter "$pattern$ext" -Recurse -ErrorAction SilentlyContinue
                if ($files) {
                    $trailerFiles += $files
                }
            }
            # Also search for trailer folders
            $trailerFolders = Get-ChildItem -Path $Path -Directory -Filter $pattern -Recurse -ErrorAction SilentlyContinue
            if ($trailerFolders) {
                foreach ($folder in $trailerFolders) {
                    $videosInFolder = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                        Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }
                    if ($videosInFolder) {
                        $trailerFiles += $videosInFolder
                    }
                }
            }
        }

        # Remove duplicates
        $trailerFiles = $trailerFiles | Select-Object -Unique

        if ($trailerFiles.Count -eq 0) {
            Write-Host "No trailer files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($trailerFiles.Count) trailer file(s)" -ForegroundColor Cyan

        # Create _Trailers folder at root
        $trailersFolder = Join-Path $Path "_Trailers"

        if (-not $script:Config.DryRun) {
            if (-not (Test-Path $trailersFolder)) {
                New-Item -Path $trailersFolder -ItemType Directory -Force | Out-Null
                Write-Host "Created _Trailers folder" -ForegroundColor Green
                Write-Log "Created _Trailers folder at: $trailersFolder" "INFO"
                $script:Stats.FoldersCreated++
            }
        }

        foreach ($trailer in $trailerFiles) {
            $destPath = Join-Path $trailersFolder $trailer.Name

            # Handle duplicate names by adding a number
            $counter = 1
            while (Test-Path $destPath) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($trailer.Name)
                $extension = $trailer.Extension
                $destPath = Join-Path $trailersFolder "$baseName`_$counter$extension"
                $counter++
            }

            if ($script:Config.DryRun) {
                Write-Host "[DRY-RUN] Would move trailer: $($trailer.Name) -> _Trailers/" -ForegroundColor Yellow
                Write-Log "Would move trailer: $($trailer.FullName) to $destPath" "DRY-RUN"
            } else {
                try {
                    Move-Item -Path $trailer.FullName -Destination $destPath -Force -ErrorAction Stop
                    Write-Host "Moved trailer: $($trailer.Name)" -ForegroundColor Green
                    Write-Log "Moved trailer: $($trailer.FullName) to $destPath" "INFO"
                    $script:Stats.TrailersMoved++
                }
                catch {
                    Write-Host "Warning: Could not move trailer $($trailer.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error moving trailer $($trailer.FullName): $_" "WARNING"
                }
            }
        }

        Write-Host "Trailer processing completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing trailers: $_" -ForegroundColor Red
        Write-Log "Error processing trailers: $_" "ERROR"
    }
}

#============================================
# NFO FILE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Parses an existing NFO file and extracts metadata
.PARAMETER NfoPath
    The path to the NFO file to parse
.OUTPUTS
    Hashtable containing extracted metadata (title, year, plot, etc.)
.DESCRIPTION
    Reads Kodi-compatible NFO files and extracts metadata including:
    - Title, original title, sort title
    - Year, release date
    - Plot, tagline
    - Rating, votes
    - IMDB/TMDB IDs
    - Genre, studio, director, actors
#>
function Read-NFOFile {
    param(
        [string]$NfoPath
    )

    $metadata = @{
        Title = $null
        OriginalTitle = $null
        SortTitle = $null
        Year = $null
        Plot = $null
        Tagline = $null
        Rating = $null
        Votes = $null
        IMDBID = $null
        TMDBID = $null
        Genres = @()
        Studios = @()
        Directors = @()
        Actors = @()
        Runtime = $null
    }

    try {
        if (-not (Test-Path $NfoPath)) {
            Write-Log "NFO file not found: $NfoPath" "WARNING"
            return $null
        }

        [xml]$nfoContent = Get-Content -Path $NfoPath -Encoding UTF8 -ErrorAction Stop
        $script:Stats.NFOFilesRead++

        # Movie NFO
        if ($nfoContent.movie) {
            $movie = $nfoContent.movie
            $metadata.Title = $movie.title
            $metadata.OriginalTitle = $movie.originaltitle
            $metadata.SortTitle = $movie.sorttitle
            $metadata.Year = $movie.year
            $metadata.Plot = $movie.plot
            $metadata.Tagline = $movie.tagline
            $metadata.Rating = $movie.rating
            $metadata.Votes = $movie.votes
            $metadata.Runtime = $movie.runtime

            # Extract IDs
            if ($movie.uniqueid) {
                foreach ($id in $movie.uniqueid) {
                    if ($id.type -eq 'imdb') { $metadata.IMDBID = $id.'#text' }
                    if ($id.type -eq 'tmdb') { $metadata.TMDBID = $id.'#text' }
                }
            }
            # Fallback for older NFO format
            if (-not $metadata.IMDBID -and $movie.imdbid) { $metadata.IMDBID = $movie.imdbid }
            if (-not $metadata.IMDBID -and $movie.id) { $metadata.IMDBID = $movie.id }

            # Extract genres
            if ($movie.genre) {
                $metadata.Genres = @($movie.genre)
            }

            # Extract studios
            if ($movie.studio) {
                $metadata.Studios = @($movie.studio)
            }

            # Extract directors
            if ($movie.director) {
                $metadata.Directors = @($movie.director)
            }

            # Extract actors
            if ($movie.actor) {
                foreach ($actor in $movie.actor) {
                    $metadata.Actors += @{
                        Name = $actor.name
                        Role = $actor.role
                        Thumb = $actor.thumb
                    }
                }
            }

            Write-Log "Parsed NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }
        # TV Show NFO
        elseif ($nfoContent.tvshow) {
            $show = $nfoContent.tvshow
            $metadata.Title = $show.title
            $metadata.OriginalTitle = $show.originaltitle
            $metadata.Year = $show.year
            $metadata.Plot = $show.plot
            $metadata.Rating = $show.rating

            if ($show.uniqueid) {
                foreach ($id in $show.uniqueid) {
                    if ($id.type -eq 'imdb') { $metadata.IMDBID = $id.'#text' }
                    if ($id.type -eq 'tmdb') { $metadata.TMDBID = $id.'#text' }
                }
            }

            if ($show.genre) {
                $metadata.Genres = @($show.genre)
            }

            Write-Log "Parsed TV Show NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }
        # Episode NFO
        elseif ($nfoContent.episodedetails) {
            $episode = $nfoContent.episodedetails
            $metadata.Title = $episode.title
            $metadata.Plot = $episode.plot
            $metadata.Rating = $episode.rating
            $metadata.Season = $episode.season
            $metadata.Episode = $episode.episode
            $metadata.Aired = $episode.aired

            Write-Log "Parsed Episode NFO file: $NfoPath - Title: $($metadata.Title)" "INFO"
        }

        return $metadata
    }
    catch {
        Write-Log "Error parsing NFO file $NfoPath : $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates a Kodi-compatible NFO file for a movie
.PARAMETER VideoPath
    The path to the video file
.PARAMETER Title
    The movie title (if not provided, extracted from folder/file name)
.PARAMETER Year
    The movie year (if not provided, extracted from folder/file name)
.DESCRIPTION
    Creates a basic NFO file that Kodi can use to identify the movie.
    The NFO file is named the same as the video file with .nfo extension.
#>
function New-MovieNFO {
    param(
        [string]$VideoPath,
        [string]$Title = $null,
        [string]$Year = $null
    )

    try {
        $videoFile = Get-Item $VideoPath -ErrorAction Stop
        $nfoPath = Join-Path $videoFile.DirectoryName "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"

        # Skip if NFO already exists
        if (Test-Path $nfoPath) {
            Write-Host "NFO already exists: $($videoFile.Name)" -ForegroundColor Cyan
            Write-Log "NFO already exists, skipping: $nfoPath" "INFO"
            return
        }

        # Extract title and year from folder name if not provided
        if (-not $Title -or -not $Year) {
            $folderName = $videoFile.Directory.Name

            # Try to extract year in parentheses format: "Movie Name (2024)"
            if ($folderName -match '^(.+?)\s*\((\d{4})\)') {
                if (-not $Title) { $Title = $Matches[1].Trim() }
                if (-not $Year) { $Year = $Matches[2] }
            }
            # Try to extract year without parentheses: "Movie Name 2024"
            elseif ($folderName -match '^(.+?)\s+(\d{4})$') {
                if (-not $Title) { $Title = $Matches[1].Trim() }
                if (-not $Year) { $Year = $Matches[2] }
            }
            # Just use folder name as title
            else {
                if (-not $Title) { $Title = $folderName }
            }
        }

        # Clean up title
        $Title = $Title -replace '\.', ' '
        $Title = $Title -replace '\s+', ' '
        $Title = $Title.Trim()

        if ($script:Config.DryRun) {
            Write-Host "[DRY-RUN] Would create NFO for: $Title $(if($Year){"($Year)"})" -ForegroundColor Yellow
            Write-Log "Would create NFO for: $Title ($Year) at $nfoPath" "DRY-RUN"
            return
        }

        # Create NFO XML content
        $nfoContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<movie>
    <title>$([System.Security.SecurityElement]::Escape($Title))</title>
    $(if($Year){"<year>$Year</year>"})
    <plot></plot>
    <outline></outline>
    <tagline></tagline>
    <runtime></runtime>
    <thumb></thumb>
    <fanart></fanart>
    <mpaa></mpaa>
    <genre></genre>
    <studio></studio>
    <director></director>
</movie>
"@

        $nfoContent | Out-File -FilePath $nfoPath -Encoding UTF8 -Force
        Write-Host "Created NFO: $Title $(if($Year){"($Year)"})" -ForegroundColor Green
        Write-Log "Created NFO file: $nfoPath" "INFO"
        $script:Stats.NFOFilesCreated++
    }
    catch {
        Write-Host "Error creating NFO for $VideoPath : $_" -ForegroundColor Red
        Write-Log "Error creating NFO for $VideoPath : $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Generates NFO files for all movies in a folder that don't have them
.PARAMETER Path
    The root path of the movie library
#>
function Invoke-NFOGeneration {
    param(
        [string]$Path
    )

    if (-not $script:Config.GenerateNFO) {
        return
    }

    Write-Host "Generating NFO files for movies..." -ForegroundColor Yellow
    Write-Log "Starting NFO generation in: $Path" "INFO"

    try {
        # Get all movie folders (folders containing video files)
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

        foreach ($folder in $movieFolders) {
            # Skip the _Trailers folder
            if ($folder.Name -eq '_Trailers') {
                continue
            }

            # Find video files in this folder
            $videoFiles = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

            foreach ($video in $videoFiles) {
                New-MovieNFO -VideoPath $video.FullName
            }
        }

        # Also check for video files directly in root
        $rootVideos = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        foreach ($video in $rootVideos) {
            New-MovieNFO -VideoPath $video.FullName
        }

        Write-Host "NFO generation completed" -ForegroundColor Green
    }
    catch {
        Write-Host "Error during NFO generation: $_" -ForegroundColor Red
        Write-Log "Error during NFO generation: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Scans and displays metadata from existing NFO files
.PARAMETER Path
    The root path to scan for NFO files
#>
function Show-NFOMetadata {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for existing NFO files..." -ForegroundColor Yellow
    Write-Log "Scanning for NFO files in: $Path" "INFO"

    try {
        $nfoFiles = Get-ChildItem -Path $Path -Filter "*.nfo" -Recurse -ErrorAction SilentlyContinue

        if ($nfoFiles.Count -eq 0) {
            Write-Host "No NFO files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($nfoFiles.Count) NFO file(s)`n" -ForegroundColor Cyan

        foreach ($nfo in $nfoFiles) {
            $metadata = Read-NFOFile -NfoPath $nfo.FullName

            if ($metadata -and $metadata.Title) {
                $displayTitle = $metadata.Title
                if ($metadata.Year) { $displayTitle += " ($($metadata.Year))" }

                Write-Host "  $displayTitle" -ForegroundColor White
                if ($metadata.IMDBID) {
                    Write-Host "    IMDB: $($metadata.IMDBID)" -ForegroundColor Gray
                }
                if ($metadata.Genres -and $metadata.Genres.Count -gt 0) {
                    Write-Host "    Genres: $($metadata.Genres -join ', ')" -ForegroundColor Gray
                }
            }
        }
    }
    catch {
        Write-Host "Error scanning NFO files: $_" -ForegroundColor Red
        Write-Log "Error scanning NFO files: $_" "ERROR"
    }
}

#============================================
# DUPLICATE DETECTION & QUALITY SCORING
#============================================

<#
.SYNOPSIS
    Calculates a quality score for a video file based on its properties
.PARAMETER FileName
    The filename to analyze
.OUTPUTS
    Hashtable with Score, Resolution, Codec, Source, and details
.DESCRIPTION
    Analyzes filename for quality indicators and assigns a score:
    - Resolution: 2160p=100, 1080p=80, 720p=60, 480p=40
    - Source: BluRay=30, WEB-DL=25, WEBRip=20, HDTV=15, DVDRip=10
    - Codec: x265/HEVC=20, x264=15, XviD=5
    - Audio: Atmos=15, TrueHD=12, DTS-HD=10, DTS=8, AC3=5, AAC=3
    - HDR: HDR10+=15, HDR10=12, HDR=10, DolbyVision=15
#>
function Get-QualityScore {
    param(
        [string]$FileName,
        [string]$FilePath = $null
    )

    $quality = @{
        Score = 0
        Resolution = "Unknown"
        Codec = "Unknown"
        Source = "Unknown"
        Audio = "Unknown"
        HDR = $false
        HDRFormat = $null
        Bitrate = 0
        Width = 0
        Height = 0
        AudioChannels = 0
        Details = @()
        DataSource = "Filename"
    }

    # Try MediaInfo first if FilePath is provided
    $mediaInfo = $null
    if ($FilePath -and (Test-Path $FilePath -PathType Leaf)) {
        $mediaInfo = Get-MediaInfoDetails -FilePath $FilePath
    }

    if ($mediaInfo) {
        $quality.DataSource = "MediaInfo"
        $quality.Width = $mediaInfo.Width
        $quality.Height = $mediaInfo.Height
        $quality.Bitrate = $mediaInfo.Bitrate
        $quality.AudioChannels = $mediaInfo.AudioChannels

        # Resolution from actual dimensions
        if ($mediaInfo.Height -ge 2160) {
            $quality.Resolution = "2160p"
            $quality.Score += 100
            $quality.Details += "4K/2160p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+100)"
        }
        elseif ($mediaInfo.Height -ge 1080) {
            $quality.Resolution = "1080p"
            $quality.Score += 80
            $quality.Details += "1080p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+80)"
        }
        elseif ($mediaInfo.Height -ge 720) {
            $quality.Resolution = "720p"
            $quality.Score += 60
            $quality.Details += "720p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+60)"
        }
        elseif ($mediaInfo.Height -ge 480) {
            $quality.Resolution = "480p"
            $quality.Score += 40
            $quality.Details += "480p [MediaInfo: $($mediaInfo.Width)x$($mediaInfo.Height)] (+40)"
        }
        elseif ($mediaInfo.Height -gt 0) {
            $quality.Resolution = "$($mediaInfo.Height)p"
            $quality.Score += 20
            $quality.Details += "$($mediaInfo.Height)p [MediaInfo] (+20)"
        }

        # Video codec from MediaInfo
        $videoCodec = $mediaInfo.VideoCodec
        if ($videoCodec) {
            switch -Regex ($videoCodec) {
                'HEVC|H\.?265|V_MPEGH' {
                    $quality.Codec = "HEVC/x265"
                    $quality.Score += 20
                    $quality.Details += "HEVC/x265 [MediaInfo: $videoCodec] (+20)"
                }
                'AVC|H\.?264|V_MPEG4/ISO/AVC' {
                    $quality.Codec = "x264"
                    $quality.Score += 15
                    $quality.Details += "x264 [MediaInfo: $videoCodec] (+15)"
                }
                'AV1' {
                    $quality.Codec = "AV1"
                    $quality.Score += 25
                    $quality.Details += "AV1 [MediaInfo] (+25)"
                }
                'VP9' {
                    $quality.Codec = "VP9"
                    $quality.Score += 18
                    $quality.Details += "VP9 [MediaInfo] (+18)"
                }
                'MPEG-4|DivX|XviD' {
                    $quality.Codec = "XviD"
                    $quality.Score += 5
                    $quality.Details += "MPEG-4/XviD [MediaInfo: $videoCodec] (+5)"
                }
                'VC-1|WMV' {
                    $quality.Codec = "VC-1"
                    $quality.Score += 8
                    $quality.Details += "VC-1 [MediaInfo] (+8)"
                }
                default {
                    $quality.Codec = $videoCodec
                    $quality.Score += 10
                    $quality.Details += "$videoCodec [MediaInfo] (+10)"
                }
            }
        }

        # Audio codec from MediaInfo
        $audioCodec = $mediaInfo.AudioCodec
        $channels = $mediaInfo.AudioChannels
        if ($audioCodec) {
            $channelInfo = if ($channels -gt 0) { " ${channels}ch" } else { "" }
            switch -Regex ($audioCodec) {
                'Atmos|E-AC-3.*Atmos|TrueHD.*Atmos' {
                    $quality.Audio = "Atmos"
                    $quality.Score += 15
                    $quality.Details += "Atmos [MediaInfo$channelInfo] (+15)"
                }
                'TrueHD' {
                    $quality.Audio = "TrueHD"
                    $quality.Score += 12
                    $quality.Details += "TrueHD [MediaInfo$channelInfo] (+12)"
                }
                'DTS-HD|DTS.*HD' {
                    $quality.Audio = "DTS-HD"
                    $quality.Score += 10
                    $quality.Details += "DTS-HD [MediaInfo$channelInfo] (+10)"
                }
                'DTS.*X|DTS:X' {
                    $quality.Audio = "DTS:X"
                    $quality.Score += 14
                    $quality.Details += "DTS:X [MediaInfo$channelInfo] (+14)"
                }
                '^DTS$|^DTS\s' {
                    $quality.Audio = "DTS"
                    $quality.Score += 8
                    $quality.Details += "DTS [MediaInfo$channelInfo] (+8)"
                }
                'E-AC-3|EAC3|DD\+|Dolby Digital Plus' {
                    $quality.Audio = "EAC3"
                    $quality.Score += 7
                    $quality.Details += "EAC3/DD+ [MediaInfo$channelInfo] (+7)"
                }
                'AC-3|AC3|Dolby Digital' {
                    $quality.Audio = "AC3"
                    $quality.Score += 5
                    $quality.Details += "AC3 [MediaInfo$channelInfo] (+5)"
                }
                'AAC' {
                    $quality.Audio = "AAC"
                    $quality.Score += 3
                    $quality.Details += "AAC [MediaInfo$channelInfo] (+3)"
                }
                'FLAC' {
                    $quality.Audio = "FLAC"
                    $quality.Score += 6
                    $quality.Details += "FLAC [MediaInfo$channelInfo] (+6)"
                }
                'PCM|LPCM' {
                    $quality.Audio = "PCM"
                    $quality.Score += 4
                    $quality.Details += "PCM [MediaInfo$channelInfo] (+4)"
                }
                'Opus' {
                    $quality.Audio = "Opus"
                    $quality.Score += 4
                    $quality.Details += "Opus [MediaInfo$channelInfo] (+4)"
                }
                'Vorbis' {
                    $quality.Audio = "Vorbis"
                    $quality.Score += 2
                    $quality.Details += "Vorbis [MediaInfo$channelInfo] (+2)"
                }
                'MP3|MPEG Audio' {
                    $quality.Audio = "MP3"
                    $quality.Score += 1
                    $quality.Details += "MP3 [MediaInfo$channelInfo] (+1)"
                }
                default {
                    $quality.Audio = $audioCodec
                    $quality.Score += 2
                    $quality.Details += "$audioCodec [MediaInfo$channelInfo] (+2)"
                }
            }
        }

        # HDR detection from MediaInfo
        if ($mediaInfo.HDR) {
            $quality.HDR = $true
            # Get detailed HDR format if available
            $hdrFormat = Get-MediaInfoHDRFormat -FilePath $FilePath
            if ($hdrFormat) {
                $quality.HDRFormat = $hdrFormat.Format
                switch ($hdrFormat.Format) {
                    "Dolby Vision" {
                        $quality.Score += 18
                        $quality.Details += "Dolby Vision [MediaInfo] (+18)"
                    }
                    "HDR10+" {
                        $quality.Score += 16
                        $quality.Details += "HDR10+ [MediaInfo] (+16)"
                    }
                    "HDR10" {
                        $quality.Score += 12
                        $quality.Details += "HDR10 [MediaInfo] (+12)"
                    }
                    "HLG" {
                        $quality.Score += 10
                        $quality.Details += "HLG [MediaInfo] (+10)"
                    }
                    default {
                        $quality.Score += 10
                        $quality.Details += "HDR [MediaInfo] (+10)"
                    }
                }
            }
            else {
                $quality.Score += 10
                $quality.Details += "HDR [MediaInfo] (+10)"
            }
        }

        # Source detection still from filename (MediaInfo can't detect source)
        $fileNameLower = $FileName.ToLower()
        if ($fileNameLower -match 'bluray|blu-ray|bdrip|brrip') {
            $quality.Source = "BluRay"
            $quality.Score += 30
            $quality.Details += "BluRay (+30)"
        }
        elseif ($fileNameLower -match 'remux') {
            $quality.Source = "Remux"
            $quality.Score += 35
            $quality.Details += "Remux (+35)"
        }
        elseif ($fileNameLower -match 'web-dl|webdl') {
            $quality.Source = "WEB-DL"
            $quality.Score += 25
            $quality.Details += "WEB-DL (+25)"
        }
        elseif ($fileNameLower -match 'webrip') {
            $quality.Source = "WEBRip"
            $quality.Score += 20
            $quality.Details += "WEBRip (+20)"
        }
        elseif ($fileNameLower -match 'hdtv') {
            $quality.Source = "HDTV"
            $quality.Score += 15
            $quality.Details += "HDTV (+15)"
        }
        elseif ($fileNameLower -match 'dvdrip') {
            $quality.Source = "DVDRip"
            $quality.Score += 10
            $quality.Details += "DVDRip (+10)"
        }

        # Bitrate bonus (higher bitrate = better quality)
        if ($quality.Bitrate -gt 0) {
            $bitrateMbps = [math]::Round($quality.Bitrate / 1000000, 1)
            if ($bitrateMbps -ge 40) {
                $quality.Score += 20
                $quality.Details += "High Bitrate [${bitrateMbps} Mbps] (+20)"
            }
            elseif ($bitrateMbps -ge 20) {
                $quality.Score += 15
                $quality.Details += "Good Bitrate [${bitrateMbps} Mbps] (+15)"
            }
            elseif ($bitrateMbps -ge 10) {
                $quality.Score += 10
                $quality.Details += "Moderate Bitrate [${bitrateMbps} Mbps] (+10)"
            }
            elseif ($bitrateMbps -ge 5) {
                $quality.Score += 5
                $quality.Details += "Low Bitrate [${bitrateMbps} Mbps] (+5)"
            }
        }

        return $quality
    }

    # Fallback to filename parsing if MediaInfo not available
    $quality.DataSource = "Filename"
    $fileNameLower = $FileName.ToLower()

    # Resolution scoring
    if ($fileNameLower -match '2160p|4k|uhd') {
        $quality.Resolution = "2160p"
        $quality.Score += 100
        $quality.Details += "4K/2160p (+100)"
    }
    elseif ($fileNameLower -match '1080p') {
        $quality.Resolution = "1080p"
        $quality.Score += 80
        $quality.Details += "1080p (+80)"
    }
    elseif ($fileNameLower -match '720p') {
        $quality.Resolution = "720p"
        $quality.Score += 60
        $quality.Details += "720p (+60)"
    }
    elseif ($fileNameLower -match '480p|dvd') {
        $quality.Resolution = "480p"
        $quality.Score += 40
        $quality.Details += "480p (+40)"
    }

    # Source scoring
    if ($fileNameLower -match 'remux') {
        $quality.Source = "Remux"
        $quality.Score += 35
        $quality.Details += "Remux (+35)"
    }
    elseif ($fileNameLower -match 'bluray|blu-ray|bdrip|brrip') {
        $quality.Source = "BluRay"
        $quality.Score += 30
        $quality.Details += "BluRay (+30)"
    }
    elseif ($fileNameLower -match 'web-dl|webdl') {
        $quality.Source = "WEB-DL"
        $quality.Score += 25
        $quality.Details += "WEB-DL (+25)"
    }
    elseif ($fileNameLower -match 'webrip') {
        $quality.Source = "WEBRip"
        $quality.Score += 20
        $quality.Details += "WEBRip (+20)"
    }
    elseif ($fileNameLower -match 'hdtv') {
        $quality.Source = "HDTV"
        $quality.Score += 15
        $quality.Details += "HDTV (+15)"
    }
    elseif ($fileNameLower -match 'dvdrip') {
        $quality.Source = "DVDRip"
        $quality.Score += 10
        $quality.Details += "DVDRip (+10)"
    }

    # Codec scoring
    if ($fileNameLower -match 'av1') {
        $quality.Codec = "AV1"
        $quality.Score += 25
        $quality.Details += "AV1 (+25)"
    }
    elseif ($fileNameLower -match 'x265|h\.?265|hevc') {
        $quality.Codec = "HEVC/x265"
        $quality.Score += 20
        $quality.Details += "HEVC/x265 (+20)"
    }
    elseif ($fileNameLower -match 'vp9') {
        $quality.Codec = "VP9"
        $quality.Score += 18
        $quality.Details += "VP9 (+18)"
    }
    elseif ($fileNameLower -match 'x264|h\.?264|avc') {
        $quality.Codec = "x264"
        $quality.Score += 15
        $quality.Details += "x264 (+15)"
    }
    elseif ($fileNameLower -match 'xvid|divx') {
        $quality.Codec = "XviD"
        $quality.Score += 5
        $quality.Details += "XviD (+5)"
    }

    # Audio scoring
    if ($fileNameLower -match 'atmos') {
        $quality.Audio = "Atmos"
        $quality.Score += 15
        $quality.Details += "Atmos (+15)"
    }
    elseif ($fileNameLower -match 'dts[\s\.\-]?x|dtsx') {
        $quality.Audio = "DTS:X"
        $quality.Score += 14
        $quality.Details += "DTS:X (+14)"
    }
    elseif ($fileNameLower -match 'truehd') {
        $quality.Audio = "TrueHD"
        $quality.Score += 12
        $quality.Details += "TrueHD (+12)"
    }
    elseif ($fileNameLower -match 'dts-hd|dtshd|dts[\s\.\-]?hd[\s\.\-]?ma') {
        $quality.Audio = "DTS-HD"
        $quality.Score += 10
        $quality.Details += "DTS-HD (+10)"
    }
    elseif ($fileNameLower -match 'dts') {
        $quality.Audio = "DTS"
        $quality.Score += 8
        $quality.Details += "DTS (+8)"
    }
    elseif ($fileNameLower -match 'eac3|ddp|dd\+|dolby\s*digital\s*plus') {
        $quality.Audio = "EAC3"
        $quality.Score += 7
        $quality.Details += "EAC3/DD+ (+7)"
    }
    elseif ($fileNameLower -match 'ac3|dd5\.?1') {
        $quality.Audio = "AC3"
        $quality.Score += 5
        $quality.Details += "AC3 (+5)"
    }
    elseif ($fileNameLower -match 'flac') {
        $quality.Audio = "FLAC"
        $quality.Score += 6
        $quality.Details += "FLAC (+6)"
    }
    elseif ($fileNameLower -match 'aac') {
        $quality.Audio = "AAC"
        $quality.Score += 3
        $quality.Details += "AAC (+3)"
    }
    elseif ($fileNameLower -match 'opus') {
        $quality.Audio = "Opus"
        $quality.Score += 4
        $quality.Details += "Opus (+4)"
    }

    # HDR scoring
    if ($fileNameLower -match 'dolby[\s\.\-]?vision|dovi|dv[\s\.\-]hdr|\.dv\.') {
        $quality.HDR = $true
        $quality.HDRFormat = "Dolby Vision"
        $quality.Score += 18
        $quality.Details += "Dolby Vision (+18)"
    }
    elseif ($fileNameLower -match 'hdr10\+|hdr10plus') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10+"
        $quality.Score += 16
        $quality.Details += "HDR10+ (+16)"
    }
    elseif ($fileNameLower -match 'hdr10') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR10"
        $quality.Score += 12
        $quality.Details += "HDR10 (+12)"
    }
    elseif ($fileNameLower -match 'hlg') {
        $quality.HDR = $true
        $quality.HDRFormat = "HLG"
        $quality.Score += 10
        $quality.Details += "HLG (+10)"
    }
    elseif ($fileNameLower -match 'hdr') {
        $quality.HDR = $true
        $quality.HDRFormat = "HDR"
        $quality.Score += 10
        $quality.Details += "HDR (+10)"
    }

    return $quality
}

<#
.SYNOPSIS
    Extracts a normalized title from a movie folder or filename for duplicate matching
.PARAMETER Name
    The folder or file name to normalize
.OUTPUTS
    Hashtable with NormalizedTitle and Year
#>
function Get-NormalizedTitle {
    param(
        [string]$Name
    )

    $result = @{
        NormalizedTitle = $null
        Year = $null
    }

    # Remove extension if present
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Name)

    # Try to extract year
    if ($name -match '[\(\[\s]*(19|20)\d{2}[\)\]\s]*') {
        $yearMatch = [regex]::Match($name, '(19|20)\d{2}')
        if ($yearMatch.Success) {
            $result.Year = $yearMatch.Value
        }
    }

    # Remove everything after year or quality tags
    $title = $name -replace '[\(\[]?(19|20)\d{2}[\)\]]?.*$', ''
    $title = $title -replace '\s*(720p|1080p|2160p|4K|HDRip|DVDRip|BRRip|BluRay|WEB-DL|WEBRip|x264|x265|HEVC).*$', ''

    # Normalize the title
    $title = $title -replace '\.', ' '
    $title = $title -replace '[_-]', ' '
    $title = $title -replace '\s+', ' '
    $title = $title.Trim().ToLower()

    # Remove common articles for better matching
    $title = $title -replace '^(the|a|an)\s+', ''

    $result.NormalizedTitle = $title

    return $result
}

<#
.SYNOPSIS
    Finds potential duplicate movies in the library
.PARAMETER Path
    The root path of the movie library
.OUTPUTS
    Array of duplicate groups, each containing matching movies
.DESCRIPTION
    Scans movie folders, normalizes titles, and groups potential duplicates.
    Each group contains the full paths and quality scores of matching items.
#>
function Find-DuplicateMovies {
    param(
        [string]$Path
    )

    Write-Host "`nScanning for duplicate movies..." -ForegroundColor Yellow
    Write-Log "Starting duplicate scan in: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return @()
        }

        # Build a lookup of normalized titles
        $titleLookup = @{}
        $totalFolders = $movieFolders.Count
        $currentIndex = 0

        foreach ($folder in $movieFolders) {
            $currentIndex++
            $percentComplete = [math]::Round(($currentIndex / $totalFolders) * 100)
            Write-Progress -Activity "Scanning for duplicates" -Status "Processing $currentIndex of $totalFolders - $($folder.Name)" -PercentComplete $percentComplete

            $titleInfo = Get-NormalizedTitle -Name $folder.Name

            # Find the main video file for size info
            $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Sort-Object Length -Descending |
                Select-Object -First 1

            $quality = Get-QualityScore -FileName $folder.Name -FilePath $(if ($videoFile) { $videoFile.FullName } else { $null })

            $entry = @{
                Path = $folder.FullName
                OriginalName = $folder.Name
                Year = $titleInfo.Year
                Quality = $quality
                FileSize = if ($videoFile) { $videoFile.Length } else { 0 }
            }

            $key = $titleInfo.NormalizedTitle
            if ($titleInfo.Year) {
                $key = "$($titleInfo.NormalizedTitle)|$($titleInfo.Year)"
            }

            if (-not $titleLookup.ContainsKey($key)) {
                $titleLookup[$key] = @()
            }
            $titleLookup[$key] += $entry
        }

        Write-Progress -Activity "Scanning for duplicates" -Completed

        # Find duplicates (groups with more than 1 entry)
        $duplicates = @()
        foreach ($key in $titleLookup.Keys) {
            if ($titleLookup[$key].Count -gt 1) {
                $duplicates += ,@($titleLookup[$key])
            }
        }

        return $duplicates
    }
    catch {
        Write-Host "Error scanning for duplicates: $_" -ForegroundColor Red
        Write-Log "Error scanning for duplicates: $_" "ERROR"
        return @()
    }
}

<#
.SYNOPSIS
    Displays duplicate movies and suggests which to keep based on quality
.PARAMETER Path
    The root path of the movie library
#>
function Show-DuplicateReport {
    param(
        [string]$Path
    )

    $duplicates = Find-DuplicateMovies -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "No duplicate movies found!" -ForegroundColor Green
        Write-Log "No duplicates found" "INFO"
        return
    }

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║                    DUPLICATE REPORT                          ║" -ForegroundColor Yellow
    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "║  Found $($duplicates.Count) potential duplicate group(s)".PadRight(63) + "║" -ForegroundColor Yellow
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow

    $groupNum = 1
    foreach ($group in $duplicates) {
        Write-Host "`n[$groupNum] Potential Duplicates:" -ForegroundColor Cyan

        # Sort by quality score descending
        $sorted = $group | Sort-Object { $_.Quality.Score } -Descending

        $first = $true
        foreach ($movie in $sorted) {
            $sizeStr = Format-FileSize $movie.FileSize
            $scoreStr = "Score: $($movie.Quality.Score)"

            if ($first) {
                Write-Host "  [KEEP] " -ForegroundColor Green -NoNewline
                $first = $false
            } else {
                Write-Host "  [DEL?] " -ForegroundColor Red -NoNewline
            }

            Write-Host "$($movie.OriginalName)" -ForegroundColor White
            Write-Host "         $scoreStr | $sizeStr | $($movie.Quality.Resolution) $($movie.Quality.Source)" -ForegroundColor Gray
        }

        $groupNum++
    }

    Write-Host "`n" -NoNewline
    Write-Log "Found $($duplicates.Count) duplicate groups" "INFO"
}

<#
.SYNOPSIS
    Interactively removes duplicate movies, keeping highest quality versions
.PARAMETER Path
    The root path of the movie library
#>
function Remove-DuplicateMovies {
    param(
        [string]$Path
    )

    $duplicates = Find-DuplicateMovies -Path $Path

    if ($duplicates.Count -eq 0) {
        Write-Host "No duplicate movies found!" -ForegroundColor Green
        return
    }

    Write-Host "`nFound $($duplicates.Count) duplicate group(s)" -ForegroundColor Yellow

    $confirmAll = Read-Host "Delete lower quality duplicates? (Y/N/Review each) [N]"

    if ($confirmAll -eq 'Y' -or $confirmAll -eq 'y') {
        foreach ($group in $duplicates) {
            $sorted = $group | Sort-Object { $_.Quality.Score } -Descending
            # Keep first (highest quality), delete the rest
            $toDelete = $sorted | Select-Object -Skip 1

            foreach ($movie in $toDelete) {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would delete: $($movie.OriginalName)" -ForegroundColor Yellow
                    Write-Log "Would delete duplicate: $($movie.Path)" "DRY-RUN"
                } else {
                    try {
                        $size = (Get-ChildItem -Path $movie.Path -Recurse | Measure-Object -Property Length -Sum).Sum
                        Remove-Item -Path $movie.Path -Recurse -Force -ErrorAction Stop
                        Write-Host "Deleted: $($movie.OriginalName)" -ForegroundColor Red
                        Write-Log "Deleted duplicate: $($movie.Path)" "INFO"
                        $script:Stats.FilesDeleted++
                        $script:Stats.BytesDeleted += $size
                    }
                    catch {
                        Write-Host "Error deleting $($movie.OriginalName): $_" -ForegroundColor Red
                        Write-Log "Error deleting duplicate $($movie.Path): $_" "ERROR"
                    }
                }
            }
        }
    }
    elseif ($confirmAll -eq 'R' -or $confirmAll -eq 'r') {
        # Review each duplicate group
        foreach ($group in $duplicates) {
            $sorted = $group | Sort-Object { $_.Quality.Score } -Descending

            Write-Host "`nDuplicate Group:" -ForegroundColor Cyan
            $i = 1
            foreach ($movie in $sorted) {
                Write-Host "  [$i] $($movie.OriginalName) (Score: $($movie.Quality.Score))" -ForegroundColor White
                $i++
            }

            $choice = Read-Host "Enter number to KEEP (others deleted), or S to skip"
            if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $sorted.Count) {
                $keepIndex = [int]$choice - 1
                for ($j = 0; $j -lt $sorted.Count; $j++) {
                    if ($j -ne $keepIndex) {
                        $movie = $sorted[$j]
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would delete: $($movie.OriginalName)" -ForegroundColor Yellow
                        } else {
                            try {
                                $size = (Get-ChildItem -Path $movie.Path -Recurse | Measure-Object -Property Length -Sum).Sum
                                Remove-Item -Path $movie.Path -Recurse -Force -ErrorAction Stop
                                Write-Host "Deleted: $($movie.OriginalName)" -ForegroundColor Red
                                $script:Stats.FilesDeleted++
                                $script:Stats.BytesDeleted += $size
                            }
                            catch {
                                Write-Host "Error: $_" -ForegroundColor Red
                            }
                        }
                    }
                }
            } else {
                Write-Host "Skipped" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "No duplicates removed" -ForegroundColor Cyan
    }
}

#============================================
# HEALTH CHECK & CODEC ANALYSIS
#============================================

<#
.SYNOPSIS
    Performs a health check on the media library
.PARAMETER Path
    The root path of the media library
.DESCRIPTION
    Validates the media library by checking for:
    - Empty folders
    - Missing video files in movie folders
    - Corrupted/zero-byte files
    - Orphaned subtitle files (no matching video)
    - Missing NFO files
    - Naming issues
    - Very small video files (likely samples)
#>
function Invoke-LibraryHealthCheck {
    param(
        [string]$Path,
        [string]$MediaType = "Movies"  # Movies or TVShows
    )

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                  LIBRARY HEALTH CHECK                        ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Log "Starting health check for: $Path" "INFO"

    $issues = @{
        EmptyFolders = @()
        NoVideoFiles = @()
        ZeroByteFiles = @()
        OrphanedSubtitles = @()
        MissingNFO = @()
        SmallVideos = @()
        NamingIssues = @()
    }

    try {
        # Check for empty folders
        Write-Host "`nChecking for empty folders..." -ForegroundColor Yellow
        $emptyFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { (Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0 }
        $issues.EmptyFolders = $emptyFolders

        # Check movie folders for missing videos
        Write-Host "Checking for folders without video files..." -ForegroundColor Yellow
        $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        foreach ($folder in $folders) {
            $hasVideo = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Select-Object -First 1

            if (-not $hasVideo) {
                $issues.NoVideoFiles += $folder
            }
        }

        # Check for zero-byte files
        Write-Host "Checking for corrupted/zero-byte files..." -ForegroundColor Yellow
        $zeroByteFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -eq 0 }
        $issues.ZeroByteFiles = $zeroByteFiles

        # Check for very small video files (likely samples, under 50MB)
        Write-Host "Checking for suspiciously small video files..." -ForegroundColor Yellow
        $smallVideos = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object {
                $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                $_.Length -lt 50MB -and
                $_.Name -notmatch 'sample|trailer|teaser'
            }
        $issues.SmallVideos = $smallVideos

        # Check for orphaned subtitle files
        Write-Host "Checking for orphaned subtitle files..." -ForegroundColor Yellow
        $subtitleFiles = Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.SubtitleExtensions -contains $_.Extension.ToLower() }

        foreach ($sub in $subtitleFiles) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sub.Name)
            # Remove language suffix if present
            $baseName = $baseName -replace '\.(eng|en|english|spa|es|fre|fr|ger|de)$', ''

            $hasMatchingVideo = Get-ChildItem -Path $sub.DirectoryName -File -ErrorAction SilentlyContinue |
                Where-Object {
                    $script:Config.VideoExtensions -contains $_.Extension.ToLower() -and
                    [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $baseName
                } | Select-Object -First 1

            if (-not $hasMatchingVideo) {
                $issues.OrphanedSubtitles += $sub
            }
        }

        # Check for missing NFO files (movies only)
        if ($MediaType -eq "Movies") {
            Write-Host "Checking for missing NFO files..." -ForegroundColor Yellow
            foreach ($folder in $folders) {
                $hasNFO = Get-ChildItem -Path $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                $hasVideo = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                    Select-Object -First 1

                if ($hasVideo -and -not $hasNFO) {
                    $issues.MissingNFO += $folder
                }
            }
        }

        # Check for naming issues
        Write-Host "Checking for naming issues..." -ForegroundColor Yellow
        foreach ($folder in $folders) {
            # Check for dots in folder names (should be spaces)
            if ($folder.Name -match '\..*\.' -and $folder.Name -notmatch '\(\d{4}\)') {
                $issues.NamingIssues += @{
                    Path = $folder.FullName
                    Issue = "Contains dots (should be spaces)"
                }
            }
            # Check for missing year
            if ($MediaType -eq "Movies" -and $folder.Name -notmatch '(19|20)\d{2}') {
                $issues.NamingIssues += @{
                    Path = $folder.FullName
                    Issue = "Missing year"
                }
            }
        }

        # Display results
        Write-Host "`n=== Health Check Results ===" -ForegroundColor Cyan

        $totalIssues = 0

        if ($issues.EmptyFolders.Count -gt 0) {
            Write-Host "`nEmpty Folders ($($issues.EmptyFolders.Count)):" -ForegroundColor Yellow
            $issues.EmptyFolders | ForEach-Object {
                Write-Host "  - $($_.FullName)" -ForegroundColor Gray
            }
            $totalIssues += $issues.EmptyFolders.Count
        }

        if ($issues.NoVideoFiles.Count -gt 0) {
            Write-Host "`nFolders Without Video Files ($($issues.NoVideoFiles.Count)):" -ForegroundColor Yellow
            $issues.NoVideoFiles | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            $totalIssues += $issues.NoVideoFiles.Count
        }

        if ($issues.ZeroByteFiles.Count -gt 0) {
            Write-Host "`nZero-Byte Files ($($issues.ZeroByteFiles.Count)):" -ForegroundColor Red
            $issues.ZeroByteFiles | ForEach-Object {
                Write-Host "  - $($_.FullName)" -ForegroundColor Gray
            }
            $totalIssues += $issues.ZeroByteFiles.Count
        }

        if ($issues.SmallVideos.Count -gt 0) {
            Write-Host "`nSuspiciously Small Videos ($($issues.SmallVideos.Count)):" -ForegroundColor Yellow
            $issues.SmallVideos | ForEach-Object {
                Write-Host "  - $($_.Name) ($(Format-FileSize $_.Length))" -ForegroundColor Gray
            }
            $totalIssues += $issues.SmallVideos.Count
        }

        if ($issues.OrphanedSubtitles.Count -gt 0) {
            Write-Host "`nOrphaned Subtitle Files ($($issues.OrphanedSubtitles.Count)):" -ForegroundColor Yellow
            $issues.OrphanedSubtitles | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            $totalIssues += $issues.OrphanedSubtitles.Count
        }

        if ($issues.MissingNFO.Count -gt 0) {
            Write-Host "`nMovies Missing NFO Files ($($issues.MissingNFO.Count)):" -ForegroundColor Yellow
            $issues.MissingNFO | Select-Object -First 10 | ForEach-Object {
                Write-Host "  - $($_.Name)" -ForegroundColor Gray
            }
            if ($issues.MissingNFO.Count -gt 10) {
                Write-Host "  ... and $($issues.MissingNFO.Count - 10) more" -ForegroundColor Gray
            }
            $totalIssues += $issues.MissingNFO.Count
        }

        if ($issues.NamingIssues.Count -gt 0) {
            Write-Host "`nNaming Issues ($($issues.NamingIssues.Count)):" -ForegroundColor Yellow
            $issues.NamingIssues | Select-Object -First 10 | ForEach-Object {
                Write-Host "  - $($_.Path): $($_.Issue)" -ForegroundColor Gray
            }
            if ($issues.NamingIssues.Count -gt 10) {
                Write-Host "  ... and $($issues.NamingIssues.Count - 10) more" -ForegroundColor Gray
            }
            $totalIssues += $issues.NamingIssues.Count
        }

        # Summary
        Write-Host "`n" -NoNewline
        if ($totalIssues -eq 0) {
            Write-Host "Library is healthy! No issues found." -ForegroundColor Green
        } else {
            Write-Host "Found $totalIssues issue(s) in the library." -ForegroundColor Yellow
        }

        Write-Log "Health check completed: $totalIssues issues found" "INFO"
        return $issues
    }
    catch {
        Write-Host "Error during health check: $_" -ForegroundColor Red
        Write-Log "Error during health check: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Analyzes video files using MediaInfo or file properties
.PARAMETER Path
    The path to the video file or folder
.DESCRIPTION
    Extracts codec information from video files and generates reports.
    Uses file naming patterns if MediaInfo is not available.
#>
function Get-VideoCodecInfo {
    param(
        [string]$FilePath
    )

    $info = @{
        FileName = [System.IO.Path]::GetFileName($FilePath)
        FileSize = 0
        VideoCodec = "Unknown"
        AudioCodec = "Unknown"
        Resolution = "Unknown"
        HDR = $false
        Container = "Unknown"
        NeedsTranscode = $false
        TranscodeReason = @()
    }

    try {
        $file = Get-Item $FilePath -ErrorAction Stop
        $info.FileSize = $file.Length
        $info.Container = $file.Extension.TrimStart('.').ToUpper()

        # Get quality info using MediaInfo when available
        $quality = Get-QualityScore -FileName $file.Name -FilePath $file.FullName
        $info.Resolution = $quality.Resolution
        $info.VideoCodec = $quality.Codec
        $info.AudioCodec = $quality.Audio
        $info.HDR = $quality.HDR

        # Determine processing needed based on codec and container
        # TranscodeMode: "none" = keep as-is, "remux" = copy streams to MKV, "transcode" = re-encode
        $info.TranscodeMode = "none"

        # XviD/DivX are legacy codecs - need full transcode to H.264
        if ($info.VideoCodec -eq "XviD" -or $info.VideoCodec -eq "DivX" -or $info.VideoCodec -eq "MPEG-4") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "XviD/DivX/MPEG-4 is a legacy codec - will transcode to H.264"
        }
        # H.264 in AVI container - remux only (no quality loss)
        elseif (($info.VideoCodec -eq "H264" -or $info.VideoCodec -eq "AVC" -or $info.VideoCodec -eq "H.264") -and $info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "remux"
            $info.TranscodeReason += "H.264 in AVI container - will remux to MKV (no re-encoding)"
        }
        # AVI with unknown codec - likely legacy, transcode it
        elseif ($info.Container -eq "AVI") {
            $info.NeedsTranscode = $true
            $info.TranscodeMode = "transcode"
            $info.TranscodeReason += "AVI with legacy codec - will transcode to H.264"
        }
        # HEVC/x265 - keep as-is (it's a modern efficient codec)
        # H.264 in modern containers (MKV, MP4) - keep as-is
        # Note: HDR and 4K are NOT flagged for transcode - they're features, not problems

        return $info
    }
    catch {
        Write-Log "Error analyzing file $FilePath : $_" "ERROR"
        return $info
    }
}

<#
.SYNOPSIS
    Generates a codec analysis report and transcoding queue
.PARAMETER Path
    The root path of the media library
.PARAMETER ExportPath
    Optional path to export the transcoding queue CSV
#>
function Invoke-CodecAnalysis {
    param(
        [string]$Path,
        [string]$ExportPath = $null
    )

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    CODEC ANALYSIS                            ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Write-Log "Starting codec analysis for: $Path" "INFO"

    try {
        # Find all video files
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        Write-Host "`nAnalyzing $($videoFiles.Count) video file(s)..." -ForegroundColor Yellow

        $analysis = @{
            TotalFiles = $videoFiles.Count
            TotalSize = 0
            ByResolution = @{}
            ByCodec = @{}
            ByContainer = @{}
            NeedTranscode = @()
        }

        $current = 1
        foreach ($file in $videoFiles) {
            $percentComplete = ($current / $videoFiles.Count) * 100
            Write-Progress -Activity "Analyzing codecs" -Status "[$current/$($videoFiles.Count)] $($file.Name)" -PercentComplete $percentComplete

            $info = Get-VideoCodecInfo -FilePath $file.FullName
            $analysis.TotalSize += $info.FileSize

            # Count by resolution
            if (-not $analysis.ByResolution.ContainsKey($info.Resolution)) {
                $analysis.ByResolution[$info.Resolution] = 0
            }
            $analysis.ByResolution[$info.Resolution]++

            # Count by codec
            if (-not $analysis.ByCodec.ContainsKey($info.VideoCodec)) {
                $analysis.ByCodec[$info.VideoCodec] = 0
            }
            $analysis.ByCodec[$info.VideoCodec]++

            # Count by container
            if (-not $analysis.ByContainer.ContainsKey($info.Container)) {
                $analysis.ByContainer[$info.Container] = 0
            }
            $analysis.ByContainer[$info.Container]++

            # Add to transcode queue if needed
            if ($info.NeedsTranscode) {
                $analysis.NeedTranscode += @{
                    Path = $file.FullName
                    FileName = $info.FileName
                    Size = $info.FileSize
                    Resolution = $info.Resolution
                    Codec = $info.VideoCodec
                    Container = $info.Container
                    Reasons = $info.TranscodeReason -join "; "
                    TranscodeMode = $info.TranscodeMode
                }
            }

            $current++
        }
        Write-Progress -Activity "Analyzing codecs" -Completed

        # Display results
        Write-Host "`n=== Library Statistics ===" -ForegroundColor Cyan
        Write-Host "Total Files: $($analysis.TotalFiles)" -ForegroundColor White
        Write-Host "Total Size: $(Format-FileSize $analysis.TotalSize)" -ForegroundColor White

        Write-Host "`n=== By Resolution ===" -ForegroundColor Cyan
        $analysis.ByResolution.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Video Codec ===" -ForegroundColor Cyan
        $analysis.ByCodec.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        Write-Host "`n=== By Container ===" -ForegroundColor Cyan
        $analysis.ByContainer.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
            $pct = [math]::Round(($_.Value / $analysis.TotalFiles) * 100, 1)
            Write-Host "  $($_.Key): $($_.Value) ($pct%)" -ForegroundColor White
        }

        # Transcode queue
        if ($analysis.NeedTranscode.Count -gt 0) {
            Write-Host "`n=== Potential Transcoding Queue ===" -ForegroundColor Yellow
            Write-Host "Found $($analysis.NeedTranscode.Count) file(s) that may benefit from transcoding:" -ForegroundColor White

            $analysis.NeedTranscode | Select-Object -First 10 | ForEach-Object {
                Write-Host "`n  $($_.FileName)" -ForegroundColor White
                Write-Host "    Size: $(Format-FileSize $_.Size) | $($_.Resolution) | $($_.Codec)" -ForegroundColor Gray
                Write-Host "    Reason: $($_.Reasons)" -ForegroundColor Yellow
            }

            if ($analysis.NeedTranscode.Count -gt 10) {
                Write-Host "`n  ... and $($analysis.NeedTranscode.Count - 10) more files" -ForegroundColor Gray
            }

            # Export option
            if (-not $ExportPath) {
                $exportInput = Read-Host "`nExport transcoding queue to CSV? (Y/N) [N]"
                if ($exportInput -eq 'Y' -or $exportInput -eq 'y') {
                    $ExportPath = Join-Path $PSScriptRoot "TranscodeQueue_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                }
            }

            if ($ExportPath) {
                $analysis.NeedTranscode | Select-Object FileName, Path, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, Resolution, Codec, Container, Reasons |
                    Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nTranscoding queue exported to: $ExportPath" -ForegroundColor Green
                Write-Log "Transcoding queue exported to: $ExportPath" "INFO"
            }
        } else {
            Write-Host "`nNo files require transcoding for compatibility." -ForegroundColor Green
        }

        Write-Log "Codec analysis completed: $($analysis.TotalFiles) files analyzed" "INFO"
        return $analysis
    }
    catch {
        Write-Host "Error during codec analysis: $_" -ForegroundColor Red
        Write-Log "Error during codec analysis: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Generates FFmpeg commands for transcoding files in the queue
.PARAMETER TranscodeQueue
    Array of files needing transcoding (from Invoke-CodecAnalysis)
.PARAMETER TargetCodec
    Target video codec (default: libx264)
.DESCRIPTION
    Creates a PowerShell script that processes files in-place:
    - Remux: Copies streams to MKV container (fast, no quality loss)
    - Transcode: Re-encodes legacy codecs to H.264
    Output files replace originals in their original folders.
    Original files are deleted only after successful processing.
#>
function New-TranscodeScript {
    param(
        [array]$TranscodeQueue,
        [string]$TargetCodec = "libx264",
        [string]$TargetResolution = $null  # e.g., "1920:1080" for 1080p
    )

    if ($TranscodeQueue.Count -eq 0) {
        Write-Host "No files in transcode queue" -ForegroundColor Cyan
        return
    }

    # Check FFmpeg installation
    if (-not (Test-FFmpegInstallation)) {
        Write-Host "`nFFmpeg not found!" -ForegroundColor Red
        Write-Host "Please install FFmpeg to use transcoding features." -ForegroundColor Yellow
        Write-Host "Download from: https://ffmpeg.org/download.html" -ForegroundColor Cyan
        Write-Host "Or install via: winget install ffmpeg" -ForegroundColor Cyan
        Write-Log "FFmpeg not found - transcode script generation aborted" "ERROR"
        return $null
    }

    # Count files by mode
    $remuxCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "remux" }).Count
    $transcodeCount = ($TranscodeQueue | Where-Object { $_.TranscodeMode -eq "transcode" -or $_.TranscodeMode -eq $null }).Count

    $ffmpegPath = $script:Config.FFmpegPath
    $scriptPath = Join-Path $PSScriptRoot "transcode_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"

    $scriptContent = @"
# MediaCleaner Transcode Script
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Target Codec: $TargetCodec
# Files to process: $($TranscodeQueue.Count)
#   - Remux only (no re-encoding): $remuxCount
#   - Full transcode: $transcodeCount
#
# Output files are saved in the same folder as the original.
# Original files are deleted after successful processing.

`$ffmpegPath = "ffmpeg"  # Update this path if ffmpeg is not in PATH

`$files = @(
"@

    foreach ($item in $TranscodeQueue) {
        # Output goes to same folder as input, with .mkv extension
        $inputDir = [System.IO.Path]::GetDirectoryName($item.Path)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $outputFile = Join-Path $inputDir ($baseName + ".mkv")
        $mode = if ($item.TranscodeMode) { $item.TranscodeMode } else { "transcode" }
        # Escape single quotes by doubling them for PowerShell string literals
        $escapedInput = $item.Path -replace "'", "''"
        $escapedOutput = $outputFile -replace "'", "''"
        $scriptContent += "`n    @{ Input = '$escapedInput'; Output = '$escapedOutput'; Mode = '$mode' },"
    }

    $resolutionParam = if ($TargetResolution) { "-vf scale=$TargetResolution" } else { "" }

    $scriptContent += @"

)

`$total = `$files.Count
`$current = 1
`$remuxed = 0
`$transcoded = 0
`$failed = 0
`$deleted = 0

foreach (`$file in `$files) {
    # Skip if input and output are the same file (already .mkv with same name)
    if (`$file.Input -eq `$file.Output) {
        Write-Host "[`$current/`$total] Skipping (output same as input): `$(`$file.Input)" -ForegroundColor Gray
        `$current++
        continue
    }

    # Use temp file to avoid issues if input/output are in same folder
    `$tempOutput = `$file.Output + ".tmp.mkv"

    if (`$file.Mode -eq "remux") {
        # REMUX: Copy streams without re-encoding (fast, no quality loss)
        Write-Host "[`$current/`$total] Remuxing (no re-encode): `$(`$file.Input)" -ForegroundColor Cyan
        & `$ffmpegPath -i "`$(`$file.Input)" -c:v copy -c:a copy -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Delete original and rename temp to final
                Remove-Item -Path `$file.Input -Force
                Rename-Item -Path `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Remuxed & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$remuxed++
                `$deleted++
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    } else {
        # TRANSCODE: Re-encode video to H.264
        Write-Host "[`$current/`$total] Transcoding: `$(`$file.Input)" -ForegroundColor Yellow
        & `$ffmpegPath -i "`$(`$file.Input)" -map 0:v -map 0:a -map 0:s? -c:v $TargetCodec -crf 23 -preset medium $resolutionParam -c:a aac -b:a 192k -c:s copy "`$tempOutput" -y

        if (`$LASTEXITCODE -eq 0 -and (Test-Path `$tempOutput)) {
            # Verify output file is valid (has size > 0)
            `$outSize = (Get-Item `$tempOutput).Length
            if (`$outSize -gt 0) {
                # Delete original and rename temp to final
                Remove-Item -Path `$file.Input -Force
                Rename-Item -Path `$tempOutput -NewName ([System.IO.Path]::GetFileName(`$file.Output))
                Write-Host "  -> Transcoded & replaced: `$(`$file.Output)" -ForegroundColor Green
                `$transcoded++
                `$deleted++

                # 5-second pause to allow user to stop
                Write-Host "  Press any key within 5 seconds to stop, or wait to continue..." -ForegroundColor Magenta
                `$timeout = 5
                `$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                while (`$stopwatch.Elapsed.TotalSeconds -lt `$timeout) {
                    if ([Console]::KeyAvailable) {
                        `$null = [Console]::ReadKey(`$true)
                        Write-Host "``n  User requested stop. Exiting..." -ForegroundColor Yellow
                        Write-Host "``n========== Summary (Stopped Early) ==========" -ForegroundColor Cyan
                        Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
                        Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
                        Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
                        Write-Host "Failed: `$failed" -ForegroundColor Red
                        Write-Host "==============================================" -ForegroundColor Cyan
                        exit
                    }
                    Start-Sleep -Milliseconds 100
                }
            } else {
                Write-Host "  -> Failed (output empty): `$(`$file.Input)" -ForegroundColor Red
                Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
                `$failed++
            }
        } else {
            Write-Host "  -> Failed: `$(`$file.Input)" -ForegroundColor Red
            Remove-Item -Path `$tempOutput -Force -ErrorAction SilentlyContinue
            `$failed++
        }
    }

    `$current++
}

Write-Host "`n========== Summary ==========" -ForegroundColor Cyan
Write-Host "Remuxed (fast, no quality loss): `$remuxed" -ForegroundColor Cyan
Write-Host "Transcoded (re-encoded): `$transcoded" -ForegroundColor Yellow
Write-Host "Original files deleted: `$deleted" -ForegroundColor Green
Write-Host "Failed: `$failed" -ForegroundColor $(if (`$failed -gt 0) { 'Red' } else { 'Green' })
Write-Host "==============================" -ForegroundColor Cyan
"@

    $scriptContent | Out-File -FilePath $scriptPath -Encoding UTF8
    Write-Host "`nTranscode script created: $scriptPath" -ForegroundColor Green
    Write-Host "  - Remux only (fast, no quality loss): $remuxCount files" -ForegroundColor Cyan
    Write-Host "  - Full transcode (re-encode): $transcodeCount files" -ForegroundColor Yellow
    Write-Host "`nRun the script to start processing, or edit it to customize parameters." -ForegroundColor Cyan
    Write-Log "Transcode script created: $scriptPath - Remux: $remuxCount, Transcode: $transcodeCount" "INFO"

    return $scriptPath
}

#============================================
# TMDB API INTEGRATION
#============================================

<#
.SYNOPSIS
    Searches TMDB for a movie by title and year
.PARAMETER Title
    The movie title to search for
.PARAMETER Year
    The movie year (optional but recommended)
.PARAMETER ApiKey
    TMDB API key (get one free at https://www.themoviedb.org/settings/api)
.OUTPUTS
    Hashtable with movie metadata from TMDB
#>
function Search-TMDBMovie {
    param(
        [string]$Title,
        [string]$Year = $null,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Log "TMDB API key not provided" "WARNING"
        return $null
    }

    try {
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/movie?api_key=$ApiKey&query=$encodedTitle"

        if ($Year) {
            $url += "&year=$Year"
        }

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            $movie = $response.results[0]

            return @{
                Id = $movie.id
                Title = $movie.title
                OriginalTitle = $movie.original_title
                Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
                Overview = $movie.overview
                Rating = $movie.vote_average
                Votes = $movie.vote_count
                PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
                BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
            }
        }

        return $null
    }
    catch {
        Write-Log "Error searching TMDB: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Gets detailed movie information from TMDB by ID
.PARAMETER MovieId
    The TMDB movie ID
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with detailed movie metadata
#>
function Get-TMDBMovieDetails {
    param(
        [int]$MovieId,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/movie/$MovieId`?api_key=$ApiKey&append_to_response=credits,external_ids"
        $movie = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        $directors = @()
        $cast = @()

        if ($movie.credits) {
            $directors = $movie.credits.crew | Where-Object { $_.job -eq 'Director' } | Select-Object -ExpandProperty name
            $cast = $movie.credits.cast | Select-Object -First 10 | ForEach-Object {
                @{
                    Name = $_.name
                    Role = $_.character
                    Thumb = if ($_.profile_path) { "https://image.tmdb.org/t/p/w185$($_.profile_path)" } else { $null }
                }
            }
        }

        return @{
            Id = $movie.id
            Title = $movie.title
            OriginalTitle = $movie.original_title
            Tagline = $movie.tagline
            Year = if ($movie.release_date) { $movie.release_date.Substring(0,4) } else { $null }
            ReleaseDate = $movie.release_date
            Overview = $movie.overview
            Rating = $movie.vote_average
            Votes = $movie.vote_count
            Runtime = $movie.runtime
            Genres = $movie.genres | Select-Object -ExpandProperty name
            Studios = $movie.production_companies | Select-Object -ExpandProperty name
            Directors = $directors
            Cast = $cast
            IMDBID = $movie.external_ids.imdb_id
            PosterPath = if ($movie.poster_path) { "https://image.tmdb.org/t/p/w500$($movie.poster_path)" } else { $null }
            BackdropPath = if ($movie.backdrop_path) { "https://image.tmdb.org/t/p/original$($movie.backdrop_path)" } else { $null }
        }
    }
    catch {
        Write-Log "Error getting TMDB movie details: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Searches TMDB for a TV show by title
.PARAMETER Title
    The TV show title to search for
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with TV show metadata from TMDB
#>
function Search-TMDBTVShow {
    param(
        [string]$Title,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $encodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)
        $url = "https://api.themoviedb.org/3/search/tv?api_key=$ApiKey&query=$encodedTitle"

        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.results -and $response.results.Count -gt 0) {
            $show = $response.results[0]

            return @{
                Id = $show.id
                Title = $show.name
                OriginalTitle = $show.original_name
                FirstAirDate = $show.first_air_date
                Year = if ($show.first_air_date) { $show.first_air_date.Substring(0,4) } else { $null }
                Overview = $show.overview
                Rating = $show.vote_average
                PosterPath = if ($show.poster_path) { "https://image.tmdb.org/t/p/w500$($show.poster_path)" } else { $null }
            }
        }

        return $null
    }
    catch {
        Write-Log "Error searching TMDB TV: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Gets episode details from TMDB
.PARAMETER ShowId
    The TMDB TV show ID
.PARAMETER Season
    The season number
.PARAMETER Episode
    The episode number
.PARAMETER ApiKey
    TMDB API key
.OUTPUTS
    Hashtable with episode metadata
#>
function Get-TMDBEpisode {
    param(
        [int]$ShowId,
        [int]$Season,
        [int]$Episode,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        return $null
    }

    try {
        $url = "https://api.themoviedb.org/3/tv/$ShowId/season/$Season/episode/$Episode`?api_key=$ApiKey"
        $ep = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        return @{
            Title = $ep.name
            Overview = $ep.overview
            AirDate = $ep.air_date
            Season = $ep.season_number
            Episode = $ep.episode_number
            Rating = $ep.vote_average
            StillPath = if ($ep.still_path) { "https://image.tmdb.org/t/p/w300$($ep.still_path)" } else { $null }
        }
    }
    catch {
        Write-Log "Error getting TMDB episode: $_" "ERROR"
        return $null
    }
}

<#
.SYNOPSIS
    Creates a Kodi NFO file with TMDB metadata
.PARAMETER VideoPath
    Path to the video file
.PARAMETER Metadata
    Hashtable with TMDB metadata
#>
function New-MovieNFOFromTMDB {
    param(
        [string]$VideoPath,
        [hashtable]$Metadata
    )

    try {
        $videoFile = Get-Item $VideoPath -ErrorAction Stop
        $nfoPath = Join-Path $videoFile.DirectoryName "$([System.IO.Path]::GetFileNameWithoutExtension($videoFile.Name)).nfo"

        if ($script:Config.DryRun) {
            Write-Host "[DRY-RUN] Would create NFO for: $($Metadata.Title)" -ForegroundColor Yellow
            Write-Log "Would create NFO with TMDB data for: $($Metadata.Title)" "DRY-RUN"
            return
        }

        $genreXml = ($Metadata.Genres | ForEach-Object { "    <genre>$([System.Security.SecurityElement]::Escape($_))</genre>" }) -join "`n"
        $studioXml = ($Metadata.Studios | Select-Object -First 3 | ForEach-Object { "    <studio>$([System.Security.SecurityElement]::Escape($_))</studio>" }) -join "`n"
        $directorXml = ($Metadata.Directors | ForEach-Object { "    <director>$([System.Security.SecurityElement]::Escape($_))</director>" }) -join "`n"

        $actorXml = ""
        if ($Metadata.Cast) {
            $actorXml = ($Metadata.Cast | ForEach-Object {
                @"
    <actor>
        <name>$([System.Security.SecurityElement]::Escape($_.Name))</name>
        <role>$([System.Security.SecurityElement]::Escape($_.Role))</role>
        $(if($_.Thumb){"<thumb>$($_.Thumb)</thumb>"})
    </actor>
"@
            }) -join "`n"
        }

        $nfoContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<movie>
    <title>$([System.Security.SecurityElement]::Escape($Metadata.Title))</title>
    <originaltitle>$([System.Security.SecurityElement]::Escape($Metadata.OriginalTitle))</originaltitle>
    <year>$($Metadata.Year)</year>
    <plot>$([System.Security.SecurityElement]::Escape($Metadata.Overview))</plot>
    <tagline>$([System.Security.SecurityElement]::Escape($Metadata.Tagline))</tagline>
    <runtime>$($Metadata.Runtime)</runtime>
    <rating>$($Metadata.Rating)</rating>
    <votes>$($Metadata.Votes)</votes>
    <uniqueid type="tmdb" default="true">$($Metadata.Id)</uniqueid>
    $(if($Metadata.IMDBID){"<uniqueid type=`"imdb`">$($Metadata.IMDBID)</uniqueid>"})
    <thumb aspect="poster">$($Metadata.PosterPath)</thumb>
    <fanart>
        <thumb>$($Metadata.BackdropPath)</thumb>
    </fanart>
$genreXml
$studioXml
$directorXml
$actorXml
</movie>
"@

        $nfoContent | Out-File -FilePath $nfoPath -Encoding UTF8 -Force
        Write-Host "Created NFO with TMDB data: $($Metadata.Title)" -ForegroundColor Green
        Write-Log "Created NFO file with TMDB data: $nfoPath" "INFO"
        $script:Stats.NFOFilesCreated++
    }
    catch {
        Write-Host "Error creating NFO: $_" -ForegroundColor Red
        Write-Log "Error creating NFO with TMDB data: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Fetches metadata from TMDB for all movies in a library
.PARAMETER Path
    The root path of the movie library
.PARAMETER ApiKey
    TMDB API key
#>
function Invoke-TMDBMetadataFetch {
    param(
        [string]$Path,
        [string]$ApiKey
    )

    if (-not $ApiKey) {
        Write-Host "TMDB API key required. Get one free at https://www.themoviedb.org/settings/api" -ForegroundColor Yellow
        return
    }

    Write-Host "`nFetching metadata from TMDB..." -ForegroundColor Yellow
    Write-Log "Starting TMDB metadata fetch for: $Path" "INFO"

    try {
        $movieFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne '_Trailers' }

        if ($movieFolders.Count -eq 0) {
            Write-Host "No movie folders found" -ForegroundColor Cyan
            return
        }

        $processed = 0
        $found = 0
        $total = $movieFolders.Count

        foreach ($folder in $movieFolders) {
            $processed++
            Write-Progress -Activity "Fetching TMDB metadata" -Status "[$processed/$total] $($folder.Name)" -PercentComplete (($processed / $total) * 100)

            # Check if NFO already exists with TMDB data
            $existingNfo = Get-ChildItem -Path $folder.FullName -Filter "*.nfo" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($existingNfo) {
                $nfoContent = Get-Content $existingNfo.FullName -Raw -ErrorAction SilentlyContinue
                if ($nfoContent -match 'uniqueid type="tmdb"') {
                    Write-Host "Skipping (has TMDB data): $($folder.Name)" -ForegroundColor Gray
                    continue
                }
            }

            # Get video file
            $videoFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
                Select-Object -First 1

            if (-not $videoFile) {
                continue
            }

            # Extract title and year from folder name
            $titleInfo = Get-NormalizedTitle -Name $folder.Name
            $searchTitle = $titleInfo.NormalizedTitle -replace '\s+', ' '

            # Search TMDB
            $searchResult = Search-TMDBMovie -Title $searchTitle -Year $titleInfo.Year -ApiKey $ApiKey

            if ($searchResult) {
                # Get detailed info
                $details = Get-TMDBMovieDetails -MovieId $searchResult.Id -ApiKey $ApiKey

                if ($details) {
                    Write-Host "Found: $($details.Title) ($($details.Year))" -ForegroundColor Green
                    New-MovieNFOFromTMDB -VideoPath $videoFile.FullName -Metadata $details
                    $found++
                }
            } else {
                Write-Host "Not found: $($folder.Name)" -ForegroundColor Yellow
            }

            # Rate limiting - TMDB allows 40 requests per 10 seconds
            Start-Sleep -Milliseconds 300
        }

        Write-Progress -Activity "Fetching TMDB metadata" -Completed
        Write-Host "`nTMDB fetch complete: $found/$total movies matched" -ForegroundColor Cyan
        Write-Log "TMDB metadata fetch completed: $found/$total matched" "INFO"
    }
    catch {
        Write-Host "Error during TMDB fetch: $_" -ForegroundColor Red
        Write-Log "Error during TMDB fetch: $_" "ERROR"
    }
}

#============================================
# ARCHIVE FUNCTIONS
#============================================

<#
.SYNOPSIS
    Extracts all archives in the specified path
.PARAMETER Path
    The root path to search for archives
.PARAMETER DeleteAfterExtract
    Whether to delete archives after successful extraction
#>
function Expand-Archives {
    param(
        [string]$Path,
        [switch]$DeleteAfterExtract
    )

    Write-Host "Extracting archives..." -ForegroundColor Yellow
    Write-Log "Starting archive extraction in: $Path" "INFO"

    try {
        Set-Alias sz $script:Config.SevenZipPath -Scope Script

        $unzipQueue = @()
        foreach ($ext in $script:Config.ArchiveExtensions) {
            $archives = Get-ChildItem -Path $Path -Filter $ext -Recurse -ErrorAction SilentlyContinue
            if ($archives) {
                $unzipQueue += $archives
            }
        }

        $count = $unzipQueue.Count

        if ($count -gt 0) {
            Write-Host "Found $count archive(s) to extract" -ForegroundColor Cyan
            Write-Log "Found $count archive(s) to extract" "INFO"

            $current = 1
            foreach ($archive in $unzipQueue) {
                $percentComplete = ($current / $count) * 100
                Write-Progress -Activity "Extracting archives" -Status "[$current/$count] $($archive.Name)" -PercentComplete $percentComplete

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] [$current/$count] Would extract: $($archive.Name)" -ForegroundColor Yellow
                    Write-Log "Would extract: $($archive.FullName)" "DRY-RUN"
                } else {
                    # Extract to archive's parent directory instead of root path
                    $extractPath = $archive.DirectoryName
                    $archiveSize = $archive.Length
                    $archiveSizeMB = [math]::Round($archiveSize / 1MB, 2)

                    Write-Host "[$current/$count] Extracting: $($archive.Name) ($archiveSizeMB MB) to $extractPath" -ForegroundColor Cyan
                    Write-Log "Extracting: $($archive.FullName) ($archiveSizeMB MB)" "INFO"

                    # Use Start-Process with timeout for better control
                    $timeoutSeconds = 1800  # 30 minute timeout
                    $process = Start-Process -FilePath $script:Config.SevenZipPath `
                        -ArgumentList "x", "-o`"$extractPath`"", "`"$($archive.FullName)`"", "-r", "-y" `
                        -NoNewWindow -PassThru -Wait:$false

                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    $lastUpdate = 0

                    # Wait for process with progress updates
                    while (-not $process.HasExited) {
                        Start-Sleep -Milliseconds 500
                        $elapsedSeconds = [math]::Floor($stopwatch.Elapsed.TotalSeconds)

                        # Update every 5 seconds
                        if ($elapsedSeconds -ge ($lastUpdate + 5)) {
                            $lastUpdate = $elapsedSeconds
                            Write-Host "  Extracting... $elapsedSeconds seconds elapsed" -ForegroundColor Gray
                        }

                        # Check timeout
                        if ($stopwatch.Elapsed.TotalSeconds -gt $timeoutSeconds) {
                            Write-Host "  Timeout reached ($timeoutSeconds seconds), killing extraction process" -ForegroundColor Red
                            Write-Log "Extraction timeout for: $($archive.FullName)" "ERROR"
                            $process.Kill()
                            $script:Stats.ArchivesFailed++
                            $stopwatch.Stop()
                            continue
                        }
                    }
                    $stopwatch.Stop()

                    $exitCode = $process.ExitCode
                    $elapsedTime = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)

                    if ($exitCode -ne 0) {
                        Write-Host "Warning: Failed to extract $($archive.Name) (exit code: $exitCode)" -ForegroundColor Yellow
                        Write-Log "Failed to extract: $($archive.FullName) (exit code: $exitCode)" "WARNING"
                        $script:Stats.ArchivesFailed++
                    } else {
                        Write-Host "Successfully extracted $($archive.Name) in $elapsedTime seconds" -ForegroundColor Green
                        Write-Log "Successfully extracted: $($archive.FullName) in $elapsedTime seconds" "INFO"
                        $script:Stats.ArchivesExtracted++
                    }
                }
                $current++
            }
            Write-Progress -Activity "Extracting archives" -Completed

            # Delete archives after extraction
            if ($DeleteAfterExtract -and -not $script:Config.DryRun) {
                Write-Host "Deleting archive files..." -ForegroundColor Yellow
                foreach ($pattern in $script:Config.ArchiveCleanupPatterns) {
                    $archivesToDelete = Get-ChildItem -Path $Path -Include $pattern -Recurse -ErrorAction SilentlyContinue
                    foreach ($archiveFile in $archivesToDelete) {
                        $archiveSize = $archiveFile.Length
                        Remove-Item -Path $archiveFile.FullName -Force -ErrorAction SilentlyContinue
                        Write-Log "Deleted archive: $($archiveFile.FullName)" "INFO"
                        $script:Stats.FilesDeleted++
                        $script:Stats.BytesDeleted += $archiveSize
                    }
                }
            } elseif ($script:Config.DryRun) {
                Write-Host "[DRY-RUN] Would delete all archive files after extraction" -ForegroundColor Yellow
            }

            Write-Host "Archives processed" -ForegroundColor Green
        } else {
            Write-Host "No archives found to extract" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error processing archives: $_" -ForegroundColor Red
        Write-Log "Error processing archives: $_" "ERROR"
    }
}

#============================================
# FOLDER ORGANIZATION FUNCTIONS
#============================================

<#
.SYNOPSIS
    Creates individual folders for loose video files
.PARAMETER Path
    The root path containing loose files
#>
function New-FoldersForLooseFiles {
    param(
        [string]$Path
    )

    Write-Host "Creating folders for loose video files..." -ForegroundColor Yellow
    Write-Log "Starting folder creation for loose files in: $Path" "INFO"

    try {
        $looseFiles = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($looseFiles) {
            foreach ($file in $looseFiles) {
                try {
                    $dir = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $fullDirPath = Join-Path -Path $Path -ChildPath $dir

                    if ($script:Config.DryRun) {
                        Write-Host "[DRY-RUN] Would create folder '$dir' and move: $($file.Name)" -ForegroundColor Yellow
                        Write-Log "Would create folder '$dir' and move: $($file.Name)" "DRY-RUN"
                    } else {
                        Write-Host "Processing: $($file.Name)" -ForegroundColor Cyan

                        if (-not (Test-Path $fullDirPath)) {
                            New-Item -Path $fullDirPath -ItemType Directory -ErrorAction Stop | Out-Null
                            $script:Stats.FoldersCreated++
                        }

                        Move-Item -Path $file.FullName -Destination $fullDirPath -Force -ErrorAction Stop
                        Write-Host "Moved to: $dir" -ForegroundColor Green
                        Write-Log "Moved '$($file.Name)' to folder '$dir'" "INFO"
                        $script:Stats.FilesMoved++
                    }
                }
                catch {
                    Write-Host "Warning: Could not process $($file.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error processing $($file.Name): $_" "ERROR"
                }
            }
        } else {
            Write-Host "No loose video files found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error organizing loose files: $_" -ForegroundColor Red
        Write-Log "Error organizing loose files: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Cleans folder names by removing tags and formatting
.PARAMETER Path
    The root path containing folders to clean
#>
function Rename-CleanFolderNames {
    param(
        [string]$Path
    )

    Write-Host "Cleaning folder names..." -ForegroundColor Yellow
    Write-Log "Starting folder name cleaning in: $Path" "INFO"

    try {
        # Remove tags from folder names
        foreach ($tag in $script:Config.Tags) {
            Get-ChildItem -Path $Path -Directory -Filter "*$tag*" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    try {
                        $newName = ($_.Name -split [regex]::Escape($tag))[0].TrimEnd(' ', '.', '-')
                        if ($newName -and $newName -ne $_.Name) {
                            if ($script:Config.DryRun) {
                                Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                                Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                            } else {
                                Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                                Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                                $script:Stats.FoldersRenamed++
                            }
                        }
                    }
                    catch {
                        Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error renaming $($_.Name): $_" "ERROR"
                    }
                }
        }

        # Replace dots with spaces in folder names
        Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\.' } |
            ForEach-Object {
                try {
                    $newName = $_.Name -replace '\.', ' '
                    $newName = $newName -replace '\s+', ' '  # Remove multiple spaces
                    $newName = $newName.Trim()

                    if ($newName -ne $_.Name) {
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                            Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                        } else {
                            Rename-Item -Path $_.FullName -NewName $newName -Force -ErrorAction Stop
                            Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                            $script:Stats.FoldersRenamed++
                        }
                    }
                }
                catch {
                    Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming $($_.Name): $_" "ERROR"
                }
            }

        # Format years with parentheses (both 19xx and 20xx)
        Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\s(19|20)\d{2}$' } |
            ForEach-Object {
                try {
                    $newName = $_.Name -replace '\s((19|20)\d{2})$', ' ($1)'

                    if ($newName -ne $_.Name) {
                        if ($script:Config.DryRun) {
                            Write-Host "[DRY-RUN] Would rename '$($_.Name)' to '$newName'" -ForegroundColor Yellow
                            Write-Log "Would rename '$($_.Name)' to '$newName'" "DRY-RUN"
                        } else {
                            Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                            Write-Log "Renamed '$($_.Name)' to '$newName'" "INFO"
                            $script:Stats.FoldersRenamed++
                        }
                    }
                }
                catch {
                    Write-Host "Warning: Could not rename $($_.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error renaming $($_.Name): $_" "ERROR"
                }
            }

        if ($script:Stats.FoldersRenamed -gt 0 -or $script:Config.DryRun) {
            Write-Host "Folder names cleaned" -ForegroundColor Green
        } else {
            Write-Host "No folders needed renaming" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error cleaning folder names: $_" -ForegroundColor Red
        Write-Log "Error cleaning folder names: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Removes empty folders from the specified path
.PARAMETER Path
    The root path to search for empty folders
#>
function Remove-EmptyFolders {
    param(
        [string]$Path
    )

    Write-Host "Cleaning up empty folders..." -ForegroundColor Yellow
    Write-Log "Starting empty folder cleanup in: $Path" "INFO"

    try {
        # Get all empty directories (recursively, deepest first)
        $emptyFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { (Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0 } |
            Sort-Object { $_.FullName.Length } -Descending

        if ($emptyFolders) {
            Write-Host "Found $($emptyFolders.Count) empty folder(s)" -ForegroundColor Cyan

            foreach ($folder in $emptyFolders) {
                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would remove: $($folder.FullName)" -ForegroundColor Yellow
                    Write-Log "Would remove empty folder: $($folder.FullName)" "DRY-RUN"
                } else {
                    Remove-Item -Path $folder.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed empty folder: $($folder.FullName)" "INFO"
                    $script:Stats.EmptyFoldersRemoved++
                }
            }
            Write-Host "Empty folders removed" -ForegroundColor Green
        } else {
            Write-Host "No empty folders found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Warning: Could not remove all empty folders: $_" -ForegroundColor Yellow
        Write-Log "Error removing empty folders: $_" "ERROR"
    }
}

#============================================
# TV SHOW FUNCTIONS
#============================================

<#
.SYNOPSIS
    Parses a filename to extract season and episode information
.PARAMETER FileName
    The filename to parse
.OUTPUTS
    Hashtable with Season, Episode, Episodes (for multi-ep), Title, and EpisodeTitle
.DESCRIPTION
    Supports multiple naming formats:
    - S01E01, s01e01
    - S01E01E02, S01E01-E03 (multi-episode)
    - 1x01, 01x01
    - Season 1 Episode 1
    - Part 1, Pt.1
#>
function Get-EpisodeInfo {
    param(
        [string]$FileName
    )

    $info = @{
        Season = $null
        Episode = $null
        Episodes = @()      # For multi-episode files
        ShowTitle = $null
        EpisodeTitle = $null
        IsMultiEpisode = $false
    }

    # Remove extension
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    # Pattern 1: S01E01 or S01E01E02 or S01E01-E03 (most common)
    if ($baseName -match '^(.+?)[.\s_-]+[Ss](\d{1,2})[Ee](\d{1,2})(?:[Ee-](\d{1,2}))?(?:[Ee-](\d{1,2}))?(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]

        # Check for multi-episode
        if ($Matches[4]) {
            $info.Episodes += [int]$Matches[4]
            $info.IsMultiEpisode = $true
        }
        if ($Matches[5]) {
            $info.Episodes += [int]$Matches[5]
        }

        # Try to get episode title from remainder
        if ($Matches[6]) {
            $remainder = $Matches[6] -replace '^[.\s_-]+', '' -replace '\.', ' '
            # Remove quality tags from episode title
            $remainder = $remainder -replace '\s*(720p|1080p|2160p|4K|HDTV|WEB-DL|WEBRip|BluRay|x264|x265|HEVC|AAC|AC3).*$', ''
            if ($remainder.Trim()) {
                $info.EpisodeTitle = $remainder.Trim()
            }
        }
    }
    # Pattern 2: 1x01 format
    elseif ($baseName -match '^(.+?)[.\s_-]+(\d{1,2})x(\d{1,2})(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }
    # Pattern 3: Season 1 Episode 1
    elseif ($baseName -match '^(.+?)[.\s_-]+Season\s*(\d{1,2})[.\s_-]+Episode\s*(\d{1,2})(.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }
    # Pattern 4: Show.Name.101 (season 1, episode 01)
    elseif ($baseName -match '^(.+?)[.\s_-]+(\d)(\d{2})[.\s_-](.*)$') {
        $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
        $info.Season = [int]$Matches[2]
        $info.Episode = [int]$Matches[3]
        $info.Episodes += [int]$Matches[3]
    }

    # Clean up show title
    if ($info.ShowTitle) {
        $info.ShowTitle = $info.ShowTitle.Trim()
    }

    return $info
}

<#
.SYNOPSIS
    Organizes TV show files into Season folders
.PARAMETER Path
    The root path of the TV show
.DESCRIPTION
    Scans video files, extracts season/episode info, and organizes into Season XX folders
#>
function Invoke-SeasonOrganization {
    param(
        [string]$Path
    )

    Write-Host "Organizing episodes into season folders..." -ForegroundColor Yellow
    Write-Log "Starting season organization in: $Path" "INFO"

    try {
        # Find all video files
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Cyan

        $organized = 0
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name

            if ($null -ne $epInfo.Season) {
                $seasonFolder = "Season {0:D2}" -f $epInfo.Season
                $seasonPath = Join-Path $Path $seasonFolder

                # Skip if already in correct season folder
                if ($file.Directory.Name -eq $seasonFolder) {
                    Write-Host "Already organized: $($file.Name)" -ForegroundColor Gray
                    continue
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would move to $seasonFolder : $($file.Name)" -ForegroundColor Yellow
                    Write-Log "Would move $($file.Name) to $seasonFolder" "DRY-RUN"
                } else {
                    # Create season folder if needed
                    if (-not (Test-Path $seasonPath)) {
                        New-Item -Path $seasonPath -ItemType Directory -Force | Out-Null
                        Write-Host "Created folder: $seasonFolder" -ForegroundColor Green
                        Write-Log "Created season folder: $seasonPath" "INFO"
                        $script:Stats.FoldersCreated++
                    }

                    # Move file
                    $destPath = Join-Path $seasonPath $file.Name
                    try {
                        Move-Item -Path $file.FullName -Destination $destPath -Force -ErrorAction Stop
                        Write-Host "Moved to $seasonFolder : $($file.Name)" -ForegroundColor Cyan
                        Write-Log "Moved $($file.Name) to $seasonFolder" "INFO"
                        $script:Stats.FilesMoved++
                        $organized++
                    }
                    catch {
                        Write-Host "Warning: Could not move $($file.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error moving $($file.FullName): $_" "WARNING"
                    }
                }
            } else {
                Write-Host "Could not parse: $($file.Name)" -ForegroundColor Yellow
                Write-Log "Could not parse season/episode from: $($file.Name)" "WARNING"
            }
        }

        if ($organized -gt 0) {
            Write-Host "Organized $organized episode(s) into season folders" -ForegroundColor Green
        } else {
            Write-Host "No episodes needed organizing" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error organizing seasons: $_" -ForegroundColor Red
        Write-Log "Error organizing seasons: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Renames episode files to a consistent format
.PARAMETER Path
    The root path of the TV show
.PARAMETER Format
    The naming format to use (default: "{ShowTitle} - S{Season:D2}E{Episode:D2}")
.DESCRIPTION
    Standardizes episode filenames while preserving episode information
#>
function Rename-EpisodeFiles {
    param(
        [string]$Path,
        [string]$Format = "{0} - S{1:D2}E{2:D2}"
    )

    Write-Host "Standardizing episode filenames..." -ForegroundColor Yellow
    Write-Log "Starting episode rename in: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        $renamed = 0
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name

            if ($null -ne $epInfo.Season -and $epInfo.ShowTitle) {
                # Build new filename
                if ($epInfo.IsMultiEpisode) {
                    # Multi-episode: Show - S01E01-E03
                    $episodePart = "E" + ($epInfo.Episodes | ForEach-Object { "{0:D2}" -f $_ }) -join "-E"
                    $newName = "{0} - S{1:D2}{2}" -f $epInfo.ShowTitle, $epInfo.Season, $episodePart
                } else {
                    $newName = $Format -f $epInfo.ShowTitle, $epInfo.Season, $epInfo.Episode
                }

                # Add episode title if available
                if ($epInfo.EpisodeTitle) {
                    $newName += " - $($epInfo.EpisodeTitle)"
                }

                $newName += $file.Extension

                # Skip if already correctly named
                if ($file.Name -eq $newName) {
                    continue
                }

                if ($script:Config.DryRun) {
                    Write-Host "[DRY-RUN] Would rename:" -ForegroundColor Yellow
                    Write-Host "  From: $($file.Name)" -ForegroundColor Gray
                    Write-Host "  To:   $newName" -ForegroundColor Cyan
                    Write-Log "Would rename $($file.Name) to $newName" "DRY-RUN"
                } else {
                    try {
                        Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                        Write-Host "Renamed: $($file.Name) -> $newName" -ForegroundColor Cyan
                        Write-Log "Renamed $($file.Name) to $newName" "INFO"
                        $script:Stats.FoldersRenamed++  # Reusing stat for file renames
                        $renamed++
                    }
                    catch {
                        Write-Host "Warning: Could not rename $($file.Name): $_" -ForegroundColor Yellow
                        Write-Log "Error renaming $($file.FullName): $_" "WARNING"
                    }
                }
            }
        }

        if ($renamed -gt 0) {
            Write-Host "Renamed $renamed episode file(s)" -ForegroundColor Green
        } else {
            Write-Host "No episodes needed renaming" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error renaming episodes: $_" -ForegroundColor Red
        Write-Log "Error renaming episodes: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Displays a summary of detected TV show episodes
.PARAMETER Path
    The root path of the TV show
#>
function Show-EpisodeSummary {
    param(
        [string]$Path
    )

    Write-Host "`nEpisode Summary:" -ForegroundColor Cyan
    Write-Log "Generating episode summary for: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() }

        if ($videoFiles.Count -eq 0) {
            Write-Host "No video files found" -ForegroundColor Cyan
            return
        }

        # Group by season
        $seasons = @{}
        foreach ($file in $videoFiles) {
            $epInfo = Get-EpisodeInfo -FileName $file.Name
            if ($null -ne $epInfo.Season) {
                $seasonKey = $epInfo.Season
                if (-not $seasons.ContainsKey($seasonKey)) {
                    $seasons[$seasonKey] = @()
                }
                $seasons[$seasonKey] += @{
                    Episode = $epInfo.Episode
                    Episodes = $epInfo.Episodes
                    FileName = $file.Name
                    IsMulti = $epInfo.IsMultiEpisode
                }
            }
        }

        if ($seasons.Count -eq 0) {
            Write-Host "Could not parse any episode information" -ForegroundColor Yellow
            return
        }

        # Display summary
        foreach ($season in ($seasons.Keys | Sort-Object)) {
            $episodes = $seasons[$season] | Sort-Object { $_.Episode }
            $episodeNums = ($episodes | ForEach-Object {
                if ($_.IsMulti) {
                    "E" + ($_.Episodes -join "-E")
                } else {
                    "E{0:D2}" -f $_.Episode
                }
            }) -join ", "

            Write-Host "  Season $season : $($episodes.Count) episode(s) - $episodeNums" -ForegroundColor White
        }

        # Check for gaps
        foreach ($season in ($seasons.Keys | Sort-Object)) {
            $episodes = $seasons[$season] | ForEach-Object { $_.Episode } | Sort-Object -Unique
            $expected = 1..($episodes | Measure-Object -Maximum).Maximum
            $missing = $expected | Where-Object { $_ -notin $episodes }
            if ($missing) {
                Write-Host "  Season $season missing: E$($missing -join ', E')" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Error generating summary: $_" -ForegroundColor Red
        Write-Log "Error generating episode summary: $_" "ERROR"
    }
}

<#
.SYNOPSIS
    Moves all video files to the root directory
.PARAMETER Path
    The root path to search for video files
#>
function Move-VideoFilesToRoot {
    param(
        [string]$Path
    )

    Write-Host "Moving video files to root..." -ForegroundColor Yellow
    Write-Log "Starting video file move to root in: $Path" "INFO"

    try {
        $videoFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $script:Config.VideoExtensions -contains $_.Extension.ToLower() } |
            Where-Object { $_.DirectoryName -ne $Path }  # Exclude files already in root

        if ($videoFiles) {
            Write-Host "Found $($videoFiles.Count) video file(s)" -ForegroundColor Cyan
            Write-Log "Found $($videoFiles.Count) video file(s)" "INFO"

            $current = 1
            foreach ($file in $videoFiles) {
                $percentComplete = ($current / $videoFiles.Count) * 100
                Write-Progress -Activity "Moving video files" -Status "[$current/$($videoFiles.Count)] $($file.Name)" -PercentComplete $percentComplete

                try {
                    if ($script:Config.DryRun) {
                        Write-Host "[DRY-RUN] [$current/$($videoFiles.Count)] Would move: $($file.Name)" -ForegroundColor Yellow
                        Write-Log "Would move: $($file.FullName) to root" "DRY-RUN"
                    } else {
                        Move-Item -Path $file.FullName -Destination $Path -Force -ErrorAction Stop
                        Write-Host "[$current/$($videoFiles.Count)] Moved: $($file.Name)" -ForegroundColor Cyan
                        Write-Log "Moved: $($file.FullName) to root" "INFO"
                        $script:Stats.FilesMoved++
                    }
                }
                catch {
                    Write-Host "Warning: Could not move $($file.Name): $_" -ForegroundColor Yellow
                    Write-Log "Error moving $($file.FullName): $_" "ERROR"
                }
                $current++
            }
            Write-Progress -Activity "Moving video files" -Completed
            Write-Host "Video files processed" -ForegroundColor Green
        } else {
            Write-Host "No video files found to move" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Error moving video files: $_" -ForegroundColor Red
        Write-Log "Error moving video files: $_" "ERROR"
    }
}

#============================================
# MAIN PROCESSING FUNCTIONS
#============================================

<#
.SYNOPSIS
    Processes a movie library
.PARAMETER Path
    The root path of the movie library
#>
function Invoke-MovieProcessing {
    param(
        [string]$Path
    )

    Write-Host "`nMovie Routine" -ForegroundColor Magenta
    Write-Log "Movie routine started for path: $Path" "INFO"

    if ($script:Config.DryRun) {
        Write-Host "`n*** DRY-RUN MODE - Previewing changes only ***`n" -ForegroundColor Yellow
    }

    # Step 1: Verify 7-Zip
    if (-not (Test-SevenZipInstallation)) {
        Write-Host "Cannot proceed without 7-Zip" -ForegroundColor Red
        return
    }

    # Step 2: Process trailers (move to _Trailers folder before deletion)
    if ($script:Config.KeepTrailers) {
        Move-TrailersToFolder -Path $Path
    }

    # Step 3: Remove unnecessary files
    Remove-UnnecessaryFiles -Path $Path

    # Step 4: Process subtitles
    Invoke-SubtitleProcessing -Path $Path

    # Step 5: Extract archives
    Expand-Archives -Path $Path -DeleteAfterExtract

    # Step 6: Create folders for loose files
    New-FoldersForLooseFiles -Path $Path

    # Step 7: Clean folder names
    Rename-CleanFolderNames -Path $Path

    # Step 8: Generate NFO files (if enabled)
    Invoke-NFOGeneration -Path $Path

    # Step 9: Show existing NFO metadata summary
    Show-NFOMetadata -Path $Path

    # Step 10: Check for duplicates (if enabled)
    if ($script:Config.CheckDuplicates) {
        Show-DuplicateReport -Path $Path

        $removeDupes = Read-Host "`nRemove duplicate movies? (Y/N) [N]"
        if ($removeDupes -eq 'Y' -or $removeDupes -eq 'y') {
            Remove-DuplicateMovies -Path $Path
        }
    }

    Write-Host "`nMovie routine completed!" -ForegroundColor Magenta
    Write-Log "Movie routine completed" "INFO"
}

<#
.SYNOPSIS
    Processes a TV show library
.PARAMETER Path
    The root path of the TV show library
#>
function Invoke-TVShowProcessing {
    param(
        [string]$Path
    )

    Write-Host "`nTV Show Routine" -ForegroundColor Magenta
    Write-Log "TV Show routine started for path: $Path" "INFO"

    if ($script:Config.DryRun) {
        Write-Host "`n*** DRY-RUN MODE - Previewing changes only ***`n" -ForegroundColor Yellow
    }

    # Step 1: Verify 7-Zip
    if (-not (Test-SevenZipInstallation)) {
        Write-Host "Cannot proceed without 7-Zip" -ForegroundColor Red
        return
    }

    # Step 2: Extract archives first (may contain episodes)
    Expand-Archives -Path $Path -DeleteAfterExtract

    # Step 3: Remove unnecessary files (samples, proofs, etc.)
    Remove-UnnecessaryFiles -Path $Path

    # Step 4: Process subtitles
    Invoke-SubtitleProcessing -Path $Path

    # Step 5: Organize into season folders (if enabled)
    if ($script:Config.OrganizeSeasons) {
        Invoke-SeasonOrganization -Path $Path
    }

    # Step 6: Rename episode files (if enabled)
    if ($script:Config.RenameEpisodes) {
        Rename-EpisodeFiles -Path $Path
    }

    # Step 7: Remove empty folders
    Remove-EmptyFolders -Path $Path

    # Step 8: Show episode summary
    Show-EpisodeSummary -Path $Path

    Write-Host "`nTV Show routine completed!" -ForegroundColor Magenta
    Write-Log "TV Show routine completed" "INFO"
}

#============================================
# MAIN SCRIPT EXECUTION
#============================================

# Initialize statistics
$script:Stats.StartTime = Get-Date

# Display header
Write-Host "`n=== MediaCleaner v4.1 ===" -ForegroundColor Cyan
Write-Host "Automated media library organization tool`n" -ForegroundColor Gray

# User configuration
Write-Host "=== Configuration ===" -ForegroundColor Cyan
$dryRunInput = Read-Host "Enable dry-run mode? (Y/N) [N]"
if ($dryRunInput -eq 'Y' -or $dryRunInput -eq 'y') {
    $script:Config.DryRun = $true
    Write-Host "DRY-RUN MODE ENABLED - No changes will be made" -ForegroundColor Yellow
    Write-Log "Dry-run mode enabled" "INFO"
} else {
    Write-Host "Live mode - changes will be applied" -ForegroundColor Green
    Write-Log "Live mode enabled" "INFO"
}

# Subtitle configuration
$subtitleInput = Read-Host "Keep subtitle files? (Y/N) [Y]"
if ($subtitleInput -eq 'N' -or $subtitleInput -eq 'n') {
    $script:Config.KeepSubtitles = $false
    Write-Host "Subtitles will be deleted" -ForegroundColor Yellow
    Write-Log "Subtitle keeping disabled" "INFO"
} else {
    Write-Host "English subtitles will be kept" -ForegroundColor Green
    Write-Log "Subtitle keeping enabled" "INFO"
}

# Trailer configuration
$trailerInput = Read-Host "Keep trailers (move to _Trailers folder)? (Y/N) [Y]"
if ($trailerInput -eq 'N' -or $trailerInput -eq 'n') {
    $script:Config.KeepTrailers = $false
    Write-Host "Trailers will be deleted" -ForegroundColor Yellow
    Write-Log "Trailer keeping disabled" "INFO"
} else {
    Write-Host "Trailers will be moved to _Trailers folder" -ForegroundColor Green
    Write-Log "Trailer keeping enabled" "INFO"
}

# NFO generation configuration
$nfoInput = Read-Host "Generate Kodi NFO files for movies without them? (Y/N) [N]"
if ($nfoInput -eq 'Y' -or $nfoInput -eq 'y') {
    $script:Config.GenerateNFO = $true
    Write-Host "NFO files will be generated" -ForegroundColor Green
    Write-Log "NFO generation enabled" "INFO"
} else {
    Write-Host "NFO generation skipped" -ForegroundColor Cyan
    Write-Log "NFO generation disabled" "INFO"
}

Write-Host "`nLog file: $($script:Config.LogFile)" -ForegroundColor Cyan
Write-Log "MediaCleaner session started" "INFO"

# Media type selection
Write-Host "`n=== Main Menu ===" -ForegroundColor Cyan
Write-Host "1. Process Movies"
Write-Host "2. Process TV Shows"
Write-Host "3. Health Check"
Write-Host "4. Codec Analysis"
Write-Host "5. TMDB Metadata Fetch"
Write-Host "6. Export Library Report"
Write-Host "7. Enhanced Duplicate Scan"
Write-Host "8. Undo Previous Session"
Write-Host "9. Save/Load Configuration"

$type = Read-Host "`nSelect option"
Write-Log "User selected type: $type" "INFO"

# Process based on selection
switch ($type) {
    "1" {
        # Movie specific configuration
        Write-Host "`n=== Movie Options ===" -ForegroundColor Cyan

        $dupeInput = Read-Host "Check for duplicate movies? (Y/N) [N]"
        if ($dupeInput -eq 'Y' -or $dupeInput -eq 'y') {
            $script:Config.CheckDuplicates = $true
            Write-Host "Duplicate detection enabled" -ForegroundColor Green
            Write-Log "Duplicate detection enabled" "INFO"
        } else {
            Write-Host "Duplicate detection disabled" -ForegroundColor Cyan
            Write-Log "Duplicate detection disabled" "INFO"
        }

        $path = Select-FolderDialog -Description "Select your Movies folder"
        if ($path) {
            Invoke-MovieProcessing -Path $path

            # Retry failed operations
            Invoke-RetryFailedOperations

            # Save undo manifest
            Save-UndoManifest

            # Offer to export report
            $exportReport = Read-Host "`nExport library report? (Y/N) [N]"
            if ($exportReport -eq 'Y' -or $exportReport -eq 'y') {
                Invoke-ExportMenu -Path $path -MediaType "Movies"
            }

            # Offer to save config
            Invoke-ConfigurationSavePrompt
        }
    }
    "2" {
        # TV Show specific configuration
        Write-Host "`n=== TV Show Options ===" -ForegroundColor Cyan

        $organizeInput = Read-Host "Organize episodes into Season folders? (Y/N) [Y]"
        if ($organizeInput -eq 'N' -or $organizeInput -eq 'n') {
            $script:Config.OrganizeSeasons = $false
            Write-Host "Season organization disabled" -ForegroundColor Cyan
            Write-Log "Season organization disabled" "INFO"
        } else {
            Write-Host "Episodes will be organized into Season folders" -ForegroundColor Green
            Write-Log "Season organization enabled" "INFO"
        }

        $renameInput = Read-Host "Rename episodes to standard format (ShowName - S01E01)? (Y/N) [N]"
        if ($renameInput -eq 'Y' -or $renameInput -eq 'y') {
            $script:Config.RenameEpisodes = $true
            Write-Host "Episodes will be renamed to standard format" -ForegroundColor Green
            Write-Log "Episode renaming enabled" "INFO"
        } else {
            Write-Host "Episode filenames will be preserved" -ForegroundColor Cyan
            Write-Log "Episode renaming disabled" "INFO"
        }

        $path = Select-FolderDialog -Description "Select your TV Shows folder"
        if ($path) {
            Invoke-TVShowProcessing -Path $path

            # Retry failed operations
            Invoke-RetryFailedOperations

            # Save undo manifest
            Save-UndoManifest

            # Offer to save config
            Invoke-ConfigurationSavePrompt
        }
    }
    "3" {
        # Health Check mode
        Write-Host "`n=== Health Check Mode ===" -ForegroundColor Cyan
        $mediaTypeInput = Read-Host "Media type? (1=Movies, 2=TV Shows) [1]"
        $mediaType = if ($mediaTypeInput -eq '2') { "TVShows" } else { "Movies" }

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            Invoke-LibraryHealthCheck -Path $path -MediaType $mediaType
        }
    }
    "4" {
        # Codec Analysis mode
        Write-Host "`n=== Codec Analysis Mode ===" -ForegroundColor Cyan
        Write-Host "MediaInfo integration: $(if (Test-MediaInfoInstallation) { 'Available' } else { 'Not found (using filename parsing)' })" -ForegroundColor $(if (Test-MediaInfoInstallation) { 'Green' } else { 'Yellow' })
        Write-Host "FFmpeg: $(if (Test-FFmpegInstallation) { 'Available' } else { 'Not found (required for transcoding)' })" -ForegroundColor $(if (Test-FFmpegInstallation) { 'Green' } else { 'Yellow' })

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            $analysis = Invoke-CodecAnalysis -Path $path

            if ($analysis -and $analysis.NeedTranscode.Count -gt 0) {
                $generateScript = Read-Host "`nGenerate FFmpeg transcode script? (Y/N) [N]"
                if ($generateScript -eq 'Y' -or $generateScript -eq 'y') {
                    New-TranscodeScript -TranscodeQueue $analysis.NeedTranscode
                }
            }
        }
    }
    "5" {
        # TMDB Metadata Fetch mode
        Write-Host "`n=== TMDB Metadata Fetch ===" -ForegroundColor Cyan
        Write-Host "This feature fetches movie metadata from The Movie Database (TMDB)" -ForegroundColor Gray
        Write-Host "and creates Kodi-compatible NFO files with full information." -ForegroundColor Gray
        Write-Host "`nGet a free API key at: https://www.themoviedb.org/settings/api" -ForegroundColor Yellow

        # Check if API key is in config
        if ($script:Config.TMDBApiKey) {
            Write-Host "API key found in configuration" -ForegroundColor Green
            $useExisting = Read-Host "Use saved API key? (Y/N) [Y]"
            if ($useExisting -eq 'N' -or $useExisting -eq 'n') {
                $apiKey = Read-Host "Enter new TMDB API key"
            } else {
                $apiKey = $script:Config.TMDBApiKey
            }
        } else {
            $apiKey = Read-Host "`nEnter your TMDB API key"
        }

        if ($apiKey) {
            $script:Config.TMDBApiKey = $apiKey

            $path = Select-FolderDialog -Description "Select your Movies folder"
            if ($path) {
                Invoke-TMDBMetadataFetch -Path $path -ApiKey $apiKey

                # Offer to save API key
                if (-not $script:Config.TMDBApiKey) {
                    $saveKey = Read-Host "Save API key to configuration? (Y/N) [Y]"
                    if ($saveKey -ne 'N' -and $saveKey -ne 'n') {
                        Export-Configuration
                    }
                }
            }
        } else {
            Write-Host "API key required for TMDB fetch" -ForegroundColor Red
        }
    }
    "6" {
        # Export Library Report
        Write-Host "`n=== Export Library Report ===" -ForegroundColor Cyan
        $mediaTypeInput = Read-Host "Media type? (1=Movies, 2=TV Shows) [1]"
        $mediaType = if ($mediaTypeInput -eq '2') { "TVShows" } else { "Movies" }

        $path = Select-FolderDialog -Description "Select your media library folder"
        if ($path) {
            Invoke-ExportMenu -Path $path -MediaType $mediaType
        }
    }
    "7" {
        # Enhanced Duplicate Scan
        Write-Host "`n=== Enhanced Duplicate Scan ===" -ForegroundColor Cyan
        Write-Host "This scan uses file hashing for accurate duplicate detection" -ForegroundColor Gray

        $path = Select-FolderDialog -Description "Select your Movies folder"
        if ($path) {
            Show-EnhancedDuplicateReport -Path $path
        }
    }
    "8" {
        # Undo Previous Session
        Write-Host "`n=== Undo Previous Session ===" -ForegroundColor Cyan

        # Find available undo manifests
        $manifests = Get-ChildItem -Path $script:UndoFolder -Filter "MediaCleaner_Undo_*.json" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending

        if ($manifests.Count -eq 0) {
            Write-Host "No undo manifests found" -ForegroundColor Yellow
        } else {
            Write-Host "Available undo manifests:" -ForegroundColor Cyan
            $i = 1
            foreach ($manifest in $manifests | Select-Object -First 10) {
                Write-Host "  $i. $($manifest.Name) ($(Get-Date $manifest.LastWriteTime -Format 'yyyy-MM-dd HH:mm'))" -ForegroundColor White
                $i++
            }

            $selection = Read-Host "`nSelect manifest number to undo (or Enter to cancel)"
            if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $manifests.Count) {
                $selectedManifest = $manifests[[int]$selection - 1]
                Invoke-UndoOperations -ManifestPath $selectedManifest.FullName
            } else {
                Write-Host "Undo cancelled" -ForegroundColor Yellow
            }
        }
    }
    "9" {
        # Configuration Management
        Write-Host "`n=== Configuration Management ===" -ForegroundColor Cyan
        Write-Host "1. Save current configuration"
        Write-Host "2. Load configuration from file"
        Write-Host "3. Reset to defaults"
        Write-Host "4. View current configuration"

        $configChoice = Read-Host "`nSelect option"
        switch ($configChoice) {
            "1" { Export-Configuration }
            "2" {
                $customPath = Read-Host "Enter config file path (or Enter for default)"
                if ($customPath) {
                    Import-Configuration -Path $customPath
                } else {
                    Import-Configuration
                }
            }
            "3" {
                $script:Config = $script:DefaultConfig.Clone()
                $script:Config.LogFile = Join-Path $PSScriptRoot "MediaCleaner_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                Write-Host "Configuration reset to defaults" -ForegroundColor Green
            }
            "4" {
                Write-Host "`nCurrent Configuration:" -ForegroundColor Cyan
                $script:Config.GetEnumerator() | Where-Object { $_.Key -ne 'Tags' } | ForEach-Object {
                    $value = if ($_.Value -is [array]) { $_.Value -join ', ' } else { $_.Value }
                    Write-Host "  $($_.Key): $value" -ForegroundColor White
                }
            }
        }
    }
    default {
        Write-Host "Invalid selection" -ForegroundColor Red
        Write-Log "Invalid type selection: $type" "ERROR"
    }
}

# Show statistics summary
Show-Statistics

Write-Host "`nLog file saved to: $($script:Config.LogFile)" -ForegroundColor Cyan
Write-Log "MediaCleaner session ended" "INFO"
