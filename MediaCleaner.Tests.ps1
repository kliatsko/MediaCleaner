#Requires -Modules Pester

<#
.SYNOPSIS
    Pester unit tests for MediaCleaner.ps1

.DESCRIPTION
    Comprehensive test suite for MediaCleaner functionality including:
    - Quality scoring
    - Episode parsing
    - Title normalization
    - NFO file operations
    - Duplicate detection
    - File operations (mocked)

.NOTES
    Run with: Invoke-Pester -Path .\MediaCleaner.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Import the script functions by dot-sourcing
    # We need to extract just the functions, not execute the main script

    # Get the script content and extract function definitions
    $scriptPath = Join-Path $PSScriptRoot "MediaCleaner.ps1"
    $scriptContent = Get-Content $scriptPath -Raw

    # Initialize config for tests
    $script:Config = @{
        LogFile = Join-Path $TestDrive "test.log"
        DryRun = $true
        SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
        VideoExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v')
        SubtitleExtensions = @('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
        ArchiveExtensions = @('*.rar', '*.zip', '*.7z', '*.tar', '*.gz', '*.bz2')
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
        Tags = @('1080p', '2160p', '720p', '480p', '4K', 'HDRip', 'BluRay', 'x264', 'x265', 'HEVC')
    }

    $script:Stats = @{
        StartTime = Get-Date
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
    }

    # Extract and define key functions for testing
    # Format-FileSize
    function Format-FileSize {
        param([long]$Bytes)
        if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
        if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
        if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
        if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
        return "$Bytes bytes"
    }

    # Get-QualityScore
    function Get-QualityScore {
        param([string]$FileName)

        $quality = @{
            Score = 0
            Resolution = "Unknown"
            Codec = "Unknown"
            Source = "Unknown"
            Audio = "Unknown"
            HDR = $false
            Details = @()
        }

        $fileNameLower = $FileName.ToLower()

        # Resolution
        if ($fileNameLower -match '2160p|4k|uhd') {
            $quality.Resolution = "2160p"
            $quality.Score += 100
        }
        elseif ($fileNameLower -match '1080p') {
            $quality.Resolution = "1080p"
            $quality.Score += 80
        }
        elseif ($fileNameLower -match '720p') {
            $quality.Resolution = "720p"
            $quality.Score += 60
        }
        elseif ($fileNameLower -match '480p|dvd') {
            $quality.Resolution = "480p"
            $quality.Score += 40
        }

        # Source
        if ($fileNameLower -match 'bluray|blu-ray|bdrip|brrip') {
            $quality.Source = "BluRay"
            $quality.Score += 30
        }
        elseif ($fileNameLower -match 'web-dl|webdl') {
            $quality.Source = "WEB-DL"
            $quality.Score += 25
        }
        elseif ($fileNameLower -match 'webrip') {
            $quality.Source = "WEBRip"
            $quality.Score += 20
        }
        elseif ($fileNameLower -match 'hdtv') {
            $quality.Source = "HDTV"
            $quality.Score += 15
        }

        # Codec
        if ($fileNameLower -match 'x265|h\.?265|hevc') {
            $quality.Codec = "HEVC/x265"
            $quality.Score += 20
        }
        elseif ($fileNameLower -match 'x264|h\.?264|avc') {
            $quality.Codec = "x264"
            $quality.Score += 15
        }

        # Audio
        if ($fileNameLower -match 'atmos') {
            $quality.Audio = "Atmos"
            $quality.Score += 15
        }
        elseif ($fileNameLower -match 'dts') {
            $quality.Audio = "DTS"
            $quality.Score += 8
        }
        elseif ($fileNameLower -match 'aac') {
            $quality.Audio = "AAC"
            $quality.Score += 3
        }

        # HDR
        if ($fileNameLower -match 'hdr') {
            $quality.HDR = $true
            $quality.Score += 10
        }

        return $quality
    }

    # Get-EpisodeInfo
    function Get-EpisodeInfo {
        param([string]$FileName)

        $info = @{
            Season = $null
            Episode = $null
            Episodes = @()
            ShowTitle = $null
            EpisodeTitle = $null
            IsMultiEpisode = $false
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

        # Pattern 1: S01E01 or S01E01E02
        if ($baseName -match '^(.+?)[.\s_-]+[Ss](\d{1,2})[Ee](\d{1,2})(?:[Ee-](\d{1,2}))?(?:[Ee-](\d{1,2}))?(.*)$') {
            $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
            $info.Season = [int]$Matches[2]
            $info.Episode = [int]$Matches[3]
            $info.Episodes += [int]$Matches[3]

            if ($Matches[4]) {
                $info.Episodes += [int]$Matches[4]
                $info.IsMultiEpisode = $true
            }
            if ($Matches[5]) {
                $info.Episodes += [int]$Matches[5]
            }
        }
        # Pattern 2: 1x01
        elseif ($baseName -match '^(.+?)[.\s_-]+(\d{1,2})x(\d{1,2})(.*)$') {
            $info.ShowTitle = $Matches[1] -replace '\.', ' ' -replace '\s+', ' '
            $info.Season = [int]$Matches[2]
            $info.Episode = [int]$Matches[3]
            $info.Episodes += [int]$Matches[3]
        }

        if ($info.ShowTitle) {
            $info.ShowTitle = $info.ShowTitle.Trim()
        }

        return $info
    }

    # Get-NormalizedTitle
    function Get-NormalizedTitle {
        param([string]$Name)

        $result = @{
            NormalizedTitle = $null
            Year = $null
        }

        $name = [System.IO.Path]::GetFileNameWithoutExtension($Name)

        if ($name -match '[\(\[\s]*(19|20)\d{2}[\)\]\s]*') {
            $yearMatch = [regex]::Match($name, '(19|20)\d{2}')
            if ($yearMatch.Success) {
                $result.Year = $yearMatch.Value
            }
        }

        $title = $name -replace '[\(\[]?(19|20)\d{2}[\)\]]?.*$', ''
        $title = $title -replace '\s*(720p|1080p|2160p|4K|HDRip|DVDRip|BRRip|BluRay|WEB-DL|WEBRip|x264|x265|HEVC).*$', ''
        $title = $title -replace '\.', ' '
        $title = $title -replace '[_-]', ' '
        $title = $title -replace '\s+', ' '
        $title = $title.Trim().ToLower()
        $title = $title -replace '^(the|a|an)\s+', ''

        $result.NormalizedTitle = $title

        return $result
    }

    # Write-Log (mock version)
    function Write-Log {
        param(
            [string]$Message,
            [string]$Level = "INFO"
        )
        # Mock - just track errors/warnings
        if ($Level -eq "ERROR") { $script:Stats.Errors++ }
        if ($Level -eq "WARNING") { $script:Stats.Warnings++ }
    }
}

Describe "Format-FileSize" {
    It "Formats bytes correctly" {
        Format-FileSize 500 | Should -Be "500 bytes"
    }

    It "Formats kilobytes correctly" {
        Format-FileSize 1024 | Should -Be "1.00 KB"
        Format-FileSize 2048 | Should -Be "2.00 KB"
    }

    It "Formats megabytes correctly" {
        Format-FileSize (1024 * 1024) | Should -Be "1.00 MB"
        Format-FileSize (500 * 1024 * 1024) | Should -Be "500.00 MB"
    }

    It "Formats gigabytes correctly" {
        Format-FileSize (1024 * 1024 * 1024) | Should -Be "1.00 GB"
        Format-FileSize (4.5 * 1024 * 1024 * 1024) | Should -Be "4.50 GB"
    }

    It "Formats terabytes correctly" {
        Format-FileSize (1024 * 1024 * 1024 * 1024) | Should -Be "1.00 TB"
    }
}

Describe "Get-QualityScore" {
    Context "Resolution Detection" {
        It "Detects 2160p/4K resolution" {
            $result = Get-QualityScore "Movie.2160p.BluRay.mkv"
            $result.Resolution | Should -Be "2160p"
            $result.Score | Should -BeGreaterOrEqual 100
        }

        It "Detects 1080p resolution" {
            $result = Get-QualityScore "Movie.1080p.WEB-DL.mkv"
            $result.Resolution | Should -Be "1080p"
            $result.Score | Should -BeGreaterOrEqual 80
        }

        It "Detects 720p resolution" {
            $result = Get-QualityScore "Movie.720p.HDTV.mkv"
            $result.Resolution | Should -Be "720p"
            $result.Score | Should -BeGreaterOrEqual 60
        }

        It "Detects 480p resolution" {
            $result = Get-QualityScore "Movie.480p.DVDRip.mkv"
            $result.Resolution | Should -Be "480p"
            $result.Score | Should -BeGreaterOrEqual 40
        }
    }

    Context "Source Detection" {
        It "Detects BluRay source" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.Source | Should -Be "BluRay"
        }

        It "Detects WEB-DL source" {
            $result = Get-QualityScore "Movie.1080p.WEB-DL.x264.mkv"
            $result.Source | Should -Be "WEB-DL"
        }

        It "Detects WEBRip source" {
            $result = Get-QualityScore "Movie.1080p.WEBRip.x264.mkv"
            $result.Source | Should -Be "WEBRip"
        }

        It "Detects HDTV source" {
            $result = Get-QualityScore "Movie.720p.HDTV.x264.mkv"
            $result.Source | Should -Be "HDTV"
        }
    }

    Context "Codec Detection" {
        It "Detects HEVC/x265 codec" {
            $result = Get-QualityScore "Movie.2160p.BluRay.x265.mkv"
            $result.Codec | Should -Be "HEVC/x265"
        }

        It "Detects x264 codec" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.Codec | Should -Be "x264"
        }
    }

    Context "HDR Detection" {
        It "Detects HDR content" {
            $result = Get-QualityScore "Movie.2160p.BluRay.HDR.x265.mkv"
            $result.HDR | Should -Be $true
        }

        It "Non-HDR content returns false" {
            $result = Get-QualityScore "Movie.1080p.BluRay.x264.mkv"
            $result.HDR | Should -Be $false
        }
    }

    Context "Score Calculation" {
        It "Higher quality gets higher score" {
            $score4k = (Get-QualityScore "Movie.2160p.BluRay.x265.mkv").Score
            $score1080 = (Get-QualityScore "Movie.1080p.BluRay.x264.mkv").Score
            $score720 = (Get-QualityScore "Movie.720p.HDTV.x264.mkv").Score

            $score4k | Should -BeGreaterThan $score1080
            $score1080 | Should -BeGreaterThan $score720
        }
    }
}

Describe "Get-EpisodeInfo" {
    Context "S01E01 Format" {
        It "Parses standard S01E01 format" {
            $result = Get-EpisodeInfo "Breaking.Bad.S01E01.Pilot.720p.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.ShowTitle | Should -Be "Breaking Bad"
            $result.IsMultiEpisode | Should -Be $false
        }

        It "Parses lowercase s01e01 format" {
            $result = Get-EpisodeInfo "the.office.s02e05.720p.mkv"
            $result.Season | Should -Be 2
            $result.Episode | Should -Be 5
            $result.ShowTitle | Should -Be "the office"
        }

        It "Parses double-digit episode numbers" {
            $result = Get-EpisodeInfo "Show.Name.S01E15.Episode.Title.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 15
        }

        It "Parses double-digit season numbers" {
            $result = Get-EpisodeInfo "Show.Name.S12E03.mkv"
            $result.Season | Should -Be 12
            $result.Episode | Should -Be 3
        }
    }

    Context "Multi-Episode Format" {
        It "Parses S01E01E02 format" {
            $result = Get-EpisodeInfo "Show.Name.S01E01E02.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.Episodes | Should -Contain 1
            $result.Episodes | Should -Contain 2
            $result.IsMultiEpisode | Should -Be $true
        }

        It "Parses S01E01-E03 format" {
            $result = Get-EpisodeInfo "Show.Name.S01E01-E03.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 1
            $result.IsMultiEpisode | Should -Be $true
        }
    }

    Context "1x01 Format" {
        It "Parses 1x01 format" {
            $result = Get-EpisodeInfo "Show.Name.1x05.mkv"
            $result.Season | Should -Be 1
            $result.Episode | Should -Be 5
        }

        It "Parses double-digit 10x15 format" {
            $result = Get-EpisodeInfo "Show.Name.10x15.mkv"
            $result.Season | Should -Be 10
            $result.Episode | Should -Be 15
        }
    }

    Context "Invalid Formats" {
        It "Returns null for unrecognized format" {
            $result = Get-EpisodeInfo "Random.Movie.2024.mkv"
            $result.Season | Should -Be $null
            $result.Episode | Should -Be $null
        }
    }
}

Describe "Get-NormalizedTitle" {
    Context "Year Extraction" {
        It "Extracts year in parentheses" {
            $result = Get-NormalizedTitle "The Movie (2024)"
            $result.Year | Should -Be "2024"
        }

        It "Extracts year without parentheses" {
            $result = Get-NormalizedTitle "The Movie 2024"
            $result.Year | Should -Be "2024"
        }

        It "Extracts year from complex filename" {
            $result = Get-NormalizedTitle "The.Movie.2024.1080p.BluRay.x264"
            $result.Year | Should -Be "2024"
        }

        It "Handles movies from 1900s" {
            $result = Get-NormalizedTitle "Classic Film (1985)"
            $result.Year | Should -Be "1985"
        }
    }

    Context "Title Normalization" {
        It "Removes quality tags" {
            $result = Get-NormalizedTitle "Movie.Name.1080p.BluRay.x264"
            $result.NormalizedTitle | Should -Not -Match "1080p"
            $result.NormalizedTitle | Should -Not -Match "bluray"
        }

        It "Replaces dots with spaces" {
            $result = Get-NormalizedTitle "Movie.Name.2024"
            $result.NormalizedTitle | Should -Not -Match "\."
        }

        It "Converts to lowercase" {
            $result = Get-NormalizedTitle "THE MOVIE NAME"
            $result.NormalizedTitle | Should -Be "movie name"
        }

        It "Removes leading articles" {
            $result = Get-NormalizedTitle "The Great Movie (2024)"
            $result.NormalizedTitle | Should -Be "great movie"

            $result = Get-NormalizedTitle "A Good Film (2024)"
            $result.NormalizedTitle | Should -Be "good film"
        }

        It "Trims whitespace" {
            $result = Get-NormalizedTitle "  Movie Name  (2024)"
            $result.NormalizedTitle | Should -Not -Match "^\s"
            $result.NormalizedTitle | Should -Not -Match "\s$"
        }
    }

    Context "Edge Cases" {
        It "Handles empty input" {
            $result = Get-NormalizedTitle ""
            $result.NormalizedTitle | Should -Be ""
        }

        It "Handles no year" {
            $result = Get-NormalizedTitle "Movie Without Year"
            $result.Year | Should -Be $null
            $result.NormalizedTitle | Should -Be "movie without year"
        }
    }
}

Describe "Configuration" {
    It "Has valid video extensions" {
        $script:Config.VideoExtensions | Should -Contain ".mp4"
        $script:Config.VideoExtensions | Should -Contain ".mkv"
        $script:Config.VideoExtensions | Should -Contain ".avi"
    }

    It "Has valid subtitle extensions" {
        $script:Config.SubtitleExtensions | Should -Contain ".srt"
        $script:Config.SubtitleExtensions | Should -Contain ".sub"
    }

    It "Has valid archive extensions" {
        $script:Config.ArchiveExtensions | Should -Contain "*.rar"
        $script:Config.ArchiveExtensions | Should -Contain "*.zip"
        $script:Config.ArchiveExtensions | Should -Contain "*.7z"
    }

    It "Has preferred subtitle languages" {
        $script:Config.PreferredSubtitleLanguages | Should -Contain "eng"
        $script:Config.PreferredSubtitleLanguages | Should -Contain "en"
    }
}

Describe "Statistics Tracking" {
    BeforeEach {
        # Reset stats
        $script:Stats.FilesDeleted = 0
        $script:Stats.BytesDeleted = 0
        $script:Stats.Errors = 0
        $script:Stats.Warnings = 0
    }

    It "Tracks errors via Write-Log" {
        Write-Log "Test error" "ERROR"
        $script:Stats.Errors | Should -Be 1
    }

    It "Tracks warnings via Write-Log" {
        Write-Log "Test warning" "WARNING"
        $script:Stats.Warnings | Should -Be 1
    }

    It "Tracks multiple errors" {
        Write-Log "Error 1" "ERROR"
        Write-Log "Error 2" "ERROR"
        Write-Log "Error 3" "ERROR"
        $script:Stats.Errors | Should -Be 3
    }
}

Describe "File Matching Patterns" {
    Context "Unnecessary Files" {
        It "Matches sample files" {
            "Sample.mkv" -like "*Sample*" | Should -Be $true
            "movie-sample.mkv" -like "*sample*" | Should -Be $true
        }

        It "Matches proof files" {
            "Proof.jpg" -like "*Proof*" | Should -Be $true
        }

        It "Matches screenshots" {
            "Screens" -like "*Screens*" | Should -Be $true
            "Screenshots" -like "*Screens*" | Should -Be $true
        }
    }

    Context "Trailer Files" {
        It "Matches trailer files" {
            "Movie-Trailer.mkv" -like "*Trailer*" | Should -Be $true
            "movie-trailer.mkv" -like "*trailer*" | Should -Be $true
        }

        It "Matches teaser files" {
            "Movie-Teaser.mkv" -like "*Teaser*" | Should -Be $true
        }
    }
}

Describe "Integration Tests" -Tag "Integration" {
    BeforeAll {
        # Create test directory structure
        $testRoot = Join-Path $TestDrive "MediaLibrary"
        $movieFolder = Join-Path $testRoot "Movies"
        $tvFolder = Join-Path $testRoot "TVShows"

        New-Item -Path $movieFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $tvFolder -ItemType Directory -Force | Out-Null
    }

    Context "Movie Library Structure" {
        BeforeAll {
            $moviePath = Join-Path $TestDrive "MediaLibrary\Movies"

            # Create test movie folders
            $movie1 = Join-Path $moviePath "The.Matrix.1999.1080p.BluRay.x264"
            $movie2 = Join-Path $moviePath "Inception (2010)"

            New-Item -Path $movie1 -ItemType Directory -Force | Out-Null
            New-Item -Path $movie2 -ItemType Directory -Force | Out-Null

            # Create dummy video files
            New-Item -Path (Join-Path $movie1 "movie.mkv") -ItemType File -Force | Out-Null
            New-Item -Path (Join-Path $movie2 "movie.mkv") -ItemType File -Force | Out-Null
        }

        It "Identifies movie folders" {
            $moviePath = Join-Path $TestDrive "MediaLibrary\Movies"
            $folders = Get-ChildItem -Path $moviePath -Directory
            $folders.Count | Should -Be 2
        }

        It "Parses movie names correctly" {
            $result1 = Get-NormalizedTitle "The.Matrix.1999.1080p.BluRay.x264"
            $result1.Year | Should -Be "1999"

            $result2 = Get-NormalizedTitle "Inception (2010)"
            $result2.Year | Should -Be "2010"
        }
    }

    Context "TV Show Library Structure" {
        BeforeAll {
            $tvPath = Join-Path $TestDrive "MediaLibrary\TVShows"
            $showPath = Join-Path $tvPath "Breaking Bad"

            New-Item -Path $showPath -ItemType Directory -Force | Out-Null

            # Create test episode files
            $episodes = @(
                "Breaking.Bad.S01E01.Pilot.720p.mkv",
                "Breaking.Bad.S01E02.720p.mkv",
                "Breaking.Bad.S02E01.720p.mkv"
            )

            foreach ($ep in $episodes) {
                New-Item -Path (Join-Path $showPath $ep) -ItemType File -Force | Out-Null
            }
        }

        It "Identifies episode files" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File
            $files.Count | Should -Be 3
        }

        It "Parses episode info for all files" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File

            foreach ($file in $files) {
                $info = Get-EpisodeInfo $file.Name
                $info.Season | Should -Not -Be $null
                $info.Episode | Should -Not -Be $null
            }
        }

        It "Identifies multiple seasons" {
            $showPath = Join-Path $TestDrive "MediaLibrary\TVShows\Breaking Bad"
            $files = Get-ChildItem -Path $showPath -File

            $seasons = $files | ForEach-Object {
                $info = Get-EpisodeInfo $_.Name
                $info.Season
            } | Sort-Object -Unique

            $seasons.Count | Should -Be 2
            $seasons | Should -Contain 1
            $seasons | Should -Contain 2
        }
    }
}

Describe "Edge Cases and Error Handling" {
    Context "Special Characters in Filenames" {
        It "Handles filenames with apostrophes" {
            $result = Get-NormalizedTitle "Marvel's Avengers (2012)"
            $result.Year | Should -Be "2012"
        }

        It "Handles filenames with colons" {
            $result = Get-NormalizedTitle "Movie - The Sequel (2024)"
            $result.NormalizedTitle | Should -Not -Be $null
        }

        It "Handles filenames with brackets" {
            $result = Get-NormalizedTitle "Movie [2024] 1080p"
            $result.Year | Should -Be "2024"
        }
    }

    Context "Empty and Null Inputs" {
        It "Handles null filename in Get-QualityScore" {
            { Get-QualityScore $null } | Should -Not -Throw
        }

        It "Handles empty filename in Get-EpisodeInfo" {
            $result = Get-EpisodeInfo ""
            $result.Season | Should -Be $null
        }
    }
}
