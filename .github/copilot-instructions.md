# MediaCleaner - AI Coding Agent Instructions

## Project Overview
MediaCleaner is a Windows PowerShell utility that automates media library cleanup for downloaded movies and TV shows. The single script (`MediaCleaner.ps1`) handles file organization, archive extraction, naming normalization, and metadata removal.

## Architecture & Core Patterns

### Single-Script Architecture
- **File**: `MediaCleaner.ps1` (184 lines)
- All functionality consolidates into one entry point with conditional branches for movies vs. shows
- Uses interactive user selection (FolderBrowserDialog) to determine target path and media type

### Two Execution Paths
1. **Movies mode** (`$type -eq 1`): Comprehensive cleanup with tag removal
2. **Shows mode** (`$type -eq 2`): Lightweight consolidation focusing on video extraction

### Key Dependencies
- **7-Zip** (`C:\Program Files\7-Zip\7z.exe`): Required for `.rar` extraction; auto-installed if missing
- **.NET Assembly**: `System.Windows.Forms` for folder browser dialog
- **Windows filesystem**: Assumes standard `NTFS` paths

## Critical Workflows

### Movie Cleanup Workflow
1. **Dependency check**: Verifies 7-Zip installation; downloads/installs if absent
2. **Delete unwanted content**: Removes folders matching patterns: `*Subs*`, `*Sample*`, `*Trailer*`, `*Proof*`, `*Screens*`
3. **Extract archives**: Processes all `.rar` files recursively, extracting to root path
4. **Organize loose files**: Moves uncontained files into new folders named after the file (minus extension)
5. **Normalize folder names**: Removes quality/format tags via regex splits (e.g., `1080p`, `BRRip`, `x264`)
6. **Format year suffixes**: Converts years to parenthetical format (e.g., `Movie 2020` → `Movie (2020)`)
7. **File extension cleanup**: Replaces dots in filenames with spaces

### Show Cleanup Workflow
- Simpler than movies: Extracts loose video files to root, unzips `.rar` archives, removes empty non-video folders
- Target extensions: `.mp4`, `.mkv`

## Project Conventions & Patterns

### Tag Removal Strategy
- **Pattern**: Split filenames on known quality/codec tags, keep everything before the tag
- **Implementation**: Chained `Get-ChildItem | Rename-Item` commands with `-split` operator
- **Tags handled**: 1080p, 2160p, 720p, HDRip, DVDRip, BRRip, BDRip, WEB-DL, x264, x265, hevc, etc. (38+ variants)
- **Example**: `Movie.Title.BRRip.2020.mkv` → `Movie.Title. (2020).mkv`

### Archive Handling
- Uses 7-Zip CLI via `sz x` (extract) with flags: `-o"$destpath"` (output), `-r` (recurse), `-y` (auto-confirm)
- Processes all `.rar` files via `Get-ChildItem -Filter *.rar -Recurse`
- Deletes extracted archives via wildcard: `*.r??` (matches `.rar`, `.r01`, etc.)

### Error Handling
- Minimal: Script exits on "No folder selected"; relies on Windows error messages for other failures
- Commented-out blocks document troubleshooting approaches (e.g., manual video file copying, empty folder cleanup)

## When Modifying This Codebase

### Adding Movie Cleanup Steps
- Insert new `Get-ChildItem` filter + rename/remove logic after line 64 (before year formatting)
- Use `Get-ChildItem -path $path -Filter *YourTag* | Rename-Item -NewName` pattern
- Test regex splits with sample filenames to avoid unintended truncation

### Extending Tag Removal
- Add lines before the year-formatting block (around line 138)
- Follow existing pattern: filter → `-split` on tag → rebuild name
- Use `[0]` index to keep pre-tag portion

### Show Mode Enhancements
- Modify within the `elseif ($type -eq 2)` block (lines 156-176)
- Maintain `$FileType` array for video extension filtering
- Avoid breaking the `break` statement at line 171 (prevents accidental fall-through)

### Testing Locally
- Create a test folder with mock media files/folders
- Run: `powershell -ExecutionPolicy Bypass -File MediaCleaner.ps1`
- Verify folder structure matches expected output before modifying cleanup logic

## Known Limitations & TODOs
- Show mode cleanup documented as "TBC" (to be completed) with commented Show Tools section
- No validation of successful extraction or file move operations
- No rollback capability if unexpected files are deleted
- Year regex (`19??`, `20??`) may need updates for future centuries
- Assumes single-level folder structure for movies (not deeply nested)

## File Modification Guidelines
- **MovieCleaner.ps1** is the only code file; all changes consolidate here
- **README.md**: Update requirements if dependencies change
- **LICENSE**: Preserve copyright attribution (created by Nick Kliatsko, last updated 12/22/2024)
