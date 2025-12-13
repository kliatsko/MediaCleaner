#!/usr/bin/env python3
"""
MediaCleaner - Cross-platform media library organization and cleanup tool

This Python script automates the cleanup and organization of downloaded media files.
It's a cross-platform port of the PowerShell MediaCleaner script.

Features:
    - Archive extraction (rar, zip, 7z, tar, gz, bz2)
    - Unnecessary file removal (samples, proofs, screenshots)
    - Subtitle handling (keep preferred language)
    - Trailer management (move to _Trailers folder)
    - Folder organization and name cleaning
    - TV show episode detection and organization
    - Duplicate detection with quality scoring
    - NFO file parsing and generation
    - TMDB metadata integration
    - Health check and codec analysis

Requirements:
    - Python 3.8+
    - 7-Zip or p7zip installed
    - Optional: requests library for TMDB integration

Author: Nick Kliatsko
Version: 5.0
"""

import os
import re
import sys
import json
import shutil
import logging
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any, Tuple
from xml.etree import ElementTree as ET
from xml.dom import minidom

# Optional imports
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


# =============================================================================
# CONFIGURATION
# =============================================================================

@dataclass
class Config:
    """Global configuration settings"""
    dry_run: bool = False
    log_file: str = ""
    seven_zip_path: str = ""

    video_extensions: tuple = ('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.m4v')
    subtitle_extensions: tuple = ('.srt', '.sub', '.idx', '.ass', '.ssa', '.vtt')
    archive_extensions: tuple = ('.rar', '.zip', '.7z', '.tar', '.gz', '.bz2')

    unnecessary_patterns: tuple = ('*Sample*', '*Proof*', '*Screens*')
    trailer_patterns: tuple = ('*Trailer*', '*trailer*', '*TRAILER*', '*Teaser*', '*teaser*')
    preferred_subtitle_languages: tuple = ('eng', 'en', 'english')

    keep_subtitles: bool = True
    keep_trailers: bool = True
    generate_nfo: bool = False
    organize_seasons: bool = True
    rename_episodes: bool = False
    check_duplicates: bool = False
    tmdb_api_key: str = ""

    # Quality/release tags to remove from folder names
    tags: tuple = (
        '1080p', '2160p', '720p', '480p', '4K', '2K', 'UHD',
        'HDRip', 'DVDRip', 'BRRip', 'BR-Rip', 'BDRip', 'BD-Rip', 'WEB-DL', 'WEBRip', 'BluRay',
        'x264', 'x265', 'X265', 'H264', 'H265', 'HEVC', 'XviD', 'DivX', 'AVC',
        'AAC', 'AC3', 'DTS', 'Atmos', 'TrueHD', 'DD5.1', '5.1', '7.1',
        'HDR', 'HDR10', 'HDR10+', 'DolbyVision', 'SDR', '10bit',
        'Extended', 'Unrated', 'Remastered', 'REPACK', 'PROPER',
        'YIFY', 'YTS', 'RARBG', 'SPARKS', 'NF', 'AMZN', 'HULU', 'WEB'
    )

    def __post_init__(self):
        # Set up log file
        if not self.log_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.log_file = f"MediaCleaner_{timestamp}.log"

        # Find 7-zip
        if not self.seven_zip_path:
            self.seven_zip_path = self._find_seven_zip()

    def _find_seven_zip(self) -> str:
        """Find 7-zip executable path"""
        if sys.platform == 'win32':
            paths = [
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe"
            ]
        else:
            paths = ['/usr/bin/7z', '/usr/local/bin/7z', '/opt/homebrew/bin/7z']

        for path in paths:
            if os.path.exists(path):
                return path

        # Try to find in PATH
        try:
            result = subprocess.run(['which', '7z'] if sys.platform != 'win32' else ['where', '7z'],
                                  capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip().split('\n')[0]
        except:
            pass

        return ""


@dataclass
class Stats:
    """Statistics tracking"""
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None
    files_deleted: int = 0
    bytes_deleted: int = 0
    archives_extracted: int = 0
    archives_failed: int = 0
    folders_created: int = 0
    folders_renamed: int = 0
    files_moved: int = 0
    empty_folders_removed: int = 0
    subtitles_processed: int = 0
    subtitles_deleted: int = 0
    trailers_moved: int = 0
    nfo_files_created: int = 0
    nfo_files_read: int = 0
    errors: int = 0
    warnings: int = 0


# Global instances
config = Config()
stats = Stats()
logger: Optional[logging.Logger] = None


# =============================================================================
# LOGGING
# =============================================================================

def setup_logging(log_file: str) -> logging.Logger:
    """Set up logging to file and console"""
    log = logging.getLogger('MediaCleaner')
    log.setLevel(logging.DEBUG)

    # File handler
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    log.addHandler(fh)

    return log


def log_info(message: str):
    """Log info message"""
    if logger:
        logger.info(message)


def log_warning(message: str):
    """Log warning message"""
    if logger:
        logger.warning(message)
    stats.warnings += 1


def log_error(message: str):
    """Log error message"""
    if logger:
        logger.error(message)
    stats.errors += 1


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def format_file_size(bytes_size: int) -> str:
    """Format bytes to human-readable string"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes_size < 1024:
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024
    return f"{bytes_size:.2f} PB"


def print_color(message: str, color: str = 'white', end: str = '\n'):
    """Print colored output (cross-platform)"""
    colors = {
        'red': '\033[91m',
        'green': '\033[92m',
        'yellow': '\033[93m',
        'blue': '\033[94m',
        'magenta': '\033[95m',
        'cyan': '\033[96m',
        'white': '\033[97m',
        'gray': '\033[90m',
        'reset': '\033[0m'
    }

    # Enable ANSI on Windows
    if sys.platform == 'win32':
        os.system('')

    print(f"{colors.get(color, '')}{message}{colors['reset']}", end=end)


def matches_pattern(name: str, patterns: tuple) -> bool:
    """Check if name matches any pattern (supports wildcards)"""
    import fnmatch
    for pattern in patterns:
        if fnmatch.fnmatch(name, pattern) or fnmatch.fnmatch(name.lower(), pattern.lower()):
            return True
    return False


# =============================================================================
# QUALITY SCORING
# =============================================================================

def get_quality_score(filename: str) -> Dict[str, Any]:
    """Calculate quality score based on filename"""
    quality = {
        'score': 0,
        'resolution': 'Unknown',
        'codec': 'Unknown',
        'source': 'Unknown',
        'audio': 'Unknown',
        'hdr': False,
        'details': []
    }

    filename_lower = filename.lower()

    # Resolution scoring
    if any(x in filename_lower for x in ['2160p', '4k', 'uhd']):
        quality['resolution'] = '2160p'
        quality['score'] += 100
        quality['details'].append('4K/2160p (+100)')
    elif '1080p' in filename_lower:
        quality['resolution'] = '1080p'
        quality['score'] += 80
        quality['details'].append('1080p (+80)')
    elif '720p' in filename_lower:
        quality['resolution'] = '720p'
        quality['score'] += 60
        quality['details'].append('720p (+60)')
    elif any(x in filename_lower for x in ['480p', 'dvd']):
        quality['resolution'] = '480p'
        quality['score'] += 40
        quality['details'].append('480p (+40)')

    # Source scoring
    if any(x in filename_lower for x in ['bluray', 'blu-ray', 'bdrip', 'brrip']):
        quality['source'] = 'BluRay'
        quality['score'] += 30
        quality['details'].append('BluRay (+30)')
    elif any(x in filename_lower for x in ['web-dl', 'webdl']):
        quality['source'] = 'WEB-DL'
        quality['score'] += 25
        quality['details'].append('WEB-DL (+25)')
    elif 'webrip' in filename_lower:
        quality['source'] = 'WEBRip'
        quality['score'] += 20
        quality['details'].append('WEBRip (+20)')
    elif 'hdtv' in filename_lower:
        quality['source'] = 'HDTV'
        quality['score'] += 15
        quality['details'].append('HDTV (+15)')
    elif 'dvdrip' in filename_lower:
        quality['source'] = 'DVDRip'
        quality['score'] += 10
        quality['details'].append('DVDRip (+10)')

    # Codec scoring
    if any(x in filename_lower for x in ['x265', 'h265', 'h.265', 'hevc']):
        quality['codec'] = 'HEVC/x265'
        quality['score'] += 20
        quality['details'].append('HEVC/x265 (+20)')
    elif any(x in filename_lower for x in ['x264', 'h264', 'h.264', 'avc']):
        quality['codec'] = 'x264'
        quality['score'] += 15
        quality['details'].append('x264 (+15)')
    elif any(x in filename_lower for x in ['xvid', 'divx']):
        quality['codec'] = 'XviD'
        quality['score'] += 5
        quality['details'].append('XviD (+5)')

    # Audio scoring
    if 'atmos' in filename_lower:
        quality['audio'] = 'Atmos'
        quality['score'] += 15
        quality['details'].append('Atmos (+15)')
    elif 'truehd' in filename_lower:
        quality['audio'] = 'TrueHD'
        quality['score'] += 12
        quality['details'].append('TrueHD (+12)')
    elif any(x in filename_lower for x in ['dts-hd', 'dtshd']):
        quality['audio'] = 'DTS-HD'
        quality['score'] += 10
        quality['details'].append('DTS-HD (+10)')
    elif 'dts' in filename_lower:
        quality['audio'] = 'DTS'
        quality['score'] += 8
        quality['details'].append('DTS (+8)')
    elif any(x in filename_lower for x in ['ac3', 'dd5.1', 'dd5 1']):
        quality['audio'] = 'AC3'
        quality['score'] += 5
        quality['details'].append('AC3 (+5)')
    elif 'aac' in filename_lower:
        quality['audio'] = 'AAC'
        quality['score'] += 3
        quality['details'].append('AAC (+3)')

    # HDR scoring
    if any(x in filename_lower for x in ['hdr10+', 'hdr10plus']):
        quality['hdr'] = True
        quality['score'] += 15
        quality['details'].append('HDR10+ (+15)')
    elif any(x in filename_lower for x in ['dolby vision', 'dolbyvision', 'dovi', ' dv ']):
        quality['hdr'] = True
        quality['score'] += 15
        quality['details'].append('Dolby Vision (+15)')
    elif 'hdr10' in filename_lower:
        quality['hdr'] = True
        quality['score'] += 12
        quality['details'].append('HDR10 (+12)')
    elif 'hdr' in filename_lower:
        quality['hdr'] = True
        quality['score'] += 10
        quality['details'].append('HDR (+10)')

    return quality


# =============================================================================
# EPISODE PARSING
# =============================================================================

def get_episode_info(filename: str) -> Dict[str, Any]:
    """Parse season/episode info from filename"""
    info = {
        'season': None,
        'episode': None,
        'episodes': [],
        'show_title': None,
        'episode_title': None,
        'is_multi_episode': False
    }

    basename = Path(filename).stem

    # Pattern 1: S01E01 or S01E01E02 or S01E01-E03
    match = re.match(r'^(.+?)[.\s_-]+[Ss](\d{1,2})[Ee](\d{1,2})(?:[Ee-](\d{1,2}))?(?:[Ee-](\d{1,2}))?(.*)$', basename)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])

        if match.group(4):
            info['episodes'].append(int(match.group(4)))
            info['is_multi_episode'] = True
        if match.group(5):
            info['episodes'].append(int(match.group(5)))

        if match.group(6):
            remainder = re.sub(r'^[.\s_-]+', '', match.group(6))
            remainder = re.sub(r'\.', ' ', remainder)
            remainder = re.sub(r'\s*(720p|1080p|2160p|4K|HDTV|WEB-DL|WEBRip|BluRay|x264|x265|HEVC|AAC|AC3).*$', '', remainder, flags=re.IGNORECASE)
            if remainder.strip():
                info['episode_title'] = remainder.strip()

        return info

    # Pattern 2: 1x01 format
    match = re.match(r'^(.+?)[.\s_-]+(\d{1,2})x(\d{1,2})(.*)$', basename)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])
        return info

    # Pattern 3: Season 1 Episode 1
    match = re.match(r'^(.+?)[.\s_-]+Season\s*(\d{1,2})[.\s_-]+Episode\s*(\d{1,2})(.*)$', basename, re.IGNORECASE)
    if match:
        info['show_title'] = re.sub(r'\.', ' ', match.group(1)).strip()
        info['season'] = int(match.group(2))
        info['episode'] = int(match.group(3))
        info['episodes'].append(info['episode'])
        return info

    return info


def get_normalized_title(name: str) -> Dict[str, Any]:
    """Extract normalized title and year from folder/file name"""
    result = {'normalized_title': None, 'year': None}

    basename = Path(name).stem

    # Try to extract year
    year_match = re.search(r'(19|20)\d{2}', basename)
    if year_match:
        result['year'] = year_match.group()

    # Remove everything after year or quality tags
    title = re.sub(r'[\(\[]?(19|20)\d{2}[\)\]]?.*$', '', basename)
    title = re.sub(r'\s*(720p|1080p|2160p|4K|HDRip|DVDRip|BRRip|BluRay|WEB-DL|WEBRip|x264|x265|HEVC).*$', '', title, flags=re.IGNORECASE)

    # Normalize
    title = re.sub(r'\.', ' ', title)
    title = re.sub(r'[_-]', ' ', title)
    title = re.sub(r'\s+', ' ', title)
    title = title.strip().lower()

    # Remove common articles
    title = re.sub(r'^(the|a|an)\s+', '', title)

    result['normalized_title'] = title
    return result


# =============================================================================
# FILE OPERATIONS
# =============================================================================

def remove_unnecessary_files(path: str):
    """Remove sample, proof, and screenshot files"""
    print_color("Cleaning unnecessary files...", 'yellow')
    log_info(f"Starting unnecessary file cleanup in: {path}")

    try:
        for root, dirs, files in os.walk(path):
            # Check directories
            for d in dirs[:]:
                if matches_pattern(d, config.unnecessary_patterns):
                    dir_path = os.path.join(root, d)
                    if config.dry_run:
                        print_color(f"[DRY-RUN] Would delete: {dir_path}", 'yellow')
                        log_info(f"Would delete: {dir_path}")
                    else:
                        try:
                            dir_size = sum(f.stat().st_size for f in Path(dir_path).rglob('*') if f.is_file())
                            shutil.rmtree(dir_path)
                            print_color(f"Deleted: {d}", 'gray')
                            log_info(f"Deleted: {dir_path}")
                            stats.files_deleted += 1
                            stats.bytes_deleted += dir_size
                            dirs.remove(d)
                        except Exception as e:
                            log_error(f"Error deleting {dir_path}: {e}")

            # Check files
            for f in files:
                if matches_pattern(f, config.unnecessary_patterns):
                    file_path = os.path.join(root, f)
                    if config.dry_run:
                        print_color(f"[DRY-RUN] Would delete: {file_path}", 'yellow')
                    else:
                        try:
                            file_size = os.path.getsize(file_path)
                            os.remove(file_path)
                            print_color(f"Deleted: {f}", 'gray')
                            stats.files_deleted += 1
                            stats.bytes_deleted += file_size
                        except Exception as e:
                            log_error(f"Error deleting {file_path}: {e}")

        print_color("Unnecessary files cleaned", 'green')
    except Exception as e:
        log_error(f"Error during file cleanup: {e}")


def process_subtitles(path: str):
    """Process subtitle files - keep preferred languages"""
    print_color("Processing subtitle files...", 'yellow')
    log_info(f"Starting subtitle processing in: {path}")

    try:
        subtitle_files = []
        for ext in config.subtitle_extensions:
            subtitle_files.extend(Path(path).rglob(f'*{ext}'))

        if not subtitle_files:
            print_color("No subtitle files found", 'cyan')
            return

        print_color(f"Found {len(subtitle_files)} subtitle file(s)", 'cyan')

        for sub in subtitle_files:
            stats.subtitles_processed += 1
            sub_name_lower = sub.stem.lower()

            # Check if preferred language
            is_preferred = False
            for lang in config.preferred_subtitle_languages:
                if re.search(rf'\.{lang}$|\.{lang}\.|_{lang}$|_{lang}[_.]', sub_name_lower):
                    is_preferred = True
                    break

            # If no language tag, assume default language
            has_lang_tag = bool(re.search(r'\.(eng|en|english|spa|es|spanish|fre|fr|french|ger|de|german)(\.|$|_)', sub_name_lower))
            if not has_lang_tag:
                is_preferred = True

            if config.keep_subtitles and is_preferred:
                print_color(f"Keeping subtitle: {sub.name}", 'green')
            else:
                if config.dry_run:
                    print_color(f"[DRY-RUN] Would delete subtitle: {sub.name}", 'yellow')
                else:
                    try:
                        sub_size = sub.stat().st_size
                        sub.unlink()
                        print_color(f"Deleted subtitle: {sub.name}", 'gray')
                        stats.subtitles_deleted += 1
                        stats.bytes_deleted += sub_size
                    except Exception as e:
                        log_error(f"Error deleting subtitle {sub}: {e}")

        print_color("Subtitle processing completed", 'green')
    except Exception as e:
        log_error(f"Error processing subtitles: {e}")


def move_trailers_to_folder(path: str):
    """Move trailer files to _Trailers folder"""
    print_color("Processing trailer files...", 'yellow')
    log_info(f"Starting trailer processing in: {path}")

    try:
        trailer_files = []

        for pattern in config.trailer_patterns:
            for ext in config.video_extensions:
                # Convert glob pattern
                search_pattern = pattern.replace('*', '**/*') + ext
                trailer_files.extend(Path(path).rglob(f'*{pattern.strip("*")}*{ext}'))

        # Remove duplicates
        trailer_files = list(set(trailer_files))

        if not trailer_files:
            print_color("No trailer files found", 'cyan')
            return

        print_color(f"Found {len(trailer_files)} trailer file(s)", 'cyan')

        trailers_folder = Path(path) / '_Trailers'

        if not config.dry_run and not trailers_folder.exists():
            trailers_folder.mkdir(parents=True)
            print_color("Created _Trailers folder", 'green')
            stats.folders_created += 1

        for trailer in trailer_files:
            dest_path = trailers_folder / trailer.name

            # Handle duplicates
            counter = 1
            while dest_path.exists():
                dest_path = trailers_folder / f"{trailer.stem}_{counter}{trailer.suffix}"
                counter += 1

            if config.dry_run:
                print_color(f"[DRY-RUN] Would move trailer: {trailer.name} -> _Trailers/", 'yellow')
            else:
                try:
                    shutil.move(str(trailer), str(dest_path))
                    print_color(f"Moved trailer: {trailer.name}", 'green')
                    stats.trailers_moved += 1
                except Exception as e:
                    log_warning(f"Error moving trailer {trailer.name}: {e}")

        print_color("Trailer processing completed", 'green')
    except Exception as e:
        log_error(f"Error processing trailers: {e}")


def extract_archives(path: str, delete_after: bool = True):
    """Extract all archives in the path"""
    print_color("Extracting archives...", 'yellow')
    log_info(f"Starting archive extraction in: {path}")

    if not config.seven_zip_path:
        print_color("7-Zip not found. Please install 7-Zip.", 'red')
        log_error("7-Zip not found")
        return

    try:
        archive_files = []
        for ext in config.archive_extensions:
            archive_files.extend(Path(path).rglob(f'*{ext}'))

        if not archive_files:
            print_color("No archives found to extract", 'cyan')
            return

        print_color(f"Found {len(archive_files)} archive(s) to extract", 'cyan')

        for i, archive in enumerate(archive_files, 1):
            if config.dry_run:
                print_color(f"[DRY-RUN] [{i}/{len(archive_files)}] Would extract: {archive.name}", 'yellow')
            else:
                print_color(f"[{i}/{len(archive_files)}] Extracting: {archive.name}", 'cyan')

                try:
                    result = subprocess.run(
                        [config.seven_zip_path, 'x', f'-o{path}', str(archive), '-r', '-y'],
                        capture_output=True, text=True
                    )

                    if result.returncode != 0:
                        print_color(f"Warning: Failed to extract {archive.name}", 'yellow')
                        log_warning(f"Failed to extract: {archive}")
                        stats.archives_failed += 1
                    else:
                        print_color(f"Successfully extracted {archive.name}", 'green')
                        stats.archives_extracted += 1
                except Exception as e:
                    log_error(f"Error extracting {archive}: {e}")
                    stats.archives_failed += 1

        # Delete archives after extraction
        if delete_after and not config.dry_run:
            print_color("Deleting archive files...", 'yellow')
            for ext in config.archive_extensions:
                for archive in Path(path).rglob(f'*{ext}'):
                    try:
                        archive_size = archive.stat().st_size
                        archive.unlink()
                        stats.files_deleted += 1
                        stats.bytes_deleted += archive_size
                    except:
                        pass

            # Also delete .r00, .r01, etc.
            for archive in Path(path).rglob('*.r[0-9][0-9]'):
                try:
                    archive_size = archive.stat().st_size
                    archive.unlink()
                    stats.files_deleted += 1
                    stats.bytes_deleted += archive_size
                except:
                    pass

        print_color("Archives processed", 'green')
    except Exception as e:
        log_error(f"Error processing archives: {e}")


def create_folders_for_loose_files(path: str):
    """Create individual folders for loose video files"""
    print_color("Creating folders for loose video files...", 'yellow')
    log_info(f"Starting folder creation for loose files in: {path}")

    try:
        root_path = Path(path)
        loose_files = [f for f in root_path.iterdir()
                      if f.is_file() and f.suffix.lower() in config.video_extensions]

        if not loose_files:
            print_color("No loose video files found", 'cyan')
            return

        for file in loose_files:
            folder_name = file.stem
            folder_path = root_path / folder_name

            if config.dry_run:
                print_color(f"[DRY-RUN] Would create folder '{folder_name}' and move: {file.name}", 'yellow')
            else:
                print_color(f"Processing: {file.name}", 'cyan')

                if not folder_path.exists():
                    folder_path.mkdir(parents=True)
                    stats.folders_created += 1

                try:
                    shutil.move(str(file), str(folder_path / file.name))
                    print_color(f"Moved to: {folder_name}", 'green')
                    stats.files_moved += 1
                except Exception as e:
                    log_error(f"Error moving {file.name}: {e}")
    except Exception as e:
        log_error(f"Error organizing loose files: {e}")


def clean_folder_names(path: str):
    """Clean folder names by removing tags and formatting"""
    print_color("Cleaning folder names...", 'yellow')
    log_info(f"Starting folder name cleaning in: {path}")

    try:
        root_path = Path(path)

        for folder in root_path.iterdir():
            if not folder.is_dir() or folder.name == '_Trailers':
                continue

            new_name = folder.name

            # Remove tags
            for tag in config.tags:
                if tag.lower() in new_name.lower():
                    # Split at tag and take first part
                    pattern = re.compile(re.escape(tag), re.IGNORECASE)
                    parts = pattern.split(new_name)
                    if parts[0].strip(' .-'):
                        new_name = parts[0].strip(' .-')

            # Replace dots with spaces
            if '.' in new_name:
                new_name = re.sub(r'\.', ' ', new_name)
                new_name = re.sub(r'\s+', ' ', new_name).strip()

            # Format year with parentheses
            year_match = re.search(r'\s(19|20)(\d{2})$', new_name)
            if year_match:
                new_name = re.sub(r'\s((19|20)\d{2})$', r' (\1)', new_name)

            if new_name != folder.name:
                new_path = root_path / new_name

                if config.dry_run:
                    print_color(f"[DRY-RUN] Would rename '{folder.name}' to '{new_name}'", 'yellow')
                else:
                    try:
                        if not new_path.exists():
                            folder.rename(new_path)
                            log_info(f"Renamed '{folder.name}' to '{new_name}'")
                            stats.folders_renamed += 1
                    except Exception as e:
                        log_error(f"Error renaming {folder.name}: {e}")

        print_color("Folder names cleaned", 'green')
    except Exception as e:
        log_error(f"Error cleaning folder names: {e}")


def remove_empty_folders(path: str):
    """Remove empty folders recursively"""
    print_color("Cleaning up empty folders...", 'yellow')
    log_info(f"Starting empty folder cleanup in: {path}")

    try:
        removed = 0
        for root, dirs, files in os.walk(path, topdown=False):
            for d in dirs:
                dir_path = os.path.join(root, d)
                try:
                    if not os.listdir(dir_path):
                        if config.dry_run:
                            print_color(f"[DRY-RUN] Would remove: {dir_path}", 'yellow')
                        else:
                            os.rmdir(dir_path)
                            stats.empty_folders_removed += 1
                            removed += 1
                except:
                    pass

        if removed > 0:
            print_color(f"Removed {removed} empty folder(s)", 'green')
        else:
            print_color("No empty folders found", 'cyan')
    except Exception as e:
        log_error(f"Error removing empty folders: {e}")


def organize_seasons(path: str):
    """Organize TV episodes into season folders"""
    print_color("Organizing episodes into season folders...", 'yellow')
    log_info(f"Starting season organization in: {path}")

    try:
        video_files = []
        for ext in config.video_extensions:
            video_files.extend(Path(path).rglob(f'*{ext}'))

        if not video_files:
            print_color("No video files found", 'cyan')
            return

        print_color(f"Found {len(video_files)} video file(s)", 'cyan')
        organized = 0

        for file in video_files:
            ep_info = get_episode_info(file.name)

            if ep_info['season'] is not None:
                season_folder = f"Season {ep_info['season']:02d}"
                season_path = Path(path) / season_folder

                # Skip if already in correct folder
                if file.parent.name == season_folder:
                    print_color(f"Already organized: {file.name}", 'gray')
                    continue

                if config.dry_run:
                    print_color(f"[DRY-RUN] Would move to {season_folder}: {file.name}", 'yellow')
                else:
                    if not season_path.exists():
                        season_path.mkdir(parents=True)
                        print_color(f"Created folder: {season_folder}", 'green')
                        stats.folders_created += 1

                    try:
                        dest_path = season_path / file.name
                        shutil.move(str(file), str(dest_path))
                        print_color(f"Moved to {season_folder}: {file.name}", 'cyan')
                        stats.files_moved += 1
                        organized += 1
                    except Exception as e:
                        log_warning(f"Error moving {file.name}: {e}")
            else:
                print_color(f"Could not parse: {file.name}", 'yellow')

        if organized > 0:
            print_color(f"Organized {organized} episode(s) into season folders", 'green')
    except Exception as e:
        log_error(f"Error organizing seasons: {e}")


# =============================================================================
# NFO FUNCTIONS
# =============================================================================

def read_nfo_file(nfo_path: str) -> Optional[Dict[str, Any]]:
    """Parse an existing NFO file"""
    metadata = {
        'title': None,
        'original_title': None,
        'year': None,
        'plot': None,
        'rating': None,
        'imdb_id': None,
        'tmdb_id': None,
        'genres': [],
        'directors': [],
        'actors': []
    }

    try:
        tree = ET.parse(nfo_path)
        root = tree.getroot()
        stats.nfo_files_read += 1

        # Movie NFO
        if root.tag == 'movie':
            metadata['title'] = root.findtext('title')
            metadata['original_title'] = root.findtext('originaltitle')
            metadata['year'] = root.findtext('year')
            metadata['plot'] = root.findtext('plot')
            metadata['rating'] = root.findtext('rating')

            # Get IDs
            for uniqueid in root.findall('uniqueid'):
                if uniqueid.get('type') == 'imdb':
                    metadata['imdb_id'] = uniqueid.text
                elif uniqueid.get('type') == 'tmdb':
                    metadata['tmdb_id'] = uniqueid.text

            metadata['genres'] = [g.text for g in root.findall('genre') if g.text]
            metadata['directors'] = [d.text for d in root.findall('director') if d.text]

        return metadata
    except Exception as e:
        log_error(f"Error parsing NFO {nfo_path}: {e}")
        return None


def create_movie_nfo(video_path: str, title: str = None, year: str = None):
    """Create a basic Kodi NFO file for a movie"""
    try:
        video_file = Path(video_path)
        nfo_path = video_file.parent / f"{video_file.stem}.nfo"

        if nfo_path.exists():
            print_color(f"NFO already exists: {video_file.name}", 'cyan')
            return

        # Extract title/year from folder name if not provided
        if not title or not year:
            folder_name = video_file.parent.name
            title_info = get_normalized_title(folder_name)

            if not title:
                title = folder_name
                # Clean up
                if title_info['year']:
                    title = re.sub(rf'\s*[\(\[]?{title_info["year"]}[\)\]]?.*$', '', title)
                title = re.sub(r'\.', ' ', title).strip()

            if not year:
                year = title_info['year']

        if config.dry_run:
            print_color(f"[DRY-RUN] Would create NFO for: {title} ({year})", 'yellow')
            return

        # Create NFO XML
        movie = ET.Element('movie')
        ET.SubElement(movie, 'title').text = title
        if year:
            ET.SubElement(movie, 'year').text = year
        ET.SubElement(movie, 'plot')
        ET.SubElement(movie, 'outline')
        ET.SubElement(movie, 'tagline')
        ET.SubElement(movie, 'runtime')
        ET.SubElement(movie, 'thumb')
        ET.SubElement(movie, 'fanart')

        # Pretty print
        xml_str = minidom.parseString(ET.tostring(movie)).toprettyxml(indent='    ')
        xml_str = '\n'.join(xml_str.split('\n')[1:])  # Remove XML declaration line
        xml_str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str

        with open(nfo_path, 'w', encoding='utf-8') as f:
            f.write(xml_str)

        print_color(f"Created NFO: {title} ({year})", 'green')
        stats.nfo_files_created += 1
    except Exception as e:
        log_error(f"Error creating NFO for {video_path}: {e}")


# =============================================================================
# DUPLICATE DETECTION
# =============================================================================

def find_duplicate_movies(path: str) -> List[List[Dict]]:
    """Find potential duplicate movies"""
    print_color("\nScanning for duplicate movies...", 'yellow')
    log_info(f"Starting duplicate scan in: {path}")

    try:
        root_path = Path(path)
        movie_folders = [f for f in root_path.iterdir()
                        if f.is_dir() and f.name != '_Trailers']

        if not movie_folders:
            print_color("No movie folders found", 'cyan')
            return []

        # Build lookup
        title_lookup = {}

        for folder in movie_folders:
            title_info = get_normalized_title(folder.name)
            quality = get_quality_score(folder.name)

            # Find main video file
            video_files = [f for f in folder.iterdir()
                         if f.is_file() and f.suffix.lower() in config.video_extensions]
            file_size = max((f.stat().st_size for f in video_files), default=0) if video_files else 0

            entry = {
                'path': str(folder),
                'original_name': folder.name,
                'year': title_info['year'],
                'quality': quality,
                'file_size': file_size
            }

            key = title_info['normalized_title']
            if title_info['year']:
                key = f"{key}|{title_info['year']}"

            if key not in title_lookup:
                title_lookup[key] = []
            title_lookup[key].append(entry)

        # Find duplicates
        duplicates = [entries for entries in title_lookup.values() if len(entries) > 1]

        return duplicates
    except Exception as e:
        log_error(f"Error scanning for duplicates: {e}")
        return []


def show_duplicate_report(path: str):
    """Display duplicate movies report"""
    duplicates = find_duplicate_movies(path)

    if not duplicates:
        print_color("No duplicate movies found!", 'green')
        return

    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'yellow')
    print_color("║                    DUPLICATE REPORT                          ║", 'yellow')
    print_color("╠══════════════════════════════════════════════════════════════╣", 'yellow')
    print_color(f"║  Found {len(duplicates)} potential duplicate group(s)                        ║", 'yellow')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'yellow')

    for group_num, group in enumerate(duplicates, 1):
        print_color(f"\n[{group_num}] Potential Duplicates:", 'cyan')

        # Sort by quality score
        sorted_group = sorted(group, key=lambda x: x['quality']['score'], reverse=True)

        for i, movie in enumerate(sorted_group):
            size_str = format_file_size(movie['file_size'])
            score = movie['quality']['score']
            resolution = movie['quality']['resolution']
            source = movie['quality']['source']

            if i == 0:
                print_color(f"  [KEEP] ", 'green', end='')
            else:
                print_color(f"  [DEL?] ", 'red', end='')

            print_color(movie['original_name'], 'white')
            print_color(f"         Score: {score} | {size_str} | {resolution} {source}", 'gray')


# =============================================================================
# HEALTH CHECK
# =============================================================================

def health_check(path: str, media_type: str = "Movies"):
    """Perform library health check"""
    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                  LIBRARY HEALTH CHECK                        ║", 'cyan')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')

    log_info(f"Starting health check for: {path}")

    issues = {
        'empty_folders': [],
        'no_video_files': [],
        'zero_byte_files': [],
        'orphaned_subtitles': [],
        'missing_nfo': [],
        'small_videos': [],
        'naming_issues': []
    }

    root_path = Path(path)

    # Check empty folders
    print_color("\nChecking for empty folders...", 'yellow')
    for folder in root_path.rglob('*'):
        if folder.is_dir() and not list(folder.iterdir()):
            issues['empty_folders'].append(folder)

    # Check folders without video files
    print_color("Checking for folders without video files...", 'yellow')
    for folder in root_path.iterdir():
        if folder.is_dir() and folder.name != '_Trailers':
            has_video = any(f.suffix.lower() in config.video_extensions
                          for f in folder.iterdir() if f.is_file())
            if not has_video:
                issues['no_video_files'].append(folder)

    # Check zero-byte files
    print_color("Checking for corrupted/zero-byte files...", 'yellow')
    for file in root_path.rglob('*'):
        if file.is_file() and file.stat().st_size == 0:
            issues['zero_byte_files'].append(file)

    # Check small videos (under 50MB, not samples/trailers)
    print_color("Checking for suspiciously small video files...", 'yellow')
    for file in root_path.rglob('*'):
        if (file.is_file() and
            file.suffix.lower() in config.video_extensions and
            file.stat().st_size < 50 * 1024 * 1024 and
            not any(x in file.name.lower() for x in ['sample', 'trailer', 'teaser'])):
            issues['small_videos'].append(file)

    # Display results
    print_color("\n=== Health Check Results ===", 'cyan')

    total_issues = 0

    if issues['empty_folders']:
        print_color(f"\nEmpty Folders ({len(issues['empty_folders'])}):", 'yellow')
        for f in issues['empty_folders'][:10]:
            print_color(f"  - {f}", 'gray')
        total_issues += len(issues['empty_folders'])

    if issues['no_video_files']:
        print_color(f"\nFolders Without Video Files ({len(issues['no_video_files'])}):", 'yellow')
        for f in issues['no_video_files'][:10]:
            print_color(f"  - {f.name}", 'gray')
        total_issues += len(issues['no_video_files'])

    if issues['zero_byte_files']:
        print_color(f"\nZero-Byte Files ({len(issues['zero_byte_files'])}):", 'red')
        for f in issues['zero_byte_files'][:10]:
            print_color(f"  - {f}", 'gray')
        total_issues += len(issues['zero_byte_files'])

    if issues['small_videos']:
        print_color(f"\nSuspiciously Small Videos ({len(issues['small_videos'])}):", 'yellow')
        for f in issues['small_videos'][:10]:
            print_color(f"  - {f.name} ({format_file_size(f.stat().st_size)})", 'gray')
        total_issues += len(issues['small_videos'])

    print()
    if total_issues == 0:
        print_color("Library is healthy! No issues found.", 'green')
    else:
        print_color(f"Found {total_issues} issue(s) in the library.", 'yellow')


# =============================================================================
# CODEC ANALYSIS
# =============================================================================

def codec_analysis(path: str):
    """Analyze video codecs in library"""
    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                    CODEC ANALYSIS                            ║", 'cyan')
    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')

    log_info(f"Starting codec analysis for: {path}")

    video_files = []
    for ext in config.video_extensions:
        video_files.extend(Path(path).rglob(f'*{ext}'))

    if not video_files:
        print_color("No video files found", 'cyan')
        return

    print_color(f"\nAnalyzing {len(video_files)} video file(s)...", 'yellow')

    analysis = {
        'total_files': len(video_files),
        'total_size': 0,
        'by_resolution': {},
        'by_codec': {},
        'by_container': {},
        'need_transcode': []
    }

    for file in video_files:
        file_size = file.stat().st_size
        analysis['total_size'] += file_size

        quality = get_quality_score(file.name)
        container = file.suffix.upper().lstrip('.')

        # Count by resolution
        res = quality['resolution']
        analysis['by_resolution'][res] = analysis['by_resolution'].get(res, 0) + 1

        # Count by codec
        codec = quality['codec']
        analysis['by_codec'][codec] = analysis['by_codec'].get(codec, 0) + 1

        # Count by container
        analysis['by_container'][container] = analysis['by_container'].get(container, 0) + 1

        # Check if needs transcoding
        needs_transcode = False
        reasons = []

        if quality['codec'] == 'HEVC/x265':
            needs_transcode = True
            reasons.append("HEVC may not play on older devices")
        if quality['resolution'] == '2160p':
            needs_transcode = True
            reasons.append("4K may require transcoding for streaming")
        if quality['hdr']:
            needs_transcode = True
            reasons.append("HDR requires tone mapping for SDR displays")
        if container == 'AVI':
            needs_transcode = True
            reasons.append("AVI container is outdated")

        if needs_transcode:
            analysis['need_transcode'].append({
                'path': str(file),
                'filename': file.name,
                'size': file_size,
                'resolution': res,
                'codec': codec,
                'container': container,
                'reasons': '; '.join(reasons)
            })

    # Display results
    print_color("\n=== Library Statistics ===", 'cyan')
    print_color(f"Total Files: {analysis['total_files']}", 'white')
    print_color(f"Total Size: {format_file_size(analysis['total_size'])}", 'white')

    print_color("\n=== By Resolution ===", 'cyan')
    for res, count in sorted(analysis['by_resolution'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {res}: {count} ({pct}%)", 'white')

    print_color("\n=== By Video Codec ===", 'cyan')
    for codec, count in sorted(analysis['by_codec'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {codec}: {count} ({pct}%)", 'white')

    print_color("\n=== By Container ===", 'cyan')
    for container, count in sorted(analysis['by_container'].items(), key=lambda x: x[1], reverse=True):
        pct = round((count / analysis['total_files']) * 100, 1)
        print_color(f"  {container}: {count} ({pct}%)", 'white')

    if analysis['need_transcode']:
        print_color(f"\n=== Potential Transcoding Queue ===", 'yellow')
        print_color(f"Found {len(analysis['need_transcode'])} file(s) that may benefit from transcoding:", 'white')

        for item in analysis['need_transcode'][:10]:
            print_color(f"\n  {item['filename']}", 'white')
            print_color(f"    Size: {format_file_size(item['size'])} | {item['resolution']} | {item['codec']}", 'gray')
            print_color(f"    Reason: {item['reasons']}", 'yellow')

        if len(analysis['need_transcode']) > 10:
            print_color(f"\n  ... and {len(analysis['need_transcode']) - 10} more files", 'gray')
    else:
        print_color("\nNo files require transcoding for compatibility.", 'green')


# =============================================================================
# TMDB INTEGRATION
# =============================================================================

def search_tmdb_movie(title: str, year: str = None, api_key: str = None) -> Optional[Dict]:
    """Search TMDB for a movie"""
    if not HAS_REQUESTS or not api_key:
        return None

    try:
        url = f"https://api.themoviedb.org/3/search/movie"
        params = {'api_key': api_key, 'query': title}
        if year:
            params['year'] = year

        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()

        if data.get('results'):
            movie = data['results'][0]
            return {
                'id': movie['id'],
                'title': movie['title'],
                'original_title': movie.get('original_title'),
                'year': movie.get('release_date', '')[:4] if movie.get('release_date') else None,
                'overview': movie.get('overview'),
                'rating': movie.get('vote_average'),
                'poster_path': f"https://image.tmdb.org/t/p/w500{movie['poster_path']}" if movie.get('poster_path') else None
            }
        return None
    except Exception as e:
        log_error(f"Error searching TMDB: {e}")
        return None


def get_tmdb_movie_details(movie_id: int, api_key: str) -> Optional[Dict]:
    """Get detailed movie info from TMDB"""
    if not HAS_REQUESTS or not api_key:
        return None

    try:
        url = f"https://api.themoviedb.org/3/movie/{movie_id}"
        params = {'api_key': api_key, 'append_to_response': 'credits,external_ids'}

        response = requests.get(url, params=params)
        response.raise_for_status()
        movie = response.json()

        directors = [c['name'] for c in movie.get('credits', {}).get('crew', []) if c.get('job') == 'Director']
        cast = [{'name': c['name'], 'role': c.get('character', '')}
                for c in movie.get('credits', {}).get('cast', [])[:10]]

        return {
            'id': movie['id'],
            'title': movie['title'],
            'original_title': movie.get('original_title'),
            'tagline': movie.get('tagline'),
            'year': movie.get('release_date', '')[:4] if movie.get('release_date') else None,
            'overview': movie.get('overview'),
            'rating': movie.get('vote_average'),
            'votes': movie.get('vote_count'),
            'runtime': movie.get('runtime'),
            'genres': [g['name'] for g in movie.get('genres', [])],
            'studios': [s['name'] for s in movie.get('production_companies', [])],
            'directors': directors,
            'cast': cast,
            'imdb_id': movie.get('external_ids', {}).get('imdb_id'),
            'poster_path': f"https://image.tmdb.org/t/p/w500{movie['poster_path']}" if movie.get('poster_path') else None,
            'backdrop_path': f"https://image.tmdb.org/t/p/original{movie['backdrop_path']}" if movie.get('backdrop_path') else None
        }
    except Exception as e:
        log_error(f"Error getting TMDB details: {e}")
        return None


# =============================================================================
# MAIN PROCESSING
# =============================================================================

def process_movies(path: str):
    """Process a movie library"""
    print_color("\nMovie Routine", 'magenta')
    log_info(f"Movie routine started for path: {path}")

    if config.dry_run:
        print_color("\n*** DRY-RUN MODE - Previewing changes only ***\n", 'yellow')

    # Step 1: Process trailers
    if config.keep_trailers:
        move_trailers_to_folder(path)

    # Step 2: Remove unnecessary files
    remove_unnecessary_files(path)

    # Step 3: Process subtitles
    process_subtitles(path)

    # Step 4: Extract archives
    extract_archives(path, delete_after=True)

    # Step 5: Create folders for loose files
    create_folders_for_loose_files(path)

    # Step 6: Clean folder names
    clean_folder_names(path)

    # Step 7: Check for duplicates
    if config.check_duplicates:
        show_duplicate_report(path)

    print_color("\nMovie routine completed!", 'magenta')


def process_tv_shows(path: str):
    """Process a TV show library"""
    print_color("\nTV Show Routine", 'magenta')
    log_info(f"TV Show routine started for path: {path}")

    if config.dry_run:
        print_color("\n*** DRY-RUN MODE - Previewing changes only ***\n", 'yellow')

    # Step 1: Extract archives
    extract_archives(path, delete_after=True)

    # Step 2: Remove unnecessary files
    remove_unnecessary_files(path)

    # Step 3: Process subtitles
    process_subtitles(path)

    # Step 4: Organize into season folders
    if config.organize_seasons:
        organize_seasons(path)

    # Step 5: Remove empty folders
    remove_empty_folders(path)

    print_color("\nTV Show routine completed!", 'magenta')


def show_statistics():
    """Display processing statistics"""
    stats.end_time = datetime.now()
    duration = stats.end_time - stats.start_time

    print_color("\n╔══════════════════════════════════════════════════════════════╗", 'cyan')
    print_color("║                    PROCESSING SUMMARY                        ║", 'cyan')
    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    print_color(f"║  Duration:              {str(duration).split('.')[0]:<39}║", 'cyan')

    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    if stats.files_deleted > 0:
        print_color(f"║  Files Deleted:         {stats.files_deleted:<39}║", 'cyan')
    if stats.bytes_deleted > 0:
        print_color(f"║  Space Reclaimed:       {format_file_size(stats.bytes_deleted):<39}║", 'cyan')
    if stats.archives_extracted > 0:
        print_color(f"║  Archives Extracted:    {stats.archives_extracted:<39}║", 'cyan')
    if stats.folders_created > 0:
        print_color(f"║  Folders Created:       {stats.folders_created:<39}║", 'cyan')
    if stats.folders_renamed > 0:
        print_color(f"║  Folders Renamed:       {stats.folders_renamed:<39}║", 'cyan')
    if stats.files_moved > 0:
        print_color(f"║  Files Moved:           {stats.files_moved:<39}║", 'cyan')
    if stats.trailers_moved > 0:
        print_color(f"║  Trailers Moved:        {stats.trailers_moved:<39}║", 'cyan')
    if stats.nfo_files_created > 0:
        print_color(f"║  NFO Files Created:     {stats.nfo_files_created:<39}║", 'cyan')

    print_color("╠══════════════════════════════════════════════════════════════╣", 'cyan')

    error_color = 'red' if stats.errors > 0 else 'green'
    warning_color = 'yellow' if stats.warnings > 0 else 'green'
    print_color(f"║  Errors:                {stats.errors:<39}║", 'cyan')
    print_color(f"║  Warnings:              {stats.warnings:<39}║", 'cyan')

    print_color("╚══════════════════════════════════════════════════════════════╝", 'cyan')


# =============================================================================
# CLI INTERFACE
# =============================================================================

def main():
    """Main entry point"""
    global config, stats, logger

    parser = argparse.ArgumentParser(
        description='MediaCleaner - Cross-platform media library organization tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  %(prog)s --movies /path/to/movies
  %(prog)s --tvshows /path/to/shows --dry-run
  %(prog)s --health-check /path/to/library
  %(prog)s --codec-analysis /path/to/library
        '''
    )

    parser.add_argument('--movies', metavar='PATH', help='Process movie library at PATH')
    parser.add_argument('--tvshows', metavar='PATH', help='Process TV show library at PATH')
    parser.add_argument('--health-check', metavar='PATH', help='Run health check on library')
    parser.add_argument('--codec-analysis', metavar='PATH', help='Analyze codecs in library')
    parser.add_argument('--duplicates', metavar='PATH', help='Find duplicate movies')

    parser.add_argument('--dry-run', action='store_true', help='Preview changes without making them')
    parser.add_argument('--no-subtitles', action='store_true', help='Delete all subtitle files')
    parser.add_argument('--no-trailers', action='store_true', help='Delete trailers instead of moving')
    parser.add_argument('--no-organize-seasons', action='store_true', help='Skip season folder organization')
    parser.add_argument('--check-duplicates', action='store_true', help='Check for duplicate movies')
    parser.add_argument('--generate-nfo', action='store_true', help='Generate NFO files')

    parser.add_argument('--tmdb-key', metavar='KEY', help='TMDB API key for metadata fetching')
    parser.add_argument('--log-file', metavar='FILE', help='Custom log file path')

    parser.add_argument('--version', action='version', version='MediaCleaner 5.0')

    args = parser.parse_args()

    # Configure
    config.dry_run = args.dry_run
    config.keep_subtitles = not args.no_subtitles
    config.keep_trailers = not args.no_trailers
    config.organize_seasons = not args.no_organize_seasons
    config.check_duplicates = args.check_duplicates
    config.generate_nfo = args.generate_nfo
    config.tmdb_api_key = args.tmdb_key or ''

    if args.log_file:
        config.log_file = args.log_file

    # Setup logging
    logger = setup_logging(config.log_file)

    # Display header
    print_color("\n=== MediaCleaner v5.0 (Python) ===", 'cyan')
    print_color("Cross-platform media library organization tool\n", 'gray')

    if config.dry_run:
        print_color("DRY-RUN MODE ENABLED - No changes will be made\n", 'yellow')

    # Process based on arguments
    if args.movies:
        if not os.path.isdir(args.movies):
            print_color(f"Error: '{args.movies}' is not a valid directory", 'red')
            sys.exit(1)
        process_movies(args.movies)
        show_statistics()

    elif args.tvshows:
        if not os.path.isdir(args.tvshows):
            print_color(f"Error: '{args.tvshows}' is not a valid directory", 'red')
            sys.exit(1)
        process_tv_shows(args.tvshows)
        show_statistics()

    elif args.health_check:
        if not os.path.isdir(args.health_check):
            print_color(f"Error: '{args.health_check}' is not a valid directory", 'red')
            sys.exit(1)
        health_check(args.health_check)

    elif args.codec_analysis:
        if not os.path.isdir(args.codec_analysis):
            print_color(f"Error: '{args.codec_analysis}' is not a valid directory", 'red')
            sys.exit(1)
        codec_analysis(args.codec_analysis)

    elif args.duplicates:
        if not os.path.isdir(args.duplicates):
            print_color(f"Error: '{args.duplicates}' is not a valid directory", 'red')
            sys.exit(1)
        show_duplicate_report(args.duplicates)

    else:
        # Interactive mode
        print_color("Select an option:", 'cyan')
        print_color("  1. Process Movies", 'white')
        print_color("  2. Process TV Shows", 'white')
        print_color("  3. Health Check", 'white')
        print_color("  4. Codec Analysis", 'white')
        print_color("  5. Find Duplicates", 'white')

        choice = input("\nEnter choice (1-5): ").strip()

        if choice in ['1', '2', '3', '4', '5']:
            path = input("Enter path to media library: ").strip()

            if not os.path.isdir(path):
                print_color(f"Error: '{path}' is not a valid directory", 'red')
                sys.exit(1)

            dry_run_input = input("Enable dry-run mode? (y/N): ").strip().lower()
            config.dry_run = dry_run_input == 'y'

            if choice == '1':
                process_movies(path)
                show_statistics()
            elif choice == '2':
                process_tv_shows(path)
                show_statistics()
            elif choice == '3':
                health_check(path)
            elif choice == '4':
                codec_analysis(path)
            elif choice == '5':
                show_duplicate_report(path)
        else:
            print_color("Invalid choice", 'red')
            sys.exit(1)

    print_color(f"\nLog file saved to: {config.log_file}", 'cyan')


if __name__ == '__main__':
    main()
