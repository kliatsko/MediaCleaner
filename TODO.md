# TODO

## Features

### Initial Setup Wizard
- [ ] Create first-run setup process that detects missing config
- [ ] Prompt user for media library paths
- [ ] Prompt for preferred codecs and quality settings
- [ ] Prompt for FFmpeg path (or auto-detect)
- [ ] Save answers to config file so script doesn't ask on every run
- [ ] Add `--setup` or `--reconfigure` flag to re-run setup

### User-Friendly Distribution
- [ ] Create installer package (MSI or self-extracting exe)
- [ ] Bundle FFmpeg or provide guided download
- [ ] Add auto-updater functionality to check for new versions
- [ ] Create simple "double-click to run" experience (no PowerShell knowledge required)
- [ ] Add uninstaller
- [ ] Consider a simple GUI wrapper for non-technical users

### Version Management
- [ ] Add `$Version` variable at top of script as single source of truth
- [ ] Add `--version` / `-Version` flag to display current version
- [ ] Use Semantic Versioning (MAJOR.MINOR.PATCH) format
- [ ] Create CHANGELOG.md to document changes per release
- [ ] Tag GitHub releases to match script version
- [ ] (Optional) Add PowerShell module manifest (.psd1) for future PowerShell Gallery publishing

## Improvements

- [ ] Improve README with step-by-step first-timer instructions
- [ ] Add screenshots/examples to documentation
- [ ] Create a "Quick Start" guide

## Bugs

- (none currently)

## Technical Debt

- (none currently)
