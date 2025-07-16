# Changelog
All notable changes to the Reddit Persona Analyzer project will be documented in this file.

## [1.2.0] - 2025-07-16
### Added
- Added professional Excel (.xlsx) export with proper formatting
- Implemented cell styling, colors, and auto-adjusting column widths
- Added automatic file opening after export
- Created dedicated exports directory for generated files

### Changed
- Removed CSV export functionality completely
- Improved error handling for export operations
- Updated .gitignore to exclude exports directory
- Enhanced README.md with new Excel export instructions

### Fixed
- Resolved 'Namespace' object has no attribute 'csv' error
- Fixed command-line argument parsing for export options
- Improved file handling and cleanup

## [1.0.0] - 2025-01-01
### Added
- Initial release of Reddit Persona Analyzer
- Basic Reddit user analysis functionality
- Google Sheets export capability
- Command-line interface
