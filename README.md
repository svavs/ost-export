# OST to MBOX/EML Exporter

A Python utility to export emails and attachments from Outlook OST files to MBOX or EML format, with support for proper MIME type detection and handling of various attachments including PDFs.

## Features

- Export Outlook OST files to MBOX format (compatible with Thunderbird, Apple Mail, etc.)
- Export to individual EML files (one file per message)
- Preserves email metadata (subject, sender, recipients, dates)
- Handles various attachment types with proper MIME type detection
- Special handling for PDF attachments to ensure proper display in email clients
- Preserves folder structure in the export
- Robust error handling and logging

## Prerequisites

- Python 3.6 or higher
- libpff library (required for OST file parsing)
- Python packages listed in `requirements.txt`

## Installation

### Linux

#### Debian/Ubuntu

1. Install the required system dependencies:
   ```bash
   sudo apt-get update
   sudo apt-get install -y python3-pip python3-dev libpff-dev
   ```

#### OpenMandriva/RPM-based distributions

##### Option 1: Using ABB (recommended for OpenMandriva)

You can build the RPM package using the OpenMandriva ABF (Automatic Build Farm) system:

1. Install the ABB tool:
   ```bash
   sudo dnf install abb-core
   ```

2. Build the package locally:
   ```bash
   # Clone the repository
   git clone https://github.com/svavs/ost-export.git
   cd ost-export
   
   # Build the package
   abb build
   ```

##### Option 2: Manual RPM Build

You can also build the RPM package manually:

1. Install build dependencies:
   ```bash
   # Fedora
   sudo dnf install -y rpm-build python3-devel python3-setuptools python3-pip

   # OpenMandriva
   sudo dnf install -y rpm-build lib64python-devel python-setuptools python-pip
   ```

2. Clone the repository and build the RPM:
   ```bash
   # Clone the repository
   git clone https://github.com/svavs/ost-export.git
   cd ost-export
   
   # Create local rpmbuild directory structure
   mkdir -p rpmbuild/{BUILD,RPMS,SOURCES,SPECS,SRPMS}
   
   # Copy source files
   cp ost_export.py README.md LICENSE rpmbuild/SOURCES/
   cp ost-export.spec rpmbuild/SPECS/
   
   # Build the RPM using the local rpmbuild directory
   rpmbuild -ba rpmbuild/SPECS/ost-export.spec \
     --define "_topdir $(pwd)/rpmbuild" \
     --define "_sourcedir %{_topdir}/SOURCES"
   ```

3. Install the built package with dependencies:
   
   ```bash
   # Install beautifulsoup4 from package manager
   # Fedora
   sudo dnf install -y python3-beautifulsoup4

   # OpenMandriva
   sudo dnf install -y python-beautifulsoup4
   
   # Install the package (libpff-python will be installed via pip during post-install)
   sudo dnf install rpmbuild/RPMS/noarch/ost-export-1.0.0-1.noarch.rpm
   ```
   
   Note: The package will automatically install `libpff-python` via pip during installation.

   The package will be installed with:
   - Main script: `/usr/lib/python3.11/site-packages/ost_export.py`
   - Wrapper script: `/usr/bin/ost-export`
   - Documentation: `/usr/share/doc/ost-export/`

   Important: The package requires Python 3.6 or higher and pip3 for dependency installation.
   
   # The resulting RPM will be in the RPMS directory
   ```

3. Install the built package with dependencies:
   
   ```bash
   # Install beautifulsoup4 from package manager
   sudo dnf install -y python3-beautifulsoup4
   
   # Install the package (libpff-python will be installed via pip during post-install)
   sudo dnf install rpmbuild/RPMS/noarch/ost-export-*.rpm
   ```
   
   Note: The package will automatically install `libpff-python` via pip during installation.
         If the pypff library is not found, you can install it manually using:
         ```bash
         pip install --user libpff-python
         ```

##### Manual Installation

1. Install the required system dependencies:
   ```bash
   # OpenMandriva
   sudo dnf install python-pip lib64python-devel
   ```

#### Using pip (works on most Linux distributions)

1. Install libpff Python bindings:
   ```bash
   pip install --user libpff-python
   ```
   
   Note: You might need to add `~/.local/bin` to your PATH if it's not already there.

2. Install Python dependencies:
   ```bash
   pip install --user -r requirements.txt
   ```

### macOS (using Homebrew)

1. Install Homebrew if you don't have it:
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```

2. Install dependencies:
   ```bash
   brew install libpff
   pip install -r requirements.txt
   ```

### Windows

1. Download and install the latest Python 3.x from [python.org](https://www.python.org/downloads/)
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```
3. For libpff on Windows, you may need to use a pre-built binary or build from source.

## Usage

### Basic Usage

```bash
python ost_export.py <path_to_ost_file> <output_directory> <format>
```

- `<path_to_ost_file>`: Path to your Outlook OST file
- `<output_directory>`: Directory where the exported files will be saved
- `<format>`: Export format - either `mbox` or `eml`

### Examples

Export to MBOX format:
```bash
python ost_export.py "path/to/outlook.ost" "./exported_emails" mbox
```

Export to EML format (individual files):
```bash
python ost_export.py "path/to/outlook.ost" "./exported_emails" eml
```

### Importing into Email Clients

#### Thunderbird (MBOX)
1. Install the "ImportExportTools NG" add-on
2. Right-click on a folder in Thunderbird
3. Select "ImportExportTools NG" > "Import mbox file"
4. Choose "Import directly one or more mbox files"
5. Select the exported .mbox files

#### Thunderbird (EML)
1. In Thunderbird, right-click on the destination folder
2. Select "ImportExportTools NG" > "Import all messages from a directory"
3. Select the directory containing the .eml files

## Requirements

The `requirements.txt` file contains the following Python packages:

```
beautifulsoup4>=4.9.3
libpff-python>=20220124
```

These dependencies will be automatically installed when you run:
```bash
pip install -r requirements.txt
```

## Troubleshooting

### Common Issues

1. **Permission errors**:
   - Make sure you have read access to the OST file and write access to the output directory

2. **Attachment issues**:

   - **PDF Attachments**:
     - PDFs are automatically detected by both file extension and content
     - Special MIME headers are added to ensure proper display in email clients
     - Embedded PDFs in messages should be properly extracted

   - **Common Problems**:
     - Some attachments might be corrupted or not properly exported due to limitations in the pypff library
     - Very large attachments (over 50MB) might cause issues with some email clients
     - Some embedded objects (like Outlook-specific items) might not be properly converted

   - **Troubleshooting**:
     - If an attachment appears corrupted, try opening it with a hex editor to verify the content
     - Check the export logs for any specific error messages about the problematic attachment
     - Some email clients have size limits for attachments in MBOX format
     
   - **Known Limitations**:
     - Password-protected attachments cannot be processed
     - Some complex embedded objects may not be fully supported
     - The original attachment filenames might be lost for some message formats
     
   - **Verification**:
     - After export, verify that the number of attachments in the original message matches the exported version
     - Check that file sizes are roughly equivalent (small differences in size are normal due to encoding)

## License

This project is open source and available under the MIT License.

## Acknowledgments

- [libpff](https://github.com/libyal/libpff) - Library for accessing Personal Folder (OST/PST) files
- [pypff](https://github.com/libyal/libpff/wiki/Development) - Python bindings for libpff
