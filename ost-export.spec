Name:           ost-export
Version:        1.0.0
Release:        1%{?dist}
Summary:        A tool to export emails and attachments from Outlook OST files to MBOX or EML format

License:        MIT
Group:          Applications/System
URL:            https://github.com/svavs/ost-export

# Disable automatic dependency generation
AutoReq: 0
AutoProv: 0

BuildArch:      noarch

# Build dependencies
BuildRequires:  python3-devel
%if 0%{?fedora} || 0%{?rhel} > 0
BuildRequires:  python3-pip
%else
BuildRequires:  python-pip
%endif

# Runtime dependencies
Requires:       python3
%if 0%{?fedora} || 0%{?rhel} > 0
Requires:       python3-beautifulsoup4
%else
Requires:       python-beautifulsoup4
%endif

%description
A Python utility to export emails and attachments from Outlook OST files to MBOX or EML format.
The tool preserves email metadata and handles various attachment types with proper MIME type detection.

%prep
# No prep needed as we're using files directly

%build
# Nothing to build as it's a pure Python package

%install
# Create directories
install -d %{buildroot}%{_bindir}
install -d %{buildroot}%{python3_sitelib}
install -d %{buildroot}%{_docdir}/%{name}

# Install the script
install -p -m 755 %{_sourcedir}/ost_export.py %{buildroot}%{python3_sitelib}/

# Create a wrapper script
cat > %{buildroot}%{_bindir}/ost-export << 'EOF'
#!/bin/sh
exec python3 -m ost_export "$@"
EOF
chmod 755 %{buildroot}%{_bindir}/ost-export

# Install documentation
install -p -m 644 %{_sourcedir}/README.md %{buildroot}%{_docdir}/%{name}/
install -p -m 644 %{_sourcedir}/LICENSE %{buildroot}%{_docdir}/%{name}/

%files
%doc README.md LICENSE
%{_bindir}/ost-export
%{python3_sitelib}/ost_export.py

# Exclude __pycache__ directories to avoid file not found errors
%exclude %{python3_sitelib}/__pycache__

%post
# Install libpff-python via pip for all distributions
if ! python3 -c "import pypff" 2>/dev/null; then
    echo "Installing libpff-python via pip..."
    pip3 install --user --no-warn-script-location libpff-python || :
fi

# Create any necessary directories
mkdir -p %{_sysconfdir}/rpm/macros.d/

%preun
# Clean up pip installation on uninstall for all distributions
if [ $1 -eq 0 ]; then  # Only on complete uninstall, not upgrade
    echo "Removing libpff-python..."
    pip3 uninstall -y libpff-python || :
fi

%changelog
* Tue Jul 01 2025 Silvano Sallese <mail@mymail.com> - 1.0.0-1
- Initial package
