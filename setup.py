from setuptools import setup, find_packages

setup(
    name="ost-export",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'beautifulsoup4>=4.9.3',
        'libpff-python>=20220124',
    ],
    entry_points={
        'console_scripts': [
            'ost-export=ost_export:main',
        ],
    },
    author="Silvano Sallese",
    author_email="mail@mymail.com",
    description="A tool to export emails and attachments from Outlook OST files to MBOX or EML format",
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/svavs/ost-export",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
