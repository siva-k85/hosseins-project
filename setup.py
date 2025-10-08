"""
Setup script for eMMA Excel Updater
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="emma-scraper",
    version="2.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Enhanced Maryland eMMA procurement scraper and Excel updater",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/emma-scraper",
    packages=find_packages(exclude=["tests", "tests.*"]),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.9",
    install_requires=[
        "requests>=2.28.0",
        "beautifulsoup4>=4.11.0",
        "pandas>=1.5.0",
        "openpyxl>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "flake8>=5.0.0",
            "black>=22.0.0",
            "mypy>=1.0.0",
            "pre-commit>=3.0.0",
        ],
        "notifications": [
            "slack-sdk>=3.0.0",
            "pymsteams>=0.2.0",
            "sendgrid>=6.0.0",
        ]
    },
    entry_points={
        "console_scripts": [
            "emma-scraper=emma_scraper_enhanced:main",
        ],
    },
    include_package_data=True,
    zip_safe=False,
)