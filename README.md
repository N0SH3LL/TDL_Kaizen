TDL Kaizen
Overview

This Python project is designed to assist with audit preparation for controls from Security Configuration Checklists (SCC's) generated from STIGs (Security Technical Implementation Guides). The tool automates some of the document gathering that was extremely time intensive. 
Features

Automated Data Extraction: Pulls relevant information from STIG-generated checklists. The assumption is there are two or more tabs: a title page and the checklist. These checklists can be generated from STIG Viewer, but should include columns labelled "STIG ID", "Configuration", "Exception ID", "Compliance Method", and "Documentation". It can then specified directories for these documents, and gather information from them. 
Data Processing: Pulls this information into "progress.json" creating dictionaries for checklists, exceptions, documents, and program settings. 
Reporting: Writes information from this in various ways, including a markdown file checklist. 

Prerequisites

Python 3.x
Required Python libraries (specified in requirements.txt)

Usage

Place your STIG-generated checklists in the input directory.
Install required dependencies. 
Run GUI.py to start the program. 

Configuration

You can configure various settings by modifying the config.json file. This includes paths to input/output directories, processing options, and report formats.

For any questions or suggestions, please open an issue in the repository.
