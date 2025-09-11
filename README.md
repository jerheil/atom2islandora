# atom2islandora
Version 0.1

## Purpose

Script to export and transform descriptions from AtoM into metadata for Islandora 8 ingests

## Installation

### Install Python

Install Python. See [Python downloads](https://www.python.org/downloads/) for instructions.

### Install ExifTool

Install ExifTool. See [ExifTool](https://exiftool.org/) for instructions.

### Download atom2islandora.py

Create a folder on your computer (Desktop or wherever). Save the [atom2islandora.py]() file to a folder on your computer. FOR WINDOWS USERS: Save the [a2i.bat]() file to the same folder

## Use

### Export from AtoM

Use the Clipboard feature in AtoM to select the Items or Files to be ingested into Islandora. As of Version 0.1, this is set to run for a single parent collection in Islandora, although the resulting file (product.csv) can be edited before ingest.
1) Load the Clipboard by Clipboard ID in AtoM
2) Click Export, and set format as CSV
3) Click refresh the page, then click Download once the link appears. Save it to the folder containing atom2islandora.py

### Run ExifTool on digital objects to be ingested

4) Run ExifTool - swap out "folder" in the pathway for the folder containing atom2islandora.py
```
   exiftool.exe -csv -r -SourceFile -Title -FileName -FileCreateDate -PageCount -FileTypeExtension -MIMEType -LayerCount * > "folder\source2.csv"
```

### Run atom2islandora

5) FOR WINDOWS USERS: Double-click on a2i.bat. For Mac or Linux users, run atom2islandora.py through the command line
```
   py atom2islandora.py
```
7) Follow the prompt to input the parent id in Islandora.
8) The program with produce three files:
    a) source1.csv - this is condensed from the exported file from AtoM and can be deleted
    b) product.csv - this is the main product for ingest into Islandora
    c) error.txt - this reports on any issues you may need to address in product.csv before ingest
9) Rename product.csv to a name you'd like to use for the ingest.
