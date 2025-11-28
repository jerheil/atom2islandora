# atom2islandora
Version 1.0

## Purpose

Script to export and transform descriptions from AtoM into metadata for Islandora 8 ingests

## Installation

### Install Python

Install Python. See [Python downloads](https://www.python.org/downloads/) for instructions.

### Install ExifTool

Install ExifTool. See [ExifTool](https://exiftool.org/) for instructions.

### Install ffmpeg (optional)

Install ffmpeg. See [ffmpeg](https://www.ffmpeg.org/) for instructions.

### Download atom2islandora.py

Create a folder on your computer (Desktop or wherever). Save the [atom2islandora.py](https://github.com/jerheil/atom2islandora/blob/main/atom2islandora.py) file to a folder on your computer. FOR WINDOWS USERS: Save the [a2i.bat](https://github.com/jerheil/atom2islandora/blob/main/a2i.bat) file to the same folder

### Download convert_wav_to_mp3.bat

FOR WINDOWS USERS: Save the [convert_wav_to_mp3.bat](https://github.com/jerheil/atom2islandora/blob/main/convert_wav_to_mp3.bat) to an accessible folder on your computer.

## Use

### Create Access Derivatives (FOR WINDOWS USERS)

#### Audio
1) Ensure ffmpeg is installed and accessible through the path to access the program
2) Run convert_wav_to_mp3.bat 

### Archives

#### Export from AtoM

Use the Clipboard feature in AtoM to select the Items or Files to be ingested into Islandora. As of Version 0.1, this is set to run for a single parent collection in Islandora, although the resulting file (product.csv) can be edited before ingest.
1) Load the Clipboard by Clipboard ID in AtoM
2) Click Export, and set format as CSV
3) Click refresh the page, then click Download once the link appears. Save it to the folder containing atom2islandora.py

#### Run atom2islandora

1) FOR WINDOWS USERS: Double-click on a2i.bat. For Mac or Linux users, run atom2islandora.py through the command line
```
   py atom2islandora.py
```
2) Follow the prompt to input the parent id in Islandora.
3) The program will produce four files:
    a) source1.csv - this is condensed from the exported file from AtoM and can be deleted or kept to update AtoM after ingest to Islandora
    b) source2.csv - this is the result of the exiftool scan and can be deleted
    c) product.csv - this is the main product for ingest into Islandora
    d) error.txt - this reports on any issues you may need to address in product.csv before ingest
4) Rename product.csv to a name you'd like to use for the ingest.

### Maps

#### Prepare metadata

1) Create a folder with the following files:
   a) source1.csv
   b) source2.csv

#### Run atom2islandora
1) FOR WINDOWS USERS: Double-click on a2i.bat. For Mac or Linux users, run atom2islandora.py through the command line
```
   py atom2islandora.py
```
2) Follow the prompt to select the folder with the source.csv files
3) Enter the desired file name for the product.csv (what will be ingested into QULDC)
4) Enter the value for the parent collection (member_of_existing_entity_id)
5) Check the mapping-report.txt and correct any issues the the product.csv
6) Enter your choice for whether to delete the working files.
