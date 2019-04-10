# kml-converter

Powershell script to convert kml files to csv format.  Intended to prep data for import into MySQL

## About

I had used a pykismet script to convert a bunch of kismet files into .kml files.  This was intended for visuualization purposes in google earth.  Since .kml files are handled quite nicely with Powershell, I decided to create a script to parse the data and prep it for import into MySQL.

## Usage

Place the script in a directory where all of your .kml files are located.  The script will parse the directory for files, convert them, and move the output to a folder named after the .kml file

## Loggin'

Run the script from the command line like so:

```PowerShell
.\kml-converter.ps1 *> kenny.log
```

Note the asterisk redirecting _all_ streams to the kenny.log file.