# kml-converter
Powershell script to convert kml files to csv format.  Intended to prep data for import into MySQL

## About
I had used a pykismet script to convert a bunch of kismet files into .kml files.  This was intended for visuualization purposes in google earth.  Since .kml files are handled quite nicely with Powershell, I decided to create a script to parse the data and prep it for import into MySQL.

## Usage
Specify the path to the kml file you wish to convert.  Include .kml in the filename for this input.  For output, you do not need to append .csv to the desired filename as this will be appended for you.  Manipulate the objects created in the file to export differnt .csv dumps as needed.
