# plz
append GPS data to spreadsheets with postal codes

This is a small Python script combining the following two functionalities:
* Download Open Street Map data via the Overpass API to build a database of the geographical coordinates of postal code areas in Austria and save this database as a JSON file
* Given a spreadsheet (e.g. in Excel format) containing records with postal codes, add geo information (latitude and longitude) as new columns to that spreadsheet.

