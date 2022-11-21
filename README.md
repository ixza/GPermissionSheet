# GPermissionSheet
Python script to autodetect permissions of users on Google sheets, then fill in permisson data on a new sheet.

Written and tested on Python 3.11.0

![ezgif-5-196d6c023c](https://user-images.githubusercontent.com/116339318/202987726-6a9a50e8-87fb-46b5-b1bd-c4ec45b63075.png)

You will need:

- Service account with permissions to the Google sheet you want to read
* A blank Google sheet for the program to output to


## Usage

Put your service account .json key into the same directory as config.py.

Put in your sheet ID (found in URL) in config.py. SPREADSHEET_ID is the URL of the sheet you want to read permissions of, and WRITE_SPREADSHEET_ID is the sheet
for the program to write to (will be overwritten).

Run main.py

## Features

- Fully formatted and readable output sheet
* Works on any sheet as long as service account has permission to it

## TODO
Clean up spaghetti code 💀
