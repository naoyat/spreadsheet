# spreadsheet.py

To log something on Google spreadsheet...

## Requirements

```
python==2.7.9
gdata==2.0.18
oauth2client==1.5.2
```

## Preparation

- open https://console.developers.google.com/
- activate Google Drive API
- make a new service account and get an auth key (`*.p12`, not `*.json`)
- note the email address of that account

- create a new spreadsheet at Google Spreadsheet
- grant its read/write permission to that email address
- you can find the `spreadsheet_key` of that spreadsheet in its URL:

```
https://docs.google.com/spreadsheets/d/<<SPREADSHEET_KEY_COMES_HERE>>/edit#gid=0
```

## Configuration

```spreadsheet.conf
{
  "oauth2_key_file": "XXXXXXXX.p12",
  "client_email": "999999999999-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@developer.gserviceaccount.com",
  "default_spreadsheet_key": "XXXXXXXXXXXXXXX",
  "default_sheet_id": "od6"
}
```

- oauth2_key_file : path to `*.p12` file
- client_email : the email address of your service account
- default_spreadsheet_key : the spreadsheet you want to write in.
- default_worksheet_id (optional) : if you don't know, please say 'od6'

## License

Licensed under the Apache License, Version 2.0

http://www.apache.org/licenses/LICENSE-2.0

## Author

naoya_t
