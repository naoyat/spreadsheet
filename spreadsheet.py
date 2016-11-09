#!/usr/bin/env python
# -*- coding: utf-8 -*-
# -------------------------------------------------------------------------
#
# Copyright (C) 2014-2016 @naoya_t
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# -------------------------------------------------------------------------
#
# spreadsheet : "book" in excel
# worksheet : "(work)sheet" in excel
#
import datetime
import re
import json

import oauth2client
from gdata.spreadsheets.client import SpreadsheetsClient, CellQuery
from gdata.spreadsheets.data import ListEntry
from gdata.gauth import OAuth2TokenFromCredentials

if oauth2client.__version__ == '1.5.2':
    from oauth2client.client import SignedJwtAssertionCredentials
    _use_old_oauth2client = True
else:
    from oauth2client.service_account import ServiceAccountCredentials
    _use_old_oauth2client = False


CONFIG_PATH = 'spreadsheet.conf'


def _strip_worksheet_entry(worksheet_entry):
    worksheet_id = worksheet_entry.id.text.split('/')[-1]
    return {
        'id': worksheet_id,
        'title': worksheet_entry.title.text,
        'col_count': int(worksheet_entry.col_count.text),
        'row_count': int(worksheet_entry.row_count.text),
        }


def _convert(value):
    if isinstance(value, unicode):
        return value.encode('utf-8')

    if value is None:
        return '-'
    elif isinstance(value, bool):
        return ('NO', 'YES')[int(value)]
    elif isinstance(value, int):
        return '%d' % value
    elif isinstance(value, float):
        return '%g' % value
    elif isinstance(value, datetime.datetime):
        return value.strftime('%Y/%m/%d %H:%M:%S')
    elif isinstance(value, str):
        if re.match(r'[.:0-9]+', value):
            value = "'" + value
        return value
    else:
        return str(value)


class Spreadsheet:
    def __init__(self, *args, **kwargs):
        if len(kwargs) == 0:
            if len(args) == 1:
                config_path = args[0]
            else:
                config_path = CONFIG_PATH
            print 'loading config from %s...' % config_path
            with open(config_path, 'r') as fp:
                self.config = json.load(fp)
        else:
            self.config = kwargs

        if 'oauth2_key_file' not in self.config:
            raise Exception('oauth2_key_file is not given')

        if 'client_email' not in self.config:
            raise Exception('client_email is not given')

        self.default_spreadsheet_key = self.config.get('default_spreadsheet_key', None)
        self.default_worksheet_id = self.config.get('default_worksheet_id', None)

        if _use_old_oauth2client:
            # use .pem
            key_file_path = self.config['oauth2_key_file']
            if key_file_path[-4:] == '.p12':
                key_file_path = key_file_path.replace('.p12', '.pem')
            with open(key_file_path, 'rb') as f:
                private_key = f.read()
                self.credentials = SignedJwtAssertionCredentials(self.config['client_email'],
                                                                 private_key,
                                                                 scope=["https://spreadsheets.google.com/feeds"])
        else:
            # use .p12
            self.credentials = ServiceAccountCredentials.from_p12_keyfile(self.config['client_email'],
                                                                         self.config['oauth2_key_file'],
                                                                         scopes=['https://spreadsheets.google.com/feeds'])
        self.client = SpreadsheetsClient()
        self.auth_token = OAuth2TokenFromCredentials(self.credentials)
        self.auth_token.authorize(self.client)

    def get_spreadsheet_infos(self):
        infos = []
        spreadsheets_feed = self.client.get_spreadsheets()
        for spreadsheet in spreadsheets_feed.entry:
            spreadsheet_key = spreadsheet.id.text.split('/')[-1]
            infos.append({
                'id': spreadsheet_key,
                'title': spreadsheet.title.text,
                'sheets': self.get_worksheet_infos(spreadsheet_key)
                })
        return infos

    def get_worksheet_infos(self, spreadsheet_key=None):
        spreadsheet_key = spreadsheet_key or self.default_spreadsheet_key
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        infos = []
        worksheets_feed = self.client.get_worksheets(spreadsheet_key)
        for worksheet_entry in worksheets_feed.entry:
            worksheet_id = worksheet_entry.id.text.split('/')[-1]
            worksheet_info = self.get_worksheet_info(spreadsheet_key, worksheet_id)
            infos.append(worksheet_info)
        return infos

    def get_worksheet_info(self, spreadsheet_key=None, worksheet_id=None):
        try:
            worksheet_entry = self.client.get_worksheet(spreadsheet_key, worksheet_id)
            if not worksheet_entry:
                return None
            return _strip_worksheet_entry(worksheet_entry)
        except:
            return None

    def update_entry(self, entry):
        if entry:
            self.client.update(entry)

    def iter_rows(self, spreadsheet_key=None, worksheet_id=None):
        spreadsheet_key = spreadsheet_key or self.default_spreadsheet_key
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        worksheet_id = worksheet_id or self.default_worksheet_id
        if worksheet_id is None:
            raise Exception('sheet_id is not given')

        list_feed = self.client.get_list_feed(spreadsheet_key, worksheet_id)
        for entry in list_feed.entry:
            yield entry.to_dict()

    def add_header(self, spreadsheet_key=None, worksheet_id=None, header=[]):
        spreadsheet_key = spreadsheet_key or self.default_spreadsheet_key
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        worksheet_id = worksheet_id or self.default_worksheet_id
        if worksheet_id is None:
            raise Exception('sheet_id is not given')

        cell_query = CellQuery(min_row=1, max_row=1,
                               min_col=1, max_col=len(header), return_empty=True)
        cells = self.client.GetCells(spreadsheet_key, worksheet_id, q=cell_query)
        for i, name in enumerate(header):
            cell_entry = cells.entry[i]
            cell_entry.cell.input_value = name
            self.client.update(cell_entry) # This is the

    # add_list_entry
    def add_row(self, spreadsheet_key=None, worksheet_id=None, values={}):
        spreadsheet_key = spreadsheet_key or self.default_spreadsheet_key
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        worksheet_id = worksheet_id or self.default_worksheet_id
        if worksheet_id is None:
            raise Exception('sheet_id is not given')

        if len(values) == 0:
            return

        list_entry = ListEntry()
        list_entry.from_dict({key: _convert(value)
                              for key, value in values.iteritems()})

        self.client.add_list_entry(list_entry, spreadsheet_key, worksheet_id)

    def add_worksheet(self, spreadsheet_key=None, title=None, rows=1, cols=1): # worksheet_id, values):
        spreadsheet_key = spreadsheet_key or self.default_spreadsheet_key
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        if not title:
            raise Exception('title is not given')

        worksheet_entry = self.client.add_worksheet(spreadsheet_key, title, rows, cols)
        return _strip_worksheet_entry(worksheet_entry)

    def delete_worksheet(self, spreadsheet_key, worksheet_id):
        if spreadsheet_key is None:
            raise Exception('spreadsheet_key is not given')

        if worksheet_id is None:
            raise Exception('worksheet_id is not given')

        worksheet_entry = self.client.get_worksheet(spreadsheet_key, worksheet_id)
        self.client.delete(worksheet_entry)


if __name__ == '__main__':
    ss = Spreadsheet()

    for spreadsheet_info in ss.get_spreadsheet_infos():
        print '%(id)s : %(title)s' % spreadsheet_info
        for sheet_info in spreadsheet_info['sheets']:
            print '  %(id)-7s : %(title)s (%(row_count)d x %(col_count)d)' % sheet_info
        print
