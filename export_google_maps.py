import json
from openpyxl import Workbook
from pathlib import Path

wb = Workbook()


json_file = Path(input('Enter full path to json file (or drag the file here):\n').strip('\'\"'))

while not json_file.is_file():
    json_file = Path(input('Invalid path, please re-enter:\n').strip('\'\"'))

dest_filename = json_file.parent / 'exported_places.xlsx'

JSON = json.loads(open(json_file, encoding='utf8').read())

ws1 = wb.active
ws1.title = 'Starred Places'

FIELDS = {
    1: 'S/N',
    2: 'Title',
    3: 'Address',
    4: 'Business Name',
    5: 'Country Code',
    6: 'Updated',
    7: 'Latitude',
    8: 'Longitude',
    9: 'Google Maps URL'
}

ws1.append(FIELDS)
total = 0
for number, place in enumerate(JSON['features']):
    ws1.append({
        1: number,
        2: place['properties']['Title'],
        3: place['properties']['Location'].get('Address', ''),
        4: place['properties']['Location'].get('Business Name', ''),
        5: place['properties']['Location'].get('Country Code', ''),
        6: place['properties']['Updated'],
        7: place['properties']['Location']['Geo Coordinates']['Latitude'],
        8: place['properties']['Location']['Geo Coordinates']['Longitude'],
        9: place['properties']['Google Maps URL']
    })
    total = number

print(f'Processed {total} records')
wb.save(dest_filename)
print(f'\nSaved to {dest_filename}')
