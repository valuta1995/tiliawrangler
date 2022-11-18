#  Copyright (c) 2022. Made available per the AGPLv3.0.
#  For closed source distribution options please contact the author.
#  https://www.gnu.org/licenses/agpl-3.0.txt

import csv
import shutil
import sys
import xml.etree.ElementTree as Et
from pathlib import Path
from typing import List, Dict, Any, TextIO

INPUT_FOLDER = './input'
OUTPUT_FOLDER = './output'

SKIP_TILIA_ROWS = [2]
SKIP_TILIA_COLUMNS = [1, 3, 4, 5, 6]  # A C D E F

TILIA_START_DATA = (8, 3)  # Start at H3

CRUNCH_MODE = 1
REPORT_MODE = 2

# Use crunch mode to produce one csv data set for each TLX file.
# Use report mode to produce a folder with csv and metadata for each TLX file.
# MODE = CRUNCH_MODE
MODE = REPORT_MODE

GEOCHRONOLOGY_FIELDS = [
    "Method", "AgeUnits", "Depth", "Thickness", "LabNumber", "Age", "ErrorOlder", "ErrorYounger", "Sigma", "StdDev",
    "GreaterThan", "Parameters", "MaterialDated", "PublicationsText", "Publications"
]
GEOCHRONOLOGY_PARAMETERS = "Parameters"
GEOCHRONOLOGY_PUBLICATIONS = "Publications"

LITHOLOGY_FIELDS = [
    "DepthTop", "DepthBottom", "Description",
]

PUBLICATION_FIELDS = [
    "PublicationType", "NeotomaID", "PublicationYear", "Citation", "Authors", "Title", "SeriesNumber", "Publisher",
    "City", "Country"
]

PUBLICATION_AUTHORS = "Authors"


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


def parse_cell(tilia_row_cell) -> float | int | str | None:
    tilia_cell_text = tilia_row_cell.find('text')
    tilia_cell_value = tilia_row_cell.find('value')
    if tilia_cell_text is not None:
        # print("\"%s\", " % tilia_cell_text.text, end='')
        return tilia_cell_text.text
    elif tilia_cell_value is not None:
        numerical_value = float(tilia_cell_value.text)
        if numerical_value.is_integer():
            # print("%d, " % int(numerical_value), end='')
            return int(numerical_value)
        else:
            # print("%f, " % numerical_value, end='')
            return numerical_value
    else:
        eprint("Unknown Cell")
        return None


def parse_row(tilia_col, csv_row_id, first_col_length) -> List[float | int | str | None]:
    result_row = []
    previous_column = 0
    for tilia_row_cell in tilia_col:

        current_column = int(tilia_row_cell.get('row'))
        while current_column > (previous_column + 1):
            previous_column += 1
            # eprint("Column %d was skipped" % previous_column)
            if previous_column not in SKIP_TILIA_ROWS:
                if csv_row_id >= TILIA_START_DATA[0] and previous_column >= TILIA_START_DATA[1]:
                    result_row.append(0)
                else:
                    result_row.append('')
        previous_column = current_column

        if current_column not in SKIP_TILIA_ROWS:
            cell_content = parse_cell(tilia_row_cell)
            result_row.append(cell_content)

    while previous_column < first_col_length:
        previous_column += 1
        # eprint("Column %d was skipped" % previous_column)
        if previous_column not in SKIP_TILIA_ROWS:
            if csv_row_id >= TILIA_START_DATA[0] and previous_column >= TILIA_START_DATA[1]:
                result_row.append(0)
            else:
                result_row.append('')

    return result_row


def calculate_col_length(tilia_col):
    max_row_id = -1
    for cell in tilia_col:
        row_id = int(cell.get('row'))
        if row_id > max_row_id:
            max_row_id = row_id
    return max_row_id


def parse_sheet(data, csv_writer):
    previous_row = 0
    first_col_length = None
    for tilia_col in data:
        if first_col_length is None:
            first_col_length = calculate_col_length(tilia_col)

        current_row = int(tilia_col.get('ID'))
        while current_row > (previous_row + 1):
            previous_row += 1
            # eprint("Row %d was skipped" % previous_row)
            # print('')
        previous_row = current_row

        if current_row not in SKIP_TILIA_COLUMNS:
            result_row = parse_row(tilia_col, current_row, first_col_length)
            csv_writer.writerow(result_row)


def translate_tlx_to_csv(tlx_root: Et.Element, csv_path: Path):
    sheet: Et.Element = tlx_root.find('SpreadSheetBook')
    data: Et.Element = sheet.find('SpreadSheet')

    with open(csv_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file, dialect='excel')
        parse_sheet(data, csv_writer)


def extract_tlx_to_indexed_list(tlx_element: Et.Element) -> Dict[int, Any]:
    result = {}
    for entry in tlx_element:
        index = int(entry.get('ID'))
        result[index] = entry
    return result


def extract_site_txt(tlx_root: Et.Element, output_path: Path):
    site: Et.Element = tlx_root.find('Site')

    with open(output_path, 'w', newline='') as txt_file:
        txt_file.write("%6fN, %6fE\n" % (
            float(site.find('LatNorth').text),
            float(site.find('LongEast').text)
        ))
        txt_file.write("%dm relative to sea level.\n" % float(site.find('Altitude').text))
        txt_file.write("%s, %s.\n" % (
            site.find('Country').text,
            site.find('State').text
        ))
        txt_file.write(site.find('SiteDescription').text)


def get_author_string(contact):
    name = contact.find('ShortContactName')
    email = contact.find('Email')
    neotoma_id = contact.find('NeotomaContactID')
    if email is None:
        author_string = "%s (Neotoma: %s)" % (name.text, neotoma_id.text,)
    else:
        author_string = "%s (%s) (Neotoma: %s)" % (name.text, email.text, neotoma_id.text,)
    return author_string


def write_contact(txt_file: TextIO, contact: Et.Element):
    author_string = get_author_string(contact)
    txt_file.write(author_string)

    txt_file.write("\n")


def extract_collection_unit_txt(tlx_root: Et.Element, output_path: Path, contacts):
    site: Et.Element = tlx_root.find('CollectionUnit')

    with open(output_path, 'w', newline='') as txt_file:
        txt_file.write("%6fN, %6fE\n" % (
            float(site.find('GPSLat').text),
            float(site.find('GPSLong').text)
        ))
        txt_file.write("%dm relative to sea level.\n" % float(site.find('GPSAltitude').text))

        txt_file.write("Handle:    %s\n" % site.find('Handle').text)
        txt_file.write("Name:      %s\n" % site.find('CollectionName').text)
        txt_file.write("Type:      %s\n" % site.find('CollectionType').text)
        txt_file.write("Device:    %s\n" % site.find('CollectionDevice').text)
        txt_file.write("Substrate: %s\n" % site.find('Substrate').text)
        txt_file.write("Depositional Environment: %s\n" % site.find('DepositionalEnvironment').text)

        for collector in site.find("Collectors"):
            txt_file.write("Collector: ")
            write_contact(txt_file, contacts[int(collector.get('ID'))])


def extract_geochronology_csv(tlx_root: Et.Element, output_path: Path, publication_db):
    geochronolgy_dataset = tlx_root.find('GeochronDataset')
    geochronolgy = geochronolgy_dataset.find('Geochronology')

    analysis_unit = geochronolgy.get('AnalysisUnitID')
    if analysis_unit != "Depth":
        eprint("Potentially unsupported analysis unit: %s\n" % analysis_unit)

    with open(output_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file, dialect='excel')
        csv_writer.writerow(GEOCHRONOLOGY_FIELDS)

        for sample in geochronolgy:
            row = [sample.find(x).text for x in GEOCHRONOLOGY_FIELDS]

            parameters = [
                "%s: %s" % (x.find('Name').text, x.find('Value').text)
                for x in sample.find(GEOCHRONOLOGY_PARAMETERS)
            ]
            row[GEOCHRONOLOGY_FIELDS.index(GEOCHRONOLOGY_PARAMETERS)] = ", ".join(parameters)

            publication_elements: [Et.Element] = [
                publication_db[int(x.get("ID"))]
                for x in sample.find(GEOCHRONOLOGY_PUBLICATIONS)
            ]
            publication_strings = [
                # '%s (%s), ' % (el.find('Citation').text, el.find('NeotomaID').text)
                el.find('NeotomaID').text
                for el in publication_elements
            ]
            row[GEOCHRONOLOGY_FIELDS.index(GEOCHRONOLOGY_PUBLICATIONS)] = ", ".join(publication_strings)

            csv_writer.writerow(row)


def extract_lithology_csv(tlx_root: Et.Element, output_path: Path):
    lithology = tlx_root.find('Lithology')

    with open(output_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file, dialect='excel')
        csv_writer.writerow(LITHOLOGY_FIELDS)

        for lithology_unit in lithology:
            csv_writer.writerow([lithology_unit.find(field).text for field in LITHOLOGY_FIELDS])


def extract_publications_csv(publications: [Et.Element], output_path: Path, contacts):
    with open(output_path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file, dialect='excel')
        csv_writer.writerow(PUBLICATION_FIELDS)

        publication: Et.Element
        for publication in publications.values():
            row = [publication.find(x).text for x in PUBLICATION_FIELDS]

            authors = [
                get_author_string(contacts[int(x.find('Contact').get('ID'))])
                for x in publication.find(PUBLICATION_AUTHORS)
            ]
            row[PUBLICATION_FIELDS.index(PUBLICATION_AUTHORS)] = ", ".join(authors)
            csv_writer.writerow(row)

    pass


def translate_tlx_to_report(tlx_root: Et.Element, report_folder_path: Path):
    base_name_path = report_folder_path.joinpath(report_folder_path.name)
    translate_tlx_to_csv(tlx_root, base_name_path.with_suffix('.csv'))

    contacts = extract_tlx_to_indexed_list(tlx_root.find("Contacts"))
    publications = extract_tlx_to_indexed_list(tlx_root.find("Publications"))

    extract_site_txt(tlx_root, report_folder_path.joinpath('site.txt'))
    extract_collection_unit_txt(
        tlx_root,
        report_folder_path.joinpath('collection_unit.txt'),
        contacts
    )
    extract_geochronology_csv(tlx_root, report_folder_path.joinpath('geochronology.csv'), publications)
    extract_lithology_csv(tlx_root, report_folder_path.joinpath('lithology.csv'))
    extract_publications_csv(publications, report_folder_path.joinpath('publications.csv'), contacts)
    pass


def main():
    input_folder_path = Path(INPUT_FOLDER)
    for potential_tlx_file in input_folder_path.glob('*.tlx'):
        if not potential_tlx_file.is_file():
            continue

        root: Et.Element = Et.parse(potential_tlx_file).getroot()

        if MODE == CRUNCH_MODE:
            csv_output_path = Path(OUTPUT_FOLDER).joinpath(potential_tlx_file.with_suffix('.csv').name)
            translate_tlx_to_csv(root, csv_output_path)

        elif MODE == REPORT_MODE:
            report_output_folder_path = Path(OUTPUT_FOLDER).joinpath(potential_tlx_file.with_suffix('').name)

            if report_output_folder_path.exists():
                shutil.rmtree(report_output_folder_path.absolute())
            report_output_folder_path.mkdir()

            translate_tlx_to_report(root, report_output_folder_path)


if __name__ == '__main__':
    main()
