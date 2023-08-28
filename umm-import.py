import csv
import logging
import openpyxl


def setup_logging(verbose):
    root_logger = logging.getLogger()
    log_level = logging.INFO
    if verbose:
        log_level=logging.DEBUG
    root_logger.setLevel(log_level)


def get_klassenkuerzel(kategorie):
    if kategorie == 'WOM-10K':
        return 'WOM'
    return kategorie.replace('*', '')

def get_bewerbskuerzel(kategorie):
    kategorie_bewerbskuerzel_mapping = {
        'MAN': 'MK',
        'MAN*': '10-K',
        'U12M': 'MK', 
        'U12W': 'MK',
        'U14M': 'MK',
        'U14W': 'MK',
        'U16M*': '6-K',
        'U16W*': '5-K',
        'U17M': 'MK',
        'U17W': 'MK',
        'U18M*': '10-K',
        'U18W*': '7-K',
        'U20M': 'MK',
        'U20M*': '10-K',
        'U20W': '5-K',
        'U20W*': '7-K',
        'WOM': '5-K',
        'WOM*': '7-K',
        'WOM-10K': '10-K',
    }
    return kategorie_bewerbskuerzel_mapping[kategorie]

def amend_klassenkuerzel_and_bewerbskuerzel(ws):
    logging.debug("amend klassenkuerzel and bewerbskuerzel")
    cell = ws['K1']
    cell.value = 'Klassenkürzel'
    cell = ws['L1']
    cell.value = 'Bewerbskürzel'
    for row_index, cell in enumerate(ws['G']):
        if row_index == 0:
            assert cell.value == 'Kategorie'
        elif cell.value is not None:
            kategorie = cell.value
            klassenkuerzel = get_klassenkuerzel(kategorie)
            bewerbskuerzel = get_bewerbskuerzel(kategorie)
            logging.info('%s: %r => %r, %r' % (row_index, kategorie, klassenkuerzel, bewerbskuerzel))
            column_row_code = 'K{}'.format(cell.row)
            target_cell = ws[column_row_code]
            target_cell.value = klassenkuerzel
            column_row_code = 'L{}'.format(cell.row)
            target_cell = ws[column_row_code]
            target_cell.value = bewerbskuerzel

def write_csv(ws, filename):
    logging.debug('writing csv')
    with open(filename, 'w', newline="", encoding='latin-1') as file_handle:
        csv_writer = csv.writer(file_handle, delimiter=';', quotechar='"')
        for row in ws.iter_rows():
            csv_writer.writerow([cell.value for cell in row])

def main(args):
    setup_logging(args.verbose)
    logging.debug("main(%s)" % args)
    wb = openpyxl.load_workbook(args.subscriptions)
    ws = wb.active
    amend_klassenkuerzel_and_bewerbskuerzel(ws)
    csv_filename = args.subscriptions.rstrip('xlsx') + 'csv'
    write_csv(ws, csv_filename)
    logging.debug("done")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='UMM Import Tool')
    parser.add_argument('-v', '--verbose', action="store_true", help="be verbose")
    parser.add_argument('subscriptions', help='XLSX file')
    args = parser.parse_args()

    main(args)
