"""Common utilities for UMM Tools"""

import logging


def setup_logging(verbose):
    root_logger = logging.getLogger()
    log_level = logging.INFO
    if verbose:
        log_level=logging.DEBUG
    root_logger.setLevel(log_level)


def get_klassenkuerzel(kategorie):
    if kategorie == 'WOM-10K*':
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
        'U17W': '5-K',
        'U18M*': '10-K',
        'U18W*': '7-K',
        'U20M': 'MK',
        'U20M*': '10-K',
        'U20W': '5-K',
        'U20W*': '7-K',
        'WOM': '5-K',
        'WOM*': '7-K',
        'WOM-10K*': '10-K',
    }
    return kategorie_bewerbskuerzel_mapping[kategorie]
