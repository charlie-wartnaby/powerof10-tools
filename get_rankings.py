# First written by (c) Charlie Wartnaby 2023
# See https://github.com/charlie-wartnaby/powerof10-tools">https://github.com/charlie-wartnaby/powerof10-tools


import argparse
import datetime
import openpyxl
import os
import pandas
import pickle
import re
import requests   # Not included by default, use pip install to add
import sys

if sys.version_info.major < 3:
    print('This script needs Python 3')
    sys.exit(1)
if sys.version_info.minor < 6:
    print('WARNING: this script assumes ordered dicts (Python 3.6+) when it builds cache keys, bad things may happen here')


class Performance():
    def __init__(self, event, score, category, gender, original_special, decimal_places, athlete_name, athlete_url='', date='',
                       fixture_name='', fixture_url='', source='', wava=0.0, age=0):
        self.event = event
        self.score = score # could be time in sec, distance in m or multievent points
        self.category = category # e.g. U20 or ALL
        self.gender = gender # W or M
        self.original_special = original_special # for wind-assisted detail etc from club records
        self.decimal_places = decimal_places # so we can use original precision which may imply electronic timing etc
        self.athlete_name = athlete_name
        self.athlete_url = athlete_url
        self.date = date
        self.fixture_name = fixture_name
        self.fixture_url = fixture_url
        self.source = source
        # Added later to support marathon WAVA list, cached performances may not have:
        self.wava = wava
        self.age = age

class HtmlBlock():
    def __init__(self, tag=''):
        self.tag = tag
        self.inner_text = ''
        self.attribs = {}

record = {} # dict of age groups, each dict of events, each dict of genders, then ordered list of performance lists (allowing for ties)
wava = {} # dict of events, each dict of years and 0 for all years, then ordered list of performance lists (to allow for ties)
max_records_all = 10 # max number of records for each event/gender, including all age groups
max_regords_age_group = 3 # Similarly per age group
max_wavas_all = 20  # All-time WAVA list
max_wavas_year = 10 # WAVA list for specific year
wava_athlete_ids_done = {}

powerof10_root_url = 'https://thepowerof10.info'
runbritain_root_url = 'https://www.runbritainrankings.com'

performance_count = {'Po10'       : 0,
                     'Runbritain' : 0,
                     'File(s)'    : 0}

wava_events = ['Mar', 'HM', '10K']  # C&C trophy category but could do other events

# Smaller time is good for runs, bigger distance/score better for jumps/throws/multievents;
# some events should be in sec (1 number), some in min:sec (2 numbers), some h:m:s (3 numbers):
#                event, small-is-good, :-numbers, runbritain,  Track/Field/Road/Multievent
known_events = [
                ('1M',         True,        2,        True,     'R'     ),
                ('2M',         True,        2,        True,     'R'     ),
                ('5K',         True,        2,        True,     'R'     ),
                ('parkrun',    True,        2,        True,     'R'     ),
                ('4M',         True,        2,        True,     'R'     ),
                ('5M',         True,        2,        True,     'R'     ),
                ('10K',        True,        2,        True,     'R'     ),
                ('10M',        True,        2,        True,     'R'     ),
                ('HM',         True,        2,        True,     'R'     ),
                ('Mar',        True,        3,        True,     'R'     ),
                ('50K',        True,        3,        True,     'R'     ),
                ('100K',       True,        3,        True,     'R'     ),
                ('60' ,        True,        1,        True,     'T'     ),
                ('80' ,        True,        1,        True,     'T'     ),
                ('100',        True,        1,        True,     'T'     ),
                ('150',        True,        1,        True,     'T'     ),
                ('200',        True,        1,        True,     'T'     ),
                ('300',        True,        1,        True,     'T'     ),
                ('400',        True,        1,        True,     'T'     ),
                ('600',        True,        1,        True,     'T'     ),
                ('800',        True,        2,        True,     'T'     ),
                ('1500',       True,        2,        True,     'T'     ),
                ('Mile',       True,        2,        True,     'T'     ),
                ('3000',       True,        2,        True,     'T'     ),
                ('5000',       True,        2,        True,     'T'     ),
                ('10000',      True,        2,        True,     'T'     ),
                ('1500SC',     True,        2,        True,     'T'     ),
                ('1500SCW',    True,        2,        True,     'T'     ),
                ('2000SC',     True,        2,        True,     'T'     ),
                ('2000SCW',    True,        2,        True,     'T'     ),
                ('3000SC',     True,        2,        True,     'T'     ),
                ('3000SCW',    True,        2,        True,     'T'     ),
                ('MileW',      True,        2,        True,     'T'     ), # Walks not shown as runbritain dropdowns but are supported
                ('1500W',      True,        2,        True,     'T'     ),
                ('2000W',      True,        2,        True,     'T'     ),
                ('3000W',      True,        2,        True,     'T'     ),
                ('5000W',      True,        2,        True,     'T'     ),
                ('5MW',        True,        2,        True,     'R'     ),
                ('10000W',     True,        2,        True,     'T'     ),
                ('10KW',       True,        2,        True,     'R'     ),
                ('70HU13W',    True,        1,        True,     'T'     ),
                ('75HU13M',    True,        1,        True,     'T'     ),
                ('75HU15W',    True,        1,        True,     'T'     ),
                ('80HU15M',    True,        1,        True,     'T'     ),
                ('80HU17W',    True,        1,        True,     'T'     ),
                ('80HW40',     True,        1,        True,     'T'     ),
                ('80HW50',     True,        1,        True,     'T'     ),
                ('100HW',      True,        1,        True,     'T'     ), # Hurdles not shown on runbritain but do work
                ('100HM50',    True,        1,        True,     'T'     ),
                ('100HU17M',   True,        1,        True,     'T'     ),
                ('110HU20M',   True,        1,        True,     'T'     ),
                ('110H',       True,        1,        True,     'T'     ),
                ('110HM35',    True,        1,        True,     'T'     ),
                ('110HM50',    True,        1,        True,     'T'     ),
                ('300HW',      True,        1,        True,     'T'     ),
                ('400H',       True,        1,        True,     'T'     ),
                ('400HW',      True,        1,        True,     'T'     ),
                ('400HU17M',   True,        1,        True,     'T'     ),
                ('4x100',      True,        1,        True,     'T'     ),
                ('4x400',      True,        2,        True,     'T'     ),
                ('HJ',         False,       1,        True,     'F'     ),
                ('PV',         False,       1,        True,     'F'     ),
                ('LJ',         False,       1,        True,     'F'     ),
                ('TJ',         False,       1,        True,     'F'     ),
                ('SP2.72K',    False,       1,        True,     'F'     ), # Some of these weights only for certain age/gender groups,
                ('SP3K',       False,       1,        True,     'F'     ), # so inefficient to try them for all, but may as well
                ('SP3.25K',    False,       1,        True,     'F'     ),
                ('SP4K',       False,       1,        True,     'F'     ),
                ('SP5K',       False,       1,        True,     'F'     ),
                ('SP6K',       False,       1,        True,     'F'     ),
                ('SP7.26K',    False,       1,        True,     'F'     ),
                ('DT0.75K',    False,       1,        True,     'F'     ),
                ('DT1K',       False,       1,        True,     'F'     ),
                ('DT1.25K',    False,       1,        True,     'F'     ),
                ('DT1.5K',     False,       1,        True,     'F'     ),
                ('DT1.75K',    False,       1,        True,     'F'     ),
                ('DT2K',       False,       1,        True,     'F'     ),
                ('HT3K',       False,       1,        True,     'F'     ),
                ('HT4K',       False,       1,        True,     'F'     ),
                ('HT5K',       False,       1,        True,     'F'     ),
                ('HT6K',       False,       1,        True,     'F'     ),
                ('HT7.26K',    False,       1,        True,     'F'     ),
                ('WT5.45K',    False,       1,        True,     'F'     ),
                ('WT7.26K',    False,       1,        True,     'F'     ),
                ('WT9.08K',    False,       1,        True,     'F'     ),
                ('WT11.34K',   False,       1,        True,     'F'     ),
                ('JT400',      False,       1,        True,     'F'     ),
                ('JT500',      False,       1,        True,     'F'     ),
                ('JT600',      False,       1,        True,     'F'     ),
                ('JT600PRE86', False,       1,        False,    'F'    ), # Invented here for historical records
                ('JT600PRE99', False,       1,        False,    'F'    ), # Invented here for historical records
                ('JT700',      False,       1,        True,     'F'     ),
                ('JT800',      False,       1,        True,     'F'     ),
                ('JT800PRE86', False,       1,        False,    'F'    ), # Invented here for historical records
                ('Minithon',   False,       1,        False,    'M'    ), # from C&C club records but not in Po10
                ('Oct',        False,       1,        False,    'M'    ), # from C&C club records but not in Po10
                ('PenU13W',    False,       1,        True,     'M'     ),
                ('PenU13M',    False,       1,        True,     'M'     ),
                ('PenU15W',    False,       1,        True,     'M'     ),
                ('PenU15M',    False,       1,        True,     'M'     ),
                ('PenU17W',    False,       1,        True,     'M'     ),
                ('PenU17M',    False,       1,        True,     'M'     ),
                ('PenU20M',    False,       1,        True,     'M'     ),
                ('PenW',       False,       1,        True,     'M'     ),
                ('PenIM35',    False,       1,        True,     'M'     ),
                ('PenIM40',    False,       1,        True,     'M'     ),
                ('PenWtM40',   False,       1,        True,     'M'     ),
                ('PenWtM45',   False,       1,        True,     'M'     ),
                ('PenWtM55',   False,       1,        True,     'M'     ),
                ('PenWtW60',   False,       1,        True,     'M'     ),
                ('PenWtM60',   False,       1,        True,     'M'     ),
                ('HepW',       False,       1,        True,     'M'     ),
                ('HepU17W',    False,       1,        True,     'M'     ),
                ('Dec',        False,       1,        True,     'M'     )
 ]

known_events_lookup = {}
for (event, smaller_better, numbers, runbritain, type) in known_events:
    known_events_lookup[event] = (smaller_better, numbers, runbritain, type)

# PowerOf10 age categories (usable on club page)
powerof10_categories = ['ALL', 'U13', 'U15', 'U17', 'U20']

# Runbritain age categories
# TODO I don't understand difference between on-the-day and "season" age
# If both min and max are 0, need to search using category name not age range.
# Otherwise use age range as runbritain skips some results if use category name, oddly.
runbritain_categories = [ # name       min  max years old
                          ('ALL',        0,    0),
                          # ('Disability', 0,    0),     # Runbritain category but always zero results
                          ('U13',        0,    0), # 1,   12), # Using official season age groups for juniors...
                          ('U15',        0,    0), # 13,  14), # ... could make it different for road/parkruns?
                          ('U17',        0,    0), # 15,  16),
                          ('U20',        0,    0), # 17,  19),
                          ('U23',        0,    0), # 20,  22),
                          # Skipping senior as have all-age records
                          ('V35',        35,  39),
                          ('V40',        40,  44),
                          ('V45',        45,  49),
                          ('V50',        50,  54),
                          ('V55',        55,  59),
                          ('V60',        60,  64),
                          ('V65',        65,  69),
                          ('V70',        70,  74),
                          ('V75',        75,  79),
                          ('V80',        80,  84),
                          ('V85',        85,  89),
                          ('V90',        90, 120)
]

runbritain_category_lookup = {}
for (category, min_age, max_age) in runbritain_categories:
    runbritain_category_lookup[category] = (min_age, max_age)



def get_html_content(html_text, html_tag):
    """Extracts instances of text data enclosed by required tag"""

    open_regex = re.compile(r'<' + html_tag + r'(.*?)>', flags=re.DOTALL)
    close_regex = re.compile(r'</' + html_tag + r'>')

    contents = []
    offset = 0
    nesting_depth = 0
    inside_tag_block = False
    block_content_start_idx = -1

    while True:
        open_match = open_regex.search(html_text, pos=offset)
        close_match = close_regex.search(html_text, pos=offset)
        if close_match is None:
            if ((open_match is not None) and (nesting_depth == 0)) or (nesting_depth > 0):
                print(f'WARNING: no match for closing tag "{html_tag}"')
            break
        if open_match is not None and (open_match.start() < close_match.start()):
            if inside_tag_block:
                nesting_depth += 1
            else:
                block_content_start_idx = open_match.end()
                inside_tag_block = True
                attribs_unparsed = open_match.group(1)
                attrib_pairs = attribs_unparsed.split(' ')
                content = HtmlBlock(html_tag)
                for attrib_pair in attrib_pairs:
                    key, _, quoted_value = attrib_pair.partition('=')
                    if not (key or quoted_value):
                        continue
                    unquoted_value = quoted_value.replace('"', '')
                    content.attribs[key] = unquoted_value
            offset = open_match.end()
        else:
            if nesting_depth > 0:
                # keep searching for tag that closes original opening tag
                nesting_depth -= 1
            else:
                content.inner_text = html_text[block_content_start_idx : close_match.start()]
                contents.append(content)
                inside_tag_block = False
            offset = close_match.end()

    return contents


def debold(bold_tagged_string):
    return bold_tagged_string.replace('<b>', '').replace('</b>', '')


def make_numeric_score_from_performance_string(perf):
    # For values like 2:17:23 or 1:28.37 need to split (hours)/mins/sec

    original_special = ''

    perf = perf.replace('pts', '') # for pentathlon/decathlon etc
    perf = perf.replace(' ', '')   # Club records can have random spaces
    perf = perf.replace(';', ':')  # Semicolons sometimes
    perf = perf.strip()

    # TODO handle wind-assisted etc from club records

    # In "6m 26.5s" and "3min 17.76s" treat 'm ' or 'min ' as sexagesimal separator
    # by replacing with ':'
    perf = re.sub(r'([0-9])(m(in)? *)([0-9])', r'\g<1>:\g<4>', perf)

    for char_idx, c in enumerate(perf):
        if (not (c >= '0' and c <= '9') and
                       not c in ['.', ':']):
            original_special = perf # preserve original detail
            perf = perf[:char_idx]
            break

    total_score = 0.0
    multiplier = 1.0
    sexagesimals = perf.split(':')
    sexagesimals.reverse()
    for sexagesmial in sexagesimals:
        total_score += float(sexagesmial) * multiplier
        multiplier *= 60.0
    decimal_split = perf.split('.')
    decimal_places = 0 if len(decimal_split) < 2 else len(decimal_split[1])

    return total_score, decimal_places, original_special


def construct_performance(event, gender, category, perf, name, url, date, fixture_name, fixture_url,
                          source, age_grade='0.0', age=0):
    score, original_dp, original_special = make_numeric_score_from_performance_string(perf)
    wava = float(age_grade)
    perf = Performance(event, score, category, gender, original_special, original_dp, name, url, 
                        date, fixture_name, fixture_url, source, wava=wava, age=age)
    return perf


def process_performance(perf, types, collection_choice, year='ALL'):
    """Add performance to overall and category record tables if appropriate,
      while respecting the max size of those tables"""
    
    known_event = known_events_lookup.get(perf.event)
    if not known_event:
        print(f'Event not in known events: {event}')
        return
    if known_event[3] not in types:
        # e.g. Track event but only want Road
        return
    if collection_choice == 'record':
        collection = record
        if perf.category not in collection:
            collection[perf.category] = {}
        if perf.event not in collection[perf.category]:
            # First occurrence of this event so start new
            collection[perf.category][perf.event] = {}
        if perf.gender not in collection[perf.category][perf.event]:
            # First performance by this gender in this event so start new list
            collection[perf.category][perf.event][perf.gender] = []
        record_list = record[perf.category][perf.event][perf.gender]
        max_records = max_records_all if perf.category == 'ALL' else max_regords_age_group
        smaller_score_better = known_events_lookup[perf.event][0]
        compare_field = 'score'
    else:
        collection = wava
        if perf.event not in collection:
            collection[perf.event] = {}
        if year not in collection[perf.event]:
            collection[perf.event][year] = []
        record_list = collection[perf.event][year]
        max_records = max_wavas_all if year == 'ALL' else max_wavas_year
        smaller_score_better = False # WAVA bigger the better always
        compare_field = 'wava'


    add_record = False
    if len(record_list) < max_records:
        # We don't have enough records for this event yet so add
        add_record = True
    else:
        prev_worst_score = getattr(record_list[-1][0], compare_field)
        # For a tie, now adding new record, as club records sometimes showed two
        # record-holders; and a chance to line up data from different sources in
        # the output to show agreement where they match
        if smaller_score_better:
            if getattr(perf, compare_field) <= prev_worst_score: add_record = True
        else:
            if getattr(perf, compare_field) >= prev_worst_score: add_record = True

    if add_record:
        same_score_seen = False
        tie_same_name_managed = False
        for existing_perf_list in record_list:
            if getattr(existing_perf_list[0], compare_field) == getattr(perf, compare_field):
                same_score_seen = True
                # Same record with different source, or a tie with new person
                # Prefer Po10 over Runbritain, and don't include both as share source data
                for perf_idx, existing_perf in enumerate(existing_perf_list):
                    if existing_perf.athlete_name == perf.athlete_name:
                        if existing_perf.source.startswith('Po10') and perf.source.startswith('Runbritain'):
                            # Don't add Runbritain score if already there from Po10
                            tie_same_name_managed = True
                            break
                        elif  existing_perf.source.startswith('Runbritain') and perf.source.startswith('Po10'):
                            # Replace existing Runbritain one with Po10
                            existing_perf_list[perf_idx] = perf
                            tie_same_name_managed = True
                            break
                        else:
                            # Keep checking remaining performances for same name and score
                            pass
                if not tie_same_name_managed:
                    # Could be manual record to put alongside Po10 say
                    existing_perf_list.append(perf)
                break
        if not same_score_seen:
            record_list.append([perf])
            record_list.sort(key=lambda x: getattr(x[0], compare_field), reverse=not smaller_score_better)

        # Ensure new name only appears with their top score
        lowest_rec_idx_for_this_name = len(record_list) # i.e. not found yet
        rec_idx = 0
        while rec_idx < len(record_list):
            perf_idx = 0
            while perf_idx < len(record_list[rec_idx]):
                existing_record_name = record_list[rec_idx][perf_idx].athlete_name
                if existing_record_name == perf.athlete_name:
                    if rec_idx > lowest_rec_idx_for_this_name:
                        # Avoid same person appearing multiple times, allowing for ties,
                        # but nice to show multiple sources for identical performance
                        del record_list[rec_idx][perf_idx]
                        continue
                    elif rec_idx < lowest_rec_idx_for_this_name:
                        # Found best performance by this name so far
                        lowest_rec_idx_for_this_name = rec_idx
                    else:
                        # Someone else, ignore
                        pass
                perf_idx += 1
            if not record_list[rec_idx]:
                # Usual case after a deletion: no tie, that score was for only one athlete
                del record_list[rec_idx]
            else:
                rec_idx += 1
        # Keep list at max required length 
        del record_list[max_records :]


def process_po10_wava(perf, performance_cache, rebuild_cache, types):
    """Add a marathon performance to WAVA tables, respecting max size"""

    athlete_id_match = re.search(r'athleteid=([0-9]+)', perf.athlete_url)
    athlete_id = athlete_id_match.group(1)
    if athlete_id in wava_athlete_ids_done:
        return
    else:
        wava_athlete_ids_done[athlete_id] = True

    request_params = {'athleteid'   : athlete_id,
                      'viewby'      : 'agegraded'}

    url = powerof10_root_url + '/athletes/profile.aspx'
    cache_key = make_cache_key(url, request_params)
    if rebuild_cache:
        perf_list = None
    else:
        perf_list = performance_cache.get(cache_key, None)

    report_string_base = f'PowerOf10 WAVA for {perf.athlete_name} ID {athlete_id} '
    if perf_list is None:
        perf_list = []
        try:
            page_response = requests.get(url, request_params)
        except requests.exceptions.ConnectionError:
            print(report_string_base + ' ConnectionError')
            return

        print(report_string_base + f'page return status {page_response.status_code}')

        if page_response.status_code != 200:
            print(f'HTTP error code fetching page: {page_response.status_code}')
            return

        input_text = page_response.text
        tables = get_html_content(input_text, 'table')
        second_level_tables = []
        for table in tables:
            nested_tables = get_html_content(table.inner_text, 'table')
            second_level_tables.extend(nested_tables)
        third_level_tables = []
        for table in second_level_tables:
            nested_tables = get_html_content(table.inner_text, 'table')
            third_level_tables.extend(nested_tables)
        fourth_level_tables = []
        for table in third_level_tables:
            nested_tables = get_html_content(table.inner_text, 'table')
            fourth_level_tables.extend(nested_tables)
        all_tables = tables
        all_tables.extend(second_level_tables)
        all_tables.extend(third_level_tables)
        all_tables.extend(fourth_level_tables) # Think it is actually in here

        for table in all_tables:
            if 'class' not in table.attribs or table.attribs['class'] != 'alternatingrowspanel':
                continue
            rows = get_html_content(table.inner_text, 'tr')
            if len(rows) < 2:
                continue
            # Looks like we've found the table of results or something similar
            process_one_athlete_results_table(perf, rows, perf_list)

        performance_cache[cache_key] = perf_list
    else:
        print(report_string_base + f'{len(perf_list)} performances from cache')

    for perf in perf_list:
        year = get_year_from_po10_date(perf.date)
        process_performance(perf, types, 'wava', 'ALL')
        process_performance(perf, types, 'wava', str(year))


def get_year_from_po10_date(date_str):
    parts = date_str.split(' ')
    year_str = parts[2]
    year = int(year_str)
    year += 2000 # No Po10 results before this
    return year


def process_one_athlete_results_table(example_perf, rows, perf_list):

    heading_row = rows.pop(0)
    cells = get_html_content(heading_row.inner_text, 'td')
    heading_idx = {}
    for i, cell in enumerate(cells):
        heading = debold(cell.inner_text)
        heading_idx[heading] = i
    for expected_heading in ['Event', 'Perf', 'AGrade', 'Age', 'Venue', 'Date']:
        if expected_heading not in heading_idx:
            # Could be "UK Rankings" or "Athletes Coached" table, skip
            return
        
    for row in rows:
        cells = get_html_content(row.inner_text, 'td')
        event = cells[heading_idx['Event']].inner_text
        if event not in wava_events:
            continue
        performance = cells[heading_idx['Perf']].inner_text
        date = cells[heading_idx['Date']].inner_text
        venue_link = cells[heading_idx['Venue']]
        anchor = get_html_content(venue_link.inner_text, 'a')
        fixture_name = anchor[0].inner_text
        fixture_url = powerof10_root_url + anchor[0].attribs["href"]
        age_grade = cells[heading_idx['AGrade']].inner_text.strip()
        if not age_grade:
            # Could be multiterrain or XC or something, though should be excluded by event type anyway
            continue
        age_str = cells[heading_idx['Age']].inner_text.strip()
        age = 0 if not age_str else int(age_str)
        source = 'Po10'
        perf = construct_performance(event, example_perf.gender, 'ALL', performance, 
                                     example_perf.athlete_name, example_perf.athlete_url,
                                     date, fixture_name, fixture_url, source, age_grade=age_grade, age=age)
        perf_list.append(perf)


def process_one_rankings_table(rows, gender, category, source, perf_list, types):
    state = "seeking_title"
    row_idx = 0
    while True:
        if row_idx >= len(rows): return
        
        cells = get_html_content(rows[row_idx].inner_text, 'td')

        if state == "seeking_title":
            if 'class' not in rows[row_idx].attribs or rows[row_idx].attribs['class'] != 'rankinglisttitle':
                pass
            else:
                event_title = debold(cells[0].inner_text).strip()
                event = event_title.split(' ', 1)[0]
                state = "seeking_headings"
        elif state == "seeking_headings":
            if 'class' not in rows[row_idx].attribs or rows[row_idx].attribs['class'] != 'rankinglistheadings':
                pass
            else:
                heading_idx = {}
                for i, cell in enumerate(cells):
                    heading = debold(cell.inner_text)
                    heading_idx[heading] = i
                state = "seeking_results"
        elif state == "seeking_results":
            if 'class' not in rows[row_idx].attribs or not rows[row_idx].attribs['class'].startswith('rlr'):
                state = "seeking_title"
            else:
                name_link = cells[heading_idx['Name']]
                if name_link.inner_text: # Can get empty name if 2nd or more performances by same athlete
                    anchor = get_html_content(name_link.inner_text, 'a')
                    name = anchor[0].inner_text
                    url = powerof10_root_url + anchor[0].attribs["href"]
                    performance = cells[heading_idx['Perf']].inner_text
                    date = cells[heading_idx['Date']].inner_text
                    venue_link = cells[heading_idx['Venue']]
                    anchor = get_html_content(venue_link.inner_text, 'a')
                    fixture_name = anchor[0].inner_text
                    fixture_url = powerof10_root_url + anchor[0].attribs["href"]
                    perf = construct_performance(event, gender, category, performance, name, url, date, fixture_name, fixture_url, source)
                    perf_list.append(perf)
        else:
            # unknown state
            state = "seeking_title"
        row_idx += 1


def process_one_po10_year_gender(club_id, year, gender, category, performance_cache,
                                  rebuild_cache, first_claim_only, types, do_wava):

    request_params = {'clubid'         : str(club_id),
                      'agegroups'      : category,
                      'sex'            : gender,
                      'year'           : str(year),
                      'firstclaimonly' : 'y' if first_claim_only else 'n',
                      'limits'         : 'n'} # y faster for debug but don't want to miss rarely performed events so 'n' for completeness

    url = powerof10_root_url + '/rankings/rankinglists.aspx'
    cache_key = make_cache_key(url, request_params)
    if rebuild_cache:
        perf_list = None
    else:
        perf_list = performance_cache.get(cache_key, None)

    report_string_base = f'PowerOf10 club {club_id} year {year} gender {gender} category {category} '
    if perf_list is None:
        perf_list = []
        try:
            page_response = requests.get(url, request_params)
        except requests.exceptions.ConnectionError:
            print(report_string_base + ' ConnectionError')
            return

        print(report_string_base + f'page return status {page_response.status_code}')

        if page_response.status_code != 200:
            print(f'HTTP error code fetching page: {page_response.status_code}')
            return

        input_text = page_response.text
        source = f'Po10 {year}'

        tables = get_html_content(input_text, 'table')
        second_level_tables = []
        for table in tables:
            nested_tables = get_html_content(table.inner_text, 'table')
            second_level_tables.extend(nested_tables)
        
        for table in second_level_tables: # table of interest always a child table?
            rows = get_html_content(table.inner_text, 'tr')
            if len(rows) < 3:
                continue
            if 'class' not in rows[0].attribs or rows[0].attribs['class'] != 'rankinglisttitle':
                continue
            if 'class' not in rows[1].attribs or rows[1].attribs['class'] != 'rankinglistheadings':
                continue
            # Looks like we've found the table of results
            process_one_rankings_table(rows, gender, category, source, perf_list, types)

        performance_cache[cache_key] = perf_list
    else:
        print(report_string_base + f'{len(perf_list)} performances from cache')

    for perf in perf_list:
        process_performance(perf, types, 'record')
        performance_count['Po10'] += 1
        if do_wava and perf.event in wava_events:
            process_po10_wava(perf, performance_cache, rebuild_cache, types)


def make_cache_key(url, request_params):
    """Make unique string for web request to use as dict key for cache of previous
    web results"""

    cache_key = url + '?'
    for key, value in request_params.items(): # ordered as inserted in Python 3.6+
        cache_key += key + '=' + value + '&' # will leave trailing & but who cares
    return cache_key


def process_one_runbritain_year_gender(club_id, year, gender, category, event, performance_cache,
                                        rebuild_cache, first_claim_only, types):

    request_params = {'clubid'         : str(club_id),
                      'sex'            : gender,
                      'year'           : str(year),
                      'event'          : event,
                      'firstclaimonly' : 'y' if first_claim_only else 'n',
                      'limit'          : 'n'      } # Otherwise miss slower performances, undocumented option

    (min_age, max_age) = runbritain_category_lookup[category]
    if min_age == 0 and max_age == 0:
        # Use category name
        request_params['agegroup'] = category
    else:
        # Runbritain can miss results if use e.g. V40 category that it finds if
        # use numeric age range, so use latter instead
        request_params['agemin'] = str(min_age)
        request_params['agemax'] = str(max_age)

    url = runbritain_root_url + '/rankings/rankinglist.aspx'
    cache_key = make_cache_key(url, request_params)

    if rebuild_cache:
        perf_list = None
    else:
        perf_list = performance_cache.get(cache_key, None)

    report_string_base = f'Runbritain club {club_id} year {year} gender {gender} category {category} event {event} '

    if perf_list is None:
        perf_list = []
        try:
            page_response = requests.get(url, request_params)
        except requests.exceptions.ConnectionError:
            print(report_string_base + ' ConnectionError')
            return

        print(report_string_base + f'page return status {page_response.status_code}')

        if page_response.status_code != 200:
            print(f'HTTP error code fetching page: {page_response.status_code}')
            return

        input_text = page_response.text
        results_array_regex = re.compile(r'runners =\s*(\[.*?\]);', flags=re.DOTALL)
        array_match = results_array_regex.search(input_text)

        if array_match is None:
            print('No data found')
        else:
            source = f'Runbritain {year}'
            array_str = array_match.group(1)
            array_str = array_str.replace('\n', ' ').replace('\r', '')
            results_array = eval(array_str)
            for result in results_array:
                if not result[6] : continue # No name, could be second performance by same person
                anchor = get_html_content(result[6], 'a')
                name = anchor[0].inner_text
                url = runbritain_root_url + anchor[0].attribs["href"]
                perf = result[1] # Chip time
                if not perf:
                    perf = result[3] # Gun time
                date = result[10]
                venue_link = result[9]
                anchor = get_html_content(venue_link, 'a')
                fixture_name = anchor[0].inner_text
                fixture_url = runbritain_root_url + anchor[0].attribs["href"]
                perf = construct_performance(event, gender, category, perf, name, url, date, fixture_name, fixture_url, source)
                perf_list.append(perf)

        performance_cache[cache_key] = perf_list
    else:
        print(report_string_base + f'{len(perf_list)} performances from cache')
    
    for perf in perf_list:
        process_performance(perf, types, 'record')
        performance_count['Runbritain'] += 1


def format_sexagesimal(value, num_numbers, decimal_places):
    """Format as HH:MM:SS (3 numbers), SS.sss (1 number) etc"""
    output = ''
    divisor = 60 ** (num_numbers - 1)
    for i in range(num_numbers - 1):
        quotient = int(value / divisor)
        if i == 0:
            output += '%d:' % quotient
        else:
            output += "%.2d:" % quotient # leading zero if needed after first number
        value -= (quotient * divisor)
        divisor /= 60
    
    if   decimal_places == 0:
        fmt = '%02.0f' if num_numbers > 1 else '%.0f'
    elif decimal_places == 1:
        fmt = '%04.1f' if num_numbers > 1 else '%.1f'
    elif decimal_places == 2:
        fmt = '%05.2f' if num_numbers > 1 else '%.2f'
    else:
        fmt = '%06.3f' if num_numbers > 1 else '%.3f'
    output += fmt % value

    return output


def output_records(output_file, first_year, last_year, club_id, do_po10, do_runbritain, input_files):

    header_part = [
                    '<html>\n',
                    '<body>\n',
                    f'<h1>Club Records</h1>\n\n',
                    f'<p>Initially autogenerated  on {datetime.date.today()} from:</p>\n',
                    f'<ul>\n']
    if do_po10:
        header_part.append(f'<li><a href="{powerof10_root_url}/clubs/club.aspx?clubid={club_id}">PowerOf10 club page</a>  {first_year} - {last_year}</li>\n')
    if do_runbritain:
        header_part.append(f'<li><a href="{runbritain_root_url}/rankings/rankinglist.aspx">runbritain rankings</a>  {first_year} - {last_year}</li>\n')
    for input_file in input_files:
        header_part.append(f'<li>Local file: {input_file}</li>\n')
    header_part.append('</ul>\n\n')
    header_part.append(f'<p>Outputting maximum {max_records_all} places overall per event and {max_regords_age_group} per age group.</p>\n')
    header_part.append(f'<p>Count of performances processed...')
    for type in performance_count.keys():
        header_part.append(f' {type}: {performance_count[type]}')
    header_part.append(f'</p>\n')
    header_part.append(f'<p><em>See <a href="https://github.com/charlie-wartnaby/powerof10-tools">https://github.com/charlie-wartnaby/powerof10-tools</a> for source code.</em></p>\n\n')

    contents_part = ['<h2>Contents</h2>\n\n']
    
    bulk_part = []

    first_content = True

    year_keys = ['ALL']
    for year in range(first_year, last_year + 1):
        year_keys.append(str(year))

    for event in wava_events:
        if event not in wava:
            continue
        anchor = f'wava_{event}'.lower()
        subtitle = event
        anchor = 'wava_tables'
        subtitle = f'Age Grade: {event}'
        if first_content:
            # Avoid big gap after Contents heading
            first_content = False
        else:
            contents_part.append('<br />\n')
        contents_part.append(f'<b><a href="#{anchor}">{subtitle}</a></b><br />\n')
        bulk_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
        for year_key in year_keys:
            if year_key not in wava[event]:
                continue
            record_list = wava[event][year_key]
            anchor = f'wava_{event}_{year_key}'.lower()
            subtitle = year_key
            contents_part.append(f'<em><a href="#{anchor}">...{subtitle}</a></em>\n')
            bulk_part.append(f'<h3><a name="{anchor}" />Age Grade {event} year: {subtitle}</h3>\n\n')
            output_record_table(bulk_part, event, record_list, 'wava')

    for (category, _, _) in runbritain_categories:
        if category not in record: continue
        for gender in ['W', 'M']:
            anchor = f'category_{gender}_{category}'
            subtitle = f'Category: {gender} {category}'
            if first_content:
                # Avoid big gap after Contents heading
                first_content = False
            else:
                contents_part.append('<br />\n')
            contents_part.append(f'<b><a href="#{anchor}">{subtitle}</a></b><br />\n')
            bulk_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
            for (event, _, _, _, _) in known_events:
                if event not in record[category]: continue
                record_list = record[category][event].get(gender)
                if not record_list: continue
                anchor = f'{event}_{gender}_{category}'.lower()
                subtitle = f'{event} {gender} {category}'
                contents_part.append(f'<em><a href="#{anchor}">...{subtitle}</a></em>\n')
                bulk_part.append(f'<h3><a name="{anchor}" />Records for {subtitle}</h3>\n\n')
                output_record_table(bulk_part, event, record_list, 'record')

    contents_part.append('\n\n')

    tail_part = ['</body>\n',
                '</html>\n']

    with open(output_file, 'wt') as fd:
        for part in [header_part, contents_part, bulk_part, tail_part]:
            for line in part:
                fd.write(line)

def output_record_table(bulk_part, event, record_list, type):
    bulk_part.append('<table border="2">\n')
    bulk_part.append('<tr>\n')
    bulk_part.append('<td><center><b>Rank</b></center></td>')
    if type == 'wava':
        bulk_part.append('<td><center><b>Age Grade %</b></center></td><td><center><b>Age</b></center></td>')
    bulk_part.append('<td><center><b>Performance</b></center></td><td><center><b>Athlete</b></center></td><td><center><b>Date</b></center></td><td><center><b>Fixture</b></center><td><center><b>Source</b></center></td>\n')
    bulk_part.append('</tr>\n')
    for idx, perf_list in enumerate(record_list):
        for perf_idx, perf in enumerate(perf_list): # May be ties with same score or different sources
            if perf.original_special:
                score_str = perf.original_special
            else:
                score_str = format_sexagesimal(perf.score, known_events_lookup[event][1], perf.decimal_places)
            bulk_part.append('<tr>\n')
            rank_str = f'{idx+1}' if perf_idx == 0 else ''
            bulk_part.append(f'  <td><center>{rank_str}</center></td>\n')
            if type == 'wava':
                wava_str = '%.2f' % perf.wava
                bulk_part.append(f'  <td><center>{wava_str}</td>\n')
                bulk_part.append(f'  <td><center>{perf.age}</td>\n')
            bulk_part.append(f'  <td><center>{score_str}</td>\n')
            if perf.athlete_url:
                bulk_part.append(f'  <td><a href="{perf.athlete_url}">{perf.athlete_name}</a></td>\n')
            else:
                bulk_part.append(f'  <td>{perf.athlete_name}</td>\n')
            bulk_part.append(f'  <td>{perf.date}</td>\n')
            if perf.fixture_url:
                bulk_part.append(f'  <td><a href="{perf.fixture_url}">{perf.fixture_name}</a></td>\n')
            else:
                bulk_part.append(f'  <td>{perf.fixture_name}</td>\n')
            bulk_part.append(f'  <td>{perf.source}</td>\n')
            bulk_part.append('</tr>\n')
    bulk_part.append('</table>\n\n')


def process_one_input_file(input_file, types):

    print(f'Processing file: {input_file}')

    _, file_extension = os.path.splitext(input_file)
    if file_extension.lower() != '.xlsx':
        print(f'WARNING: ignoring input file, can only handle .xlsx currently: {input_file}')
        return
    
    workbook = openpyxl.load_workbook(filename=input_file)
    for worksheet in workbook.worksheets:
        process_one_excel_worksheet(input_file, worksheet, types)


def process_one_excel_worksheet(input_file, worksheet, types):

    print(f'Processing worksheet: {worksheet.title}')
    df = pandas.DataFrame(worksheet.values)

    # We get integers as headings instead of the intended column headings, so find those
    headings_found = False
    for row_idx, row in df.iterrows():
        for col_idx, cell_value in enumerate(row.values):
            if cell_value is None: continue
            if cell_value.lower().strip() == 'performance':
                headings_found = True
                break
        if headings_found: break

    if not headings_found:
        print(f'WARNING: could not find "Performance" heading, skipping sheet')
        return
    
    df.columns = df.iloc[row_idx]
    df.drop(df.index[0:row_idx + 1], inplace=True)
    
    # Convert all headings to lower case and strip whitespace
    renames = {}
    for col_name in df.columns:
        if not col_name : continue
        new_col_name = col_name.lower().strip()
        renames[col_name] = new_col_name
    df.rename(columns=renames, inplace=True)
    
    # C&C club records have some column headings that might not suit us for other
    # input lists
    df.rename(columns={'year' : 'date', 'record holder' : 'name'}, inplace=True)

    col_name_list = df.columns.tolist()
    for reqd_heading in ['performance', 'date', 'name', 'po10 event', 'gender', 'age code']:
        if reqd_heading not in col_name_list:
            if reqd_heading == 'date': reqd_heading = 'date [or year]'
            if reqd_heading == 'name': reqd_heading = 'date [or record holder]'
            print(f'Required heading not found (case insensitive), skipping sheet: {reqd_heading}')
            return

    for row_idx, row in df.iterrows():
        # Can get None obj references as well as empty strings
        excel_row_number = row_idx + 1
        perf = row['performance']
        name = row['name']
        if not perf and not name:
            # Assume blank row, ignore quietly
            continue
        perf = str(perf).strip() if perf else ''
        name = name.strip() if name else ''
        if not name and not perf:
            # Assume blank row, ignore quietly
            continue
        if not perf:
            print(f'WARNING: performance missing at row {excel_row_number}')
            continue
        if not name:
            print(f'WARNING: name missing at row {excel_row_number}')
            continue
        # Name URL is optional
        name_url = row['name url'] if 'name url' in row else None
        name_url = str(name_url).strip() if name_url else ''
        date = row['date']
        if date is None:
            print(f'WARNING: date missing at row {excel_row_number}')
            continue
        # Fixture is optional
        fixture = row['fixture'] if 'fixture' in df.columns else None
        fixture = str(fixture).strip() if fixture else ''
        fixture_url = row['fixture url'] if 'fixture url' in df.columns else None
        fixture_url = str(fixture_url).strip() if fixture_url else ''
        event = row['po10 event']
        if event is None:
            print(f'WARNING: Po10 event code missing at row {excel_row_number}')
            continue
        gender = row['gender']
        if gender is None:
            print(f'WARNING: gender missing at row {excel_row_number}')
            continue
        category = row['age code']
        category = str(category).strip() if category is not None else ''
        if category is None:
            print(f'WARNING: age category missing at row {excel_row_number}')
            continue
        perf = str(perf).strip()
        if not perf:
            print(f'WARNING: performance missing at row {excel_row_number}')
            continue
        date = str(date).strip()
        if not date:
            print(f'WARNING: date missing at row {excel_row_number}')
            continue
        if not name:
            print(f'WARNING: name missing at row {excel_row_number}')
            continue
        event = str(event).strip()
        if not event:
            print(f'WARNING: event missing at row {excel_row_number}')
            continue
        gender = gender.upper().strip()
        if gender not in ['M', 'W']:
            print(f'WARNING: gender not W or M at row {excel_row_number}')
            continue
        perf = construct_performance(event, gender, category, perf, name, name_url,
                            date, fixture, fixture_url, input_file + ':' + worksheet.title)
        process_performance(perf, types, 'record')
        performance_count['File(s)'] += 1


def main(club_id=238, output_file='records.htm', first_year=2005, last_year=2023, 
         do_po10=False, do_runbritain=True, input_files=[],
         cache_file='cache.pkl', rebuild_last_year=False, first_claim_only=False,
         types=['T', 'F', 'R', 'M'], do_wava=True):

    # Input files first so known club records appear above database results for same performance
    for input_file in input_files:
        process_one_input_file(input_file, types)

    # Retrieve cache of performances obtained from web trawl previously
    try:
        with open(cache_file, 'rb') as fd:
            performance_cache = pickle.load(fd)
            print(f'Cached web results retrieved from {cache_file}')
    except IOError:
        print(f"Cache file {cache_file} can't be opened, starting new cache")
        performance_cache = {}

    for year in range(first_year, last_year + 1):
        rebuild_cache = rebuild_last_year and year == last_year
        for gender in ['W', 'M']:
            if do_po10:
                for category in powerof10_categories:
                    process_one_po10_year_gender(club_id, year, gender, category,
                                                 performance_cache, rebuild_cache, first_claim_only,
                                                 types, do_wava)
            if do_runbritain:
                for (event, _, _, runbritain, type) in known_events: # debug [('Mar', True, 3, True, 'R')]:
                    if not runbritain: continue
                    if type not in types: continue
                    for (category, _, _) in runbritain_categories: # debug [('ALL', 0, 0), ('V50', 50, 54)]
                        process_one_runbritain_year_gender(club_id, year, gender, category, event,
                                                           performance_cache, rebuild_cache,
                                                           first_claim_only, types)

    # Save updated cache for next time
    try:
        with open(cache_file, 'wb') as fd:
            pickle.dump(performance_cache, fd)
        print(f'Cached web results written to {cache_file}')
    except IOError:
        print(f"Cache file {cache_file} can't be written, any new web results this time not cached")

    output_records(output_file, first_year, last_year, club_id, do_po10, do_runbritain, input_files)


if __name__ == '__main__':
    # Main entry point

    parser = argparse.ArgumentParser(description='Build club records tables from thepowerof10, runbritain and Excel files')
 
    yes_no_choices = ['y', 'Y', 'n', 'N']
    cnc_po10_club_id = 238
    this_year = datetime.datetime.now().year

    parser.add_argument(dest='excel_file', nargs ='*') # .xlsx records files
    parser.add_argument('--powerof10', dest='do_po10', choices=yes_no_choices, default='y')
    parser.add_argument('--runbritain', dest='do_runbritain', choices=yes_no_choices, default='y')
    parser.add_argument('--firstyear', dest='first_year', type=int, default=2004) # A few Po10 results in 2004
    parser.add_argument('--lastyear', dest='last_year', type=int, default=this_year)
    parser.add_argument('--clubid', dest='club_id', type=int, default=cnc_po10_club_id)
    parser.add_argument('--output', dest='output_filename', default='records.htm')
    parser.add_argument('--cache', dest='cache_filename', default='cache.pkl')
    parser.add_argument('--rebuild-last-year', dest='rebuild_last_year',  choices=yes_no_choices, default='n')
    parser.add_argument('--first-claim-only', dest='first_claim_only',  choices=yes_no_choices, default='n')
    parser.add_argument('--track', dest='track',  choices=yes_no_choices, default='y')
    parser.add_argument('--field', dest='field',  choices=yes_no_choices, default='y')
    parser.add_argument('--road', dest='road',  choices=yes_no_choices, default='y')
    parser.add_argument('--multievent', dest='multievent',  choices=yes_no_choices, default='y')
    parser.add_argument('--wava', dest='wava',  choices=yes_no_choices, default='n')

    args = parser.parse_args()

    do_po10           = args.do_po10.lower().startswith('y')
    do_runbritain     = args.do_runbritain.lower().startswith('y')
    rebuild_last_year = args.rebuild_last_year.lower().startswith('y')
    first_claim_only  = args.first_claim_only.lower().startswith('y')
    do_wava           = args.wava.lower().startswith('y')
    types = []
    if args.track.lower().startswith('y'):      types.append('T')
    if args.field.lower().startswith('y'):      types.append('F')
    if args.road.lower().startswith('y'):       types.append('R')
    if args.multievent.lower().startswith('y'): types.append('M')

    main(club_id=args.club_id, output_file=args.output_filename, first_year=args.first_year, 
         last_year=args.last_year, do_po10=do_po10, do_runbritain=do_runbritain, 
         input_files=args.excel_file, cache_file=args.cache_filename, rebuild_last_year=rebuild_last_year,
         first_claim_only=first_claim_only, types=types, do_wava=do_wava)
