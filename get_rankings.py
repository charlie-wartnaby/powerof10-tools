from array import array
import datetime
import math
import re
import requests
import sys


class Performance():
    def __init__(self, event, score, decimal_places, athlete_name, athlete_url='', date='',
                       fixture_name='', fixture_url='', source=''):
        self.event = event
        self.score = score # could be time in sec, distance in m or multievent points
        self.decimal_places = decimal_places # so we can use original precision which may imply electronic timing etc
        self.athlete_name = athlete_name
        self.athlete_url = athlete_url
        self.date = date
        self.fixture_name = fixture_name
        self.fixture_url = fixture_url
        self.source = source

class HtmlBlock():
    def __init__(self, tag=''):
        self.tag = tag
        self.inner_text = ''
        self.attribs = {}

record = {} # dict of events, each dict of genders, then ordered list of performances
max_records_all = 10 # max number of records for each event/gender, including all age groups
max_regords_age_group = 3 # Similarly per age group
powerof10_root_url = 'https://thepowerof10.info'
runbritain_root_url = 'https://www.runbritainrankings.com'

# TODO should maybe use original precision for performances as don't want to imply electronic
# timing with additional decimal places that were never originally there, say.

# Smaller time is good for runs, bigger distance/score better for jumps/throws/multievents;
# some events should be in sec (1 number), some in min:sec (2 numbers), some h:m:s (3 numbers):
#                event, small-is-good, :-numbers
known_events = [
                ('10K',     True,        2),
                ('HM',      True,        2),
                ('Mar',     True,        3),
                ('60' ,     True,        1),
                ('100',     True,        1),
                ('200',     True,        1),
                ('400',     True,        1),
                ('800',     True,        2),
                ('1500',    True,        2),
                ('3000',    True,        2),
                ('5000',    True,        2),
                ('10000',   True,        2),
                ('3000SC',  True,        2),
                ('3000SCW', True,        2),
                ('100HW',   True,        1),
                ('110H',    True,        1),
                ('400H',    True,        1),
                ('400HW',   True,        1),
                ('HJ',      False,       1),
                ('PV',      False,       1),
                ('LJ',      False,       1),
                ('TJ',      False,       1),
                ('SP4K',    False,       1),
                ('SP7.26K', False,       1),
                ('DT1K',    False,       1),
                ('DT2K',    False,       1),
                ('HT4K',    False,       1),
                ('HT7.26K', False,       1),
                ('JT600',   False,       1),
                ('JT800',   False,       1),
                ('HepW',    False,       1),
                ('Dec',     False,       1)
 ]

known_events_lookup = {}
for (event, smaller_better, numbers) in known_events:
    known_events_lookup[event] = (smaller_better, numbers)


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
                print(f'Warning: no match for closing tag "{html_tag}"')
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
    total_score = 0.0
    multiplier = 1.0
    sexagesimals = perf.split(':')
    sexagesimals.reverse()
    for sexagesmial in sexagesimals:
        total_score += float(sexagesmial) * multiplier
        multiplier *= 60.0
    decimal_split = perf.split('.')
    decimal_places = 0 if len(decimal_split) < 2 else len(decimal_split[1])

    return total_score, decimal_places


def process_performance(event, gender, perf, name, url, date, fixture_name, fixture_url, source):
    if event not in record:
        # First occurrence of this event so start new
        record[event] = {}
    if gender not in record[event]:
        # First performance by this gender in this event so start new list
        record[event][gender] = []

    if event not in known_events_lookup:
        print(f'Warning: unknown event {event}, ignoring')
        return
    smaller_score_better = known_events_lookup[event][0]

    score, original_dp = make_numeric_score_from_performance_string(perf)

    record_list = record[event][gender]
    add_record = False
    if len(record_list) < max_records_all:
        # We don't have enough records for this event yet so add
        add_record = True
    else:
        prev_worst_score = record_list[-1].score
        # For a tie, not adding new record as the earlier one should
        # take precedence, but TODO not considering fixture date yet,
        # only depending on processing rankings in year order
        if smaller_score_better:
            if score < prev_worst_score: add_record = True
        else:
            if score > prev_worst_score: add_record = True
        # TODO getting equal scores in resuls currently...

    if add_record:
        perf = Performance(event, score, original_dp, name, powerof10_root_url+url, 
                           date, fixture_name, powerof10_root_url+fixture_url, source)
        record_list.append(perf)
        record_list.sort(key=lambda x: x.score, reverse=not smaller_score_better)
        athlete_names = {}
        idx = 0
        while idx < len(record_list):
            existing_record_name = record_list[idx].athlete_name
            if existing_record_name in athlete_names:
                # Avoid same person appearing multiple times
                del record_list[idx]
            else:
                athlete_names[existing_record_name] = True
            idx += 1
        # Keep list at max required length 
        del record_list[max_records_all :]

def process_one_rankings_table(rows, gender, source):
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
                    url = anchor[0].attribs["href"]
                    perf = cells[heading_idx['Perf']].inner_text
                    date = cells[heading_idx['Date']].inner_text
                    venue_link = cells[heading_idx['Venue']]
                    anchor = get_html_content(venue_link.inner_text, 'a')
                    fixture_name = anchor[0].inner_text
                    fixture_url = anchor[0].attribs["href"]
                    process_performance(event, gender, perf, name, url, date, fixture_name, fixture_url, source)
        else:
            # unknown state
            state = "seeking_title"
        row_idx += 1

def process_one_po10_year_gender(club_id, year, gender):

    request_params = {'clubid'         : str(club_id),
                      'agegroups'      : 'ALL',
                      'sex'            : gender,
                      'year'           : str(year),
                      'firstclaimonly' : 'y',
                      'limits'         : 'n'} # y faster for debug but don't want to miss rarely performed events so 'n' for completeness

    page_response = requests.get(powerof10_root_url + '/rankings/rankinglists.aspx', request_params)

    print(f'Club {club_id} year {year} gender {gender} page return status {page_response.status_code}')

    if page_response.status_code != 200:
        print(f'HTTP error code fetching page: {page_response.status_code}')
        return

    debug = False
    if debug:
        with open('shortened_example.htm') as fd:
            input_text = fd.read()
    else:
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
        process_one_rankings_table(rows, gender, source)

    if debug:
        sys.exit(0)

def process_one_runbritain_year_gender(club_id, year, gender, event):

    request_params = {'clubid'       : str(club_id),
                      'agegroup'     : 'ALL',
                      'sex'          : gender,
                      'year'         : str(year),
                      'event'        : event}

    page_response = requests.get(runbritain_root_url + '/rankings/rankinglist.aspx', request_params)

    print(f'Runbritain club {club_id} year {year} gender {gender} page return status {page_response.status_code}')

    if page_response.status_code != 200:
        print(f'HTTP error code fetching page: {page_response.status_code}')
        return

    input_text = page_response.text
    results_array_regex = re.compile(r'runners =\s*(\[.*?\];)', flags=re.DOTALL)
    array_match = results_array_regex.search(input_text)

    if array_match is None:
        print('No data found')
    else:
        array_str = array_match.group(1)
        results_array = eval(array_str)
        #process_one_rankings_table(rows, gender)

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


def output_records(output_file, first_year, last_year, club_id):

    with open(output_file, 'wt') as fd:
        fd.write('<html>\n')
        fd.write('<body>\n')
        fd.write(f'<h1>Club Records</h1>\n')
        fd.write(f'<p>Initially autogenerated from ')
        fd.write(f'<a href="{powerof10_root_url}/clubs/club.aspx?clubid={club_id}">PowerOf10 club page</a>')
        fd.write(f' {first_year} - {last_year}')
        fd.write(f' on {datetime.date.today()}.</p>\n\n')
        for (event, _, _) in known_events: # debug: ['10K', 'HM', 'Mar', 'LJ', 'HepW', 'Dec']:
            if event not in record: continue
            for gender in ['W', 'M']:
                record_list = record[event].get(gender)
                if not record_list: continue
                fd.write(f'<h2>Records for {event} {gender}</h2>\n\n')
                fd.write('<table border="2">\n')
                fd.write('<tr>\n')
                fd.write('<td><b>Rank</b></td><td><b>Performance</b></td><td><b>Athlete</b></td><td><b>Date</b></td><td><b>Fixture</b><td><b>Source</b></td>\n')
                fd.write('</tr>\n')
                for idx, perf in enumerate(record_list):
                    score_str = format_sexagesimal(perf.score, known_events_lookup[event][1], perf.decimal_places)
                    fd.write('<tr>\n')
                    fd.write(f'  <td>{idx+1}</td>\n')
                    fd.write(f'  <td>{score_str}</td>\n')
                    fd.write(f'  <td><a href="{perf.athlete_url}"> {perf.athlete_name}</a></td>\n')
                    fd.write(f'  <td>{perf.date}</td>\n')
                    fd.write(f'  <td><a href="{perf.fixture_url}"> {perf.fixture_name}</a></td>\n')
                    fd.write(f'  <td>{perf.source}</td></n>')
                    fd.write('</tr>\n')
                fd.write('</table>\n\n')
        fd.write('</body>\n')
        fd.write('</html>\n')

def main(club_id=238, output_file='records.htm', first_year=2006, last_year=2006, do_po10=False, do_runbritain=True):

    for year in range(first_year, last_year + 1):
        for gender in ['W', 'M']:
            if do_po10:
                process_one_po10_year_gender(club_id, year, gender)
            if do_runbritain:
                for event in  ['Mar']:
                    process_one_runbritain_year_gender(club_id, year, gender, event)

    output_records(output_file, first_year, last_year, club_id)

if __name__ == '__main__':
    main()
