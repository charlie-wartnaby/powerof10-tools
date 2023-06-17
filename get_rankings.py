import re
import requests

class Performance():
    def __init__(self):
        event_name = ''
        time_sec = 0.0
        runner_name = ''
        runner_url = ''
        fixture_name = ''
        fixture_url = ''

def get_html_content(html_text, html_tag):
    """Extracts instances of text data enclosed by required tag"""

    open_regex = re.compile(r'<' + html_tag + r'.*?>', flags=re.DOTALL)
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
            print('Warning: no match for closing tag "f{html_tag}"')
            break
        if open_match is not None and (open_match.start() < close_match.start()):
            if inside_tag_block:
                nesting_depth += 1
            else:
                block_content_start_idx = open_match.end()
                inside_tag_block = True
            offset = open_match.end()
        else:
            if nesting_depth > 0:
                # keep searching for tag that closes original opening tag
                nesting_depth -= 1
            else:
                content_substr = html_text[block_content_start_idx : close_match.start()]
                contents.append(content_substr)
                inside_tag_block = False
            offset = close_match.end()

    return contents

def process_one_year_gender(club_id, year, gender):

    request_params = {'clubid' : str(club_id),
                      'agegroups' : 'ALL',
                      'sex' : gender,
                      'year' : str(year),
                      'firstclaimonly' : 'y',
                      'limits' : 'n'}

    page_response = requests.get('https://thepowerof10.info/rankings/rankinglists.aspx', request_params)

    print(f'Club {club_id} year {year} gender {gender} page return status {page_response.status_code}')

    if page_response.status_code != 200:
        raise Exception(f'HTTP error code fetching page: {page_response.status_code}')

    tables = get_html_content(page_response.text, 'table')
    second_level_tables = []
    for table in tables:
        nested_tables = get_html_content(table, 'table')
        second_level_tables.extend(nested_tables)

    pass

def main(club_id=238):
    for year in range(2005,2006):
        for gender in ['W', 'M']:
            process_one_year_gender(club_id, year, gender)

if __name__ == '__main__':
    main()
