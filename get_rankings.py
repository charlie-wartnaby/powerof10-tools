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
    debug_open_regex = re.compile(r'<' + html_tag)

    all_matches =[]
    for match in  debug_open_regex.finditer(html_text):
        all_matches.append(match)
        tag_start = match.start()
        print(f'Debug match at {tag_start}: {html_text[tag_start:tag_start + 30]}')

    contents = []
    offset = 0

    while True:
        open_match = open_regex.search(html_text, pos=offset)
        if open_match is None:
            break
        idx_char_after_opening_tag = open_match.end()
        debug_idx_char_start_tag = open_match.start()
        close_match = close_regex.search(html_text, pos=idx_char_after_opening_tag)
        if close_match is None:
            print(f'Warning: unmatched opening tag {html_tag} found and ignored')
            break
        idx_first_char_closing_tag = close_match.start()
        content_substr = html_text[idx_char_after_opening_tag : idx_first_char_closing_tag]
        contents.append(content_substr)
        idx_char_after_closing_tag = close_match.end()
        offset = idx_char_after_closing_tag

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
    pass

def main(club_id=238):
    for year in range(2005,2006):
        for gender in ['F', 'M']:
            process_one_year_gender(club_id, year, gender)

if __name__ == '__main__':
    main()
