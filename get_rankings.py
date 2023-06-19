import re
import requests

class Performance():
    def __init__(self):
        self.event_name = ''
        self.time_sec = 0.0
        self.runner_name = ''
        self.runner_url = ''
        self.fixture_name = ''
        self.fixture_url = ''

class HtmlBlock():
    def __init__(self, _tag=''):
        self.tag = _tag
        self.inner_text = ''
        self.attribs = {}

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

def process_one_rankings_table(rows):
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
                print(event_title)
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
                name = cells[heading_idx['Name']].inner_text
                perf = cells[heading_idx['Perf']].inner_text
                print(name, perf)
                pass
        else:
            # unknown state
            state = "seeking_title"
        row_idx += 1

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
        process_one_rankings_table(rows)


def main(club_id=238):
    for year in range(2005,2006):
        for gender in ['W', 'M']:
            process_one_year_gender(club_id, year, gender)

if __name__ == '__main__':
    main()
