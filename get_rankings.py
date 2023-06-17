import requests


def process_one_year_gender(club_id, year, gender):

    request_params = {'clubid' : str(club_id),
                      'agegroups' : 'ALL',
                      'sex' : gender,
                      'year' : str(year),
                      'firstclaimonly' : 'y',
                      'limits' : 'n'}

    page_response = requests.get('https://thepowerof10.info/rankings/rankinglists.aspx', request_params)

    print(f'Club {club_id} year {year} gender {gender} page return status {page_response.status_code}')

def main(club_id=238):
    for year in range(2005,2006):
        for gender in ['F', 'M']:
            process_one_year_gender(club_id, year, gender)

if __name__ == '__main__':
    main()
