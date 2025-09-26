# First written by (c) Charlie Wartnaby 2023
# See https://github.com/charlie-wartnaby/powerof10-tools">https://github.com/charlie-wartnaby/powerof10-tools

# See/use requirements.txt for additional module dependencies
import argparse
import copy
import datetime
import openpyxl
import os
import pandas
import pickle
import re
import requests
import sys

if sys.version_info.major < 3:
    print('This script needs Python 3')
    sys.exit(1)
if sys.version_info.minor < 6:
    print('WARNING: this script assumes ordered dicts (Python 3.6+) when it builds cache keys, bad things may happen here')


class Performance():

    def __init__(self, event, score, category, gender, original_special, decimal_places, athlete_name, athlete_url='', date='',
                       fixture_name='', fixture_url='', source='', wava=0.0, age=0, invalid=False, ea_pb_score=0.0):
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
        # Added so that Po10/Runbritain records could be removed (e.g. if athlete known to no longer be in club):
        self.invalid = invalid
        # Added for England Athletics PB Award scheme, cached performances will not have
        # but currently computed when used anyway:
        self.ea_pb_score = ea_pb_score

class HtmlBlock():
    def __init__(self, tag=''):
        self.tag = tag
        self.inner_text = ''
        self.attribs = {}

record = {} # dict of age groups, each dict of events, each dict of genders, then ordered list of performance lists (allowing for ties)
wava = {} # dict of events, each dict of years and 0 for all years, then ordered list of performance lists (to allow for ties)
ea_pb = {} # dict of buckets (e.g. throws or sprints)
max_records_all = 10 # max number of records for each event/gender, including all age groups
max_records_age_group = 3 # Similarly per age group
max_wavas_all = 20  # All-time WAVA list
max_wavas_year = 5 # WAVA list for specific year
wava_athlete_ids_done = {}
max_ea_pbs_all = max_wavas_all
max_ea_pbs_year = max_wavas_year
ea_pb_limit_perf_fraction = 0.25   # Improvement beyond level 9 we use for 'level 10'

powerof10_root_url = 'https://thepowerof10.info'
runbritain_root_url = 'https://www.runbritainrankings.com'

common_table_attribs = 'border="2" style="width:100%"'

performance_count = {'Po10'       : 0,
                     'Runbritain' : 0,
                     'Po10-WAVA'  : 0,
                     'File(s)'    : 0}

wava_events = ['Mar', 'HM', '10K', '5K']  # C&C trophy category but could do other events

# Smaller time is good for runs, bigger distance/score better for jumps/throws/multievents;
# some events should be in sec (1 number), some in min:sec (2 numbers), some h:m:s (3 numbers).
# Some events only relevant to certain gender and/or age category; those found by checking po10
# national rankings for different age groups to see e.g. which discus weight listed.
#                event, small-is-good, :-numbers, runbritain,  Track/Field/Road/Multievent, Categories
known_events = [
                ('1M',         True,        2,        True,     'R',                        []     ),
                ('2M',         True,        2,        True,     'R',                        []     ),
                ('5K',         True,        2,        True,     'R',                        []     ),
                ('parkrun',    True,        2,        True,     'R',                        []     ),
                ('4M',         True,        2,        True,     'R',                        []     ),
                ('5M',         True,        2,        True,     'R',                        []     ),
                ('10K',        True,        2,        True,     'R',                        []     ),
                ('10M',        True,        2,        True,     'R',                        []     ),
                ('HM',         True,        2,        True,     'R',                        []     ),
                ('Mar',        True,        3,        True,     'R',                        []     ),
                ('50K',        True,        3,        True,     'R',                        []     ),
                ('100K',       True,        3,        True,     'R',                        []     ),
                ('60' ,        True,        1,        True,     'T',                        []     ),
                ('80' ,        True,        1,        True,     'T',                        []     ),
                ('100',        True,        1,        True,     'T',                        []     ),
                ('150',        True,        1,        True,     'T',                        []     ),
                ('200',        True,        1,        True,     'T',                        []     ),
                ('300',        True,        1,        True,     'T',                        []     ),
                ('400',        True,        1,        True,     'T',                        []     ),
                ('600',        True,        1,        True,     'T',                        []     ),
                ('800',        True,        2,        True,     'T',                        []     ),
                ('1500',       True,        2,        True,     'T',                        []     ),
                ('Mile',       True,        2,        True,     'T',                        []     ),
                ('3000',       True,        2,        True,     'T',                        []     ),
                ('5000',       True,        2,        True,     'T',                        []     ),
                ('10000',      True,        2,        True,     'T',                        []     ),
                ('1500SC',     True,        2,        True,     'T',                        ['M U17']     ),
                ('1500SCW',    True,        2,        True,     'T',                        ['W ALL', 'W U17', 'W U20', 'W U23', 'W 45', 'W V50', 'W V55', 'W V60']     ), # Included W ALL because official C&C records do, and so many age groups
                ('2000SC',     True,        2,        True,     'T',                        ['M ALL', 'M U17', 'M U20', 'M U23', 'M V35', 'M V40', 'M V45']     ), # Included M ALL because official C&C records do, and so many age groups
                ('2000SCW',    True,        2,        True,     'T',                        ['W ALL', 'W U20', 'W U23', 'W V50', 'W V55']     ),  # Included W ALL because official C&C records do, and so many age groups
                ('3000SC',     True,        2,        True,     'T',                        ['M ALL']     ),
                ('3000SCW',    True,        2,        True,     'T',                        ['W ALL']     ),
                ('MileW',      True,        2,        True,     'T',                        []     ), # Walks not shown as runbritain dropdowns but are supported
                ('1500W',      True,        2,        True,     'T',                        []     ),
                ('2000W',      True,        2,        True,     'T',                        []     ),
                ('3000W',      True,        2,        True,     'T',                        []     ),
                ('5000W',      True,        2,        True,     'T',                        []     ),
                # 1MW, 2MW, 5KW, 10MW, 20MW events not supported by runbritain
                ('5MW',        True,        2,        True,     'R',                        []     ),
                ('10000W',     True,        2,        True,     'T',                        []     ),
                ('10KW',       True,        2,        True,     'R',                        []     ),
                ('70HU13W',    True,        1,        True,     'T',                        ['W U13']     ),
                ('75HU13M',    True,        1,        True,     'T',                        ['M U13']     ),
                ('75HU15W',    True,        1,        True,     'T',                        ['W U15']     ),
                ('80HU15M',    True,        1,        True,     'T',                        ['M U15']     ),
                ('80HU17W',    True,        1,        True,     'T',                        ['W U17']     ),
                ('80HW40',     True,        1,        True,     'T',                        ['W V40', 'W V45']     ), # checked runbritain for hurdles age groups
                ('80HW50',     True,        1,        True,     'T',                        ['W V50', 'W V55']     ),
                ('80HW60',     True,        1,        True,     'T',                        ['W V60', 'W V65']     ),
                ('80HW70',     True,        1,        True,     'T',                        ['W V70', 'W V75']     ),
                ('100HW',      True,        1,        True,     'T',                        ['W ALL']     ), # Hurdles not shown on runbritain but do work
                ('100HM50',    True,        1,        True,     'T',                        ['M V50', 'M V55']     ),
                ('100HU17M',   True,        1,        True,     'T',                        ['M U17']     ),
                ('110HU20M',   True,        1,        True,     'T',                        ['M U20']     ),
                ('110H',       True,        1,        True,     'T',                        ['M ALL']     ),
                ('110HM35',    True,        1,        True,     'T',                        ['M V35', 'M V40', 'M V45']     ),
                ('300HW',      True,        1,        True,     'T',                        ['W U17', 'W V50', 'W V55']     ),
                ('300HW60',    True,        1,        True,     'T',                        ['W V60', 'W V65']     ),
                ('400H',       True,        1,        True,     'T',                        ['M ALL']     ),
                ('400HW',      True,        1,        True,     'T',                        ['W ALL']     ),
                ('400HU17M',   True,        1,        True,     'T',                        ['M U17']     ),
                ('400HM50',    True,        1,        True,     'T',                        ['M V50', 'M V55']     ),
                ('4x100',      True,        1,        True,     'T',                        []     ),
                ('4x200',      True,        1,        True,     'T',                        []     ),
                ('4x400',      True,        2,        True,     'T',                        []     ),
                ('HJ',         False,       1,        True,     'F',                        []     ),
                ('PV',         False,       1,        True,     'F',                        []     ),
                ('LJ',         False,       1,        True,     'F',                        []     ),
                ('TJ',         False,       1,        True,     'F',                        []     ),
                ('SP2K',       False,       1,        True,     'F',                        ['W V75', 'W V80']     ),
                ('SP2.72K',    False,       1,        True,     'F',                        ['W U13']     ),
                ('SP3K',       False,       1,        True,     'F',                        ['M U13', 'M V80', 'W U15', 'W U17', 'W V50', 'W V55', 'W V60', 'W V65', 'W V70']     ),
                ('SP3.25K',    False,       1,        True,     'F',                        ['W U13', 'M U13', 'W U15', 'M U15', 'W U17', 'W V55']     ),  # In C&C records, not current weight
                ('SP4K',       False,       1,        True,     'F',                        ['W ALL', 'M U15', 'M V70', 'M V75']     ),
                ('SP5K',       False,       1,        True,     'F',                        ['M U17', 'M V60', 'M V65']     ),
                ('SP6K',       False,       1,        True,     'F',                        ['M U20', 'M V50', 'M V55']     ),
                ('SP7.26K',    False,       1,        True,     'F',                        ['M ALL']     ),
                ('DT0.75K',    False,       1,        True,     'F',                        ['W U13', 'W V75', 'W V80', 'W V85']     ),    # Got to here, DTs not yet done for age cats
                ('DT1K',       False,       1,        True,     'F',                        ['W ALL', 'M U13', 'M V60', 'M V65', 'M V70', 'M V75' 'M V80', 'M V85']     ),
                ('DT1.25K',    False,       1,        True,     'F',                        ['M U15']     ),
                ('DT1.5K',     False,       1,        True,     'F',                        ['M U17', 'M V50', 'M V55']     ),
                ('DT1.75K',    False,       1,        True,     'F',                        ['M U20']     ),
                ('DT2K',       False,       1,        True,     'F',                        ['M ALL']     ),
                ('HT2K',       False,       1,        True,     'F',                        ['W V75', 'W V80', 'W V85']     ),
                ('HT3K',       False,       1,        True,     'F',                        ['W U13', 'W U15', 'W U17', 'W V50', 'W V55', 'W V60', 'W V65', 'W V70', 'M U13', 'M V80', 'M V85']     ),
                ('HT4K',       False,       1,        True,     'F',                        ['W ALL', 'M U15', 'M V70', 'M V75']     ),
                ('HT5K',       False,       1,        True,     'F',                        ['M U17', 'M V60', 'M V65']     ),
                ('HT6K',       False,       1,        True,     'F',                        ['M U20', 'M V50', 'M V55']     ),
                ('HT7.26K',    False,       1,        True,     'F',                        ['M ALL']     ),
                ('WT4K',       False,       1,        True,     'F',                        ['W V75', 'W V80', 'W V85']     ),
                ('WT5.45K',    False,       1,        True,     'F',                        ['W U13', 'W U15', 'W V60', 'W V65', 'W V70', 'M U13', 'M V80', 'M V85']     ),
                ('WT7.26K',    False,       1,        True,     'F',                        ['W U17', 'W V50', 'W V55', 'M U15', 'M V70', 'M V75']     ),
                ('WT9.08K',    False,       1,        True,     'F',                        ['W ALL', 'M U17', 'M V60', 'M V65']     ),
                ('WT11.34K',   False,       1,        True,     'F',                        ['M U20', 'M V50', 'M V55']     ),
                ('WT15.88K',   False,       1,        True,     'F',                        ['M ALL']     ),
                ('JT400',      False,       1,        True,     'F',                        ['W U13', 'W V75', 'W V80', 'W V85', 'M U13', 'M V80', 'M V85', 'M V90']     ),
                ('JT500',      False,       1,        True,     'F',                        ['W U15', 'W U17', 'W V50', 'W V55', 'W V60', 'W V65', 'W V70', 'M V70', 'M V75']     ),
                ('JT600',      False,       1,        True,     'F',                        ['W ALL', 'M U15', 'M V60', 'M V65']     ),
                ('JT600PRE86', False,       1,        False,    'F',                        ['W ALL']     ), # Invented here for historical records
                ('JT600PRE99', False,       1,        False,    'F',                        ['W ALL']     ), # Invented here for historical records
                ('JT700',      False,       1,        True,     'F',                        ['M U17', 'M V50', 'M V55']     ),
                ('JT800',      False,       1,        True,     'F',                        ['M ALL']     ),
                ('JT800PRE86', False,       1,        False,    'F',                        ['M ALL']     ), # Invented here for historical records
                ('Minithon',   False,       1,        False,    'M',                        ['W U13', 'M U13']     ), # from C&C club records but not in Po10
                ('Oct',        False,       1,        False,    'M',                        []     ), # from C&C club records but not in Po10
                ('PenU13W',    False,       1,        True,     'M',                        ['W U13']     ),
                ('PenU13M',    False,       1,        True,     'M',                        ['M U13']     ),
                ('PenU15W',    False,       1,        True,     'M',                        ['W U15']     ),
                ('PenU15M',    False,       1,        True,     'M',                        ['M U15']     ),
                ('PenU17W',    False,       1,        True,     'M',                        ['W U17']     ),
                ('PenU17M',    False,       1,        True,     'M',                        ['M U17']     ),
                ('PenU20M',    False,       1,        True,     'M',                        ['M U20']     ),
                ('PenW',       False,       1,        True,     'M',                        ['W ALL']     ),
                ('Pen',        False,       1,        True,     'M',                        ['M ALL']     ),
                ('PenW35',     False,       1,        True,     'M',                        ['W V35']     ), # Outdoor
                ('PenM35',     False,       1,        True,     'M',                        ['M V35']     ), # Outdoor
                ('PenW40',     False,       1,        True,     'M',                        ['W V40']     ),
                ('PenM40',     False,       1,        True,     'M',                        ['M V40']     ),
                ('PenW45',     False,       1,        True,     'M',                        ['W V45']     ),
                ('PenM45',     False,       1,        True,     'M',                        ['M V45']     ),
                ('PenW50',     False,       1,        True,     'M',                        ['W V50']     ),
                ('PenM50',     False,       1,        True,     'M',                        ['M V50']     ),
                ('PenW55',     False,       1,        True,     'M',                        ['W V55']     ),
                ('PenM55',     False,       1,        True,     'M',                        ['M V55']     ),
                ('PenW60',     False,       1,        True,     'M',                        ['W V60']     ),
                ('PenM60',     False,       1,        True,     'M',                        ['M V60']     ),
                ('PenW65',     False,       1,        True,     'M',                        ['W V65']     ),
                ('PenM65',     False,       1,        True,     'M',                        ['M V65']     ),
                ('PenW70',     False,       1,        True,     'M',                        ['W V70']     ),
                ('PenM70',     False,       1,        True,     'M',                        ['M V70']     ),
                ('PenW75',     False,       1,        True,     'M',                        ['W V75']     ),
                ('PenM75',     False,       1,        True,     'M',                        ['M V75']     ),
                ('PenW80',     False,       1,        True,     'M',                        ['W V80']     ),
                ('PenM80',     False,       1,        True,     'M',                        ['M V80']     ),
                ('PenI',       False,       1,        True,     'M',                        ['M ALL']     ), # Indoor
                ('PenIW35',    False,       1,        True,     'M',                        ['W V35']     ),
                ('PenIM35',    False,       1,        True,     'M',                        ['M V35']     ),
                ('PenIW40',    False,       1,        True,     'M',                        ['W V40']     ),
                ('PenIM40',    False,       1,        True,     'M',                        ['M V40']     ),
                ('PenIW45',    False,       1,        True,     'M',                        ['W V45']     ),
                ('PenIM45',    False,       1,        True,     'M',                        ['M V45']     ),
                ('PenIW50',    False,       1,        True,     'M',                        ['W V50']     ),
                ('PenIM50',    False,       1,        True,     'M',                        ['M V50']     ),
                ('PenIW55',    False,       1,        True,     'M',                        ['W V55']     ),
                ('PenIM55',    False,       1,        True,     'M',                        ['M V55']     ),
                ('PenIW60',    False,       1,        True,     'M',                        ['W V60']     ),
                ('PenIM60',    False,       1,        True,     'M',                        ['M V60']     ),
                ('PenIW65',    False,       1,        True,     'M',                        ['W V65']     ),
                ('PenIM65',    False,       1,        True,     'M',                        ['M V65']     ),
                ('PenIW70',    False,       1,        True,     'M',                        ['W V70']     ),
                ('PenIM70',    False,       1,        True,     'M',                        ['M V70']     ),
                ('PenIM75',    False,       1,        True,     'M',                        ['M V75']     ), # No women's code beyond V70
                ('PenIM80',    False,       1,        True,     'M',                        ['M V80']     ),
                # Skipping weights pentathlon junior events, never seem to happen these days and no prev C&C records
                ('PenWtW',     False,       1,        True,     'F',                        ['W ALL']     ),
                ('PenWt',      False,       1,        True,     'F',                        ['M ALL']     ),
                ('PenWtW35',   False,       1,        True,     'F',                        ['W V35']     ),
                ('PenWtM35',   False,       1,        True,     'F',                        ['M V35']     ),
                ('PenWtW40',   False,       1,        True,     'F',                        ['W V40']     ),
                ('PenWtM40',   False,       1,        True,     'F',                        ['M V40']     ),
                ('PenWtW45',   False,       1,        True,     'F',                        ['M V45']     ),
                ('PenWtM45',   False,       1,        True,     'F',                        ['W V45']     ),
                ('PenWtW50',   False,       1,        True,     'F',                        ['W V50']     ),
                ('PenWtM50',   False,       1,        True,     'F',                        ['M V50']     ),
                ('PenWtW55',   False,       1,        True,     'F',                        ['W V55']     ),
                ('PenWtM55',   False,       1,        True,     'F',                        ['M V55']     ),
                ('PenWtW60',   False,       1,        True,     'F',                        ['W V60']     ),
                ('PenWtM60',   False,       1,        True,     'F',                        ['M V60']     ),
                ('PenWtW65',   False,       1,        True,     'F',                        ['W V65']     ),
                ('PenWtM65',   False,       1,        True,     'F',                        ['M V65']     ),
                ('PenWtW70',   False,       1,        True,     'F',                        ['W V70']     ),
                ('PenWtM70',   False,       1,        True,     'F',                        ['M V70']     ),
                ('PenWtW75',   False,       1,        True,     'F',                        ['W V75']     ),
                ('PenWtM75',   False,       1,        True,     'F',                        ['M V75']     ),
                ('PenWtW80',   False,       1,        True,     'F',                        ['W V80']     ),
                ('PenWtM80',   False,       1,        True,     'F',                        ['M V80']     ),
                ('HepU17W',    False,       1,        True,     'M',                        ['W U17']     ),
                ('HepW',       False,       1,        True,     'M',                        ['W ALL']     ),
                ('HepW35',     False,       1,        True,     'M',                        ['W V35']     ),
                ('HepW40',     False,       1,        True,     'M',                        ['W V35']     ),
                ('HepW45',     False,       1,        True,     'M',                        ['W V40']     ),
                ('HepW50',     False,       1,        True,     'M',                        ['W V50']     ),
                ('HepW55',     False,       1,        True,     'M',                        ['W V55']     ),
                ('HepW60',     False,       1,        True,     'M',                        ['W V60']     ),
                ('HepW65',     False,       1,        True,     'M',                        ['W V65']     ),
                ('DecU17M',    False,       1,        True,     'M',                        ['M U17']     ),
                ('DecU20M',    False,       1,        True,     'M',                        ['M U20']     ),
                ('Dec',        False,       1,        True,     'M',                        ['M ALL']     ),
                ('DecM35',     False,       1,        True,     'M',                        ['M V35']     ),
                ('DecM40',     False,       1,        True,     'M',                        ['M V40']     ),
                ('DecM45',     False,       1,        True,     'M',                        ['M V45']     ),
                ('DecM50',     False,       1,        True,     'M',                        ['M V50']     ),
                ('DecM55',     False,       1,        True,     'M',                        ['M V55']     ),
                ('DecM60',     False,       1,        True,     'M',                        ['M V60']     ),

 ]

known_events_lookup = {}
for (event, smaller_better, numbers, runbritain, type, categories) in known_events:
    known_events_lookup[event] = (smaller_better, numbers, runbritain, type, categories)

# PowerOf10 age categories (usable on club page)
powerof10_categories = ['ALL', 'U13', 'U15', 'U17', 'U20']

# Runbritain age categories
runbritain_categories = [ # name       min  max years old
                          ('ALL',        0,    0),
                          # ('Disability', 0,    0),     # Runbritain category but always zero results
                          ('U13',        1,   12),
                          ('U15',        13,  14),
                          ('U17',        15,  16),
                          ('U20',        17,  19),
                          ('U23',        20,  22),
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

age_category_lookup = {}
for age in range(1, 120):
    for (category, min_age, max_age) in runbritain_categories:
        if age >= min_age and age <= max_age:
            age_category_lookup[age] = category
            break
    if age not in age_category_lookup:
        age_category_lookup[age] = 'SEN'

# Read from file later
# Indexed by Po10 event code, age/gender category
ea_pb_award_score = {}
num_ea_pb_levels  = 9

class EaPbAwardScoreSet():
    def __init__(self, bucket, event, category, level_scores):
        self.bucket = bucket # e.g. Sprint or Endurance
        self.event = event # Po10 event e.g. '400H'
        self.category = category # age/gender e.g. 'W U20'
        self.level_scores = level_scores # list of 9 numeric values for level 1 to level 9



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
    # Also various special cases from legacy club records

    perf = perf.lower()

    # Support invalidation of records e.g. so that athletes who have left
    # club but still listed as members on Po10 can have performances removed
    invalid = False
    if 'invalid' in perf:
        invalid = True
        perf = perf.replace('invalid', '')

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

    return total_score, decimal_places, original_special, invalid


def construct_performance(event, gender, category, perf, name, url, date, fixture_name, fixture_url,
                          source, age_grade='0.0', age=0):
    score, original_dp, original_special, invalid = make_numeric_score_from_performance_string(perf)
    wava = float(age_grade)
    perf = Performance(event, score, category, gender, original_special, original_dp, name, url, 
                        date, fixture_name, fixture_url, source, wava=wava, age=age, invalid=invalid)
    return perf


def source_pref_score(source):
    if source.startswith('Po10'):
        score = 2
    elif source.startswith('Runbritain'):
        score = 1
    else:
        # E.g. legacy spreadsheets now not preferred
        score = 0
    
    return score


def process_performance_cat_and_all(perf, types, collection_choice, year='ALL'):
    """Our manual entries from spreadsheets
    should be considered for overall records even if they are noted for an age group,
    if that event is one that seniors do.
    Also now catching if powerof10 or runbritain provide an age group performance
    that's not all captured for ALL, but only found one instance of that in practice
    (W 300 from 2005). It does result in powerof10 replacing some runbritain
    equivalens though."""

    process_performance(perf, types, collection_choice, year)

    saved_category = perf.category
    if saved_category != 'ALL' and event_relevant_to_category(perf.event, perf.gender, 'ALL'):
        # Also consider if this might be an outright record, not just in this age group
        perf.category = 'ALL' # Cheeky trick to avoid reconstructing for different category
        process_performance(perf, types, collection_choice, year)
        perf.category = saved_category


def process_performance(perf, types, collection_choice, year='ALL'):
    """Add performance to overall and category record tables if appropriate,
      while respecting the max size of those tables"""
    
    known_event = known_events_lookup.get(perf.event)
    if not known_event:
        print(f'Event not in known events: {perf.event}')
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
        max_records = max_records_all if perf.category == 'ALL' else max_records_age_group
        smaller_score_better = known_events_lookup[perf.event][0]
        compare_field = 'score'
    elif collection_choice == 'wava':
        collection = wava
        if perf.event not in collection:
            collection[perf.event] = {}
        if year not in collection[perf.event]:
            collection[perf.event][year] = []
        record_list = collection[perf.event][year]
        max_records = max_wavas_all if year == 'ALL' else max_wavas_year
        smaller_score_better = False # WAVA bigger the better always
        compare_field = 'wava'
    elif collection_choice == 'ea_pb':
        if perf.event not in ea_pb_award_score:
            # No score tables loaded or event doesn't fit scheme
            return
        ea_pb_event = ea_pb_award_score[perf.event]
        gender_age_cat = perf.gender + " " + perf.category
        if gender_age_cat not in ea_pb_event:
            # No score defined
            return
        ea_pb_obj = ea_pb_event[gender_age_cat]
        smaller_score_better = known_events_lookup[perf.event][0] # Event time/distance/height
        perf.ea_pb_score = calculate_ea_pb_score(ea_pb_obj, perf.score, smaller_score_better)
        smaller_score_better = False # Now EA PB Score not event time/distance/height
        compare_field = 'ea_pb_score'
        collection = ea_pb
        if ea_pb_obj.bucket not in collection:
            collection[ea_pb_obj.bucket] = {}
        if year not in collection[ea_pb_obj.bucket]:
            collection[ea_pb_obj.bucket][year] = []
        record_list = collection[ea_pb_obj.bucket][year]
        max_records = max_ea_pbs_all if year == 'ALL' else max_ea_pbs_year
    else:
        raise ValueError(f"Unexpected collection_choice {collection_choice}")

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
                        if getattr(perf, 'invalid', False): # Po10 caches may predate this
                            # Manual anti-record to delete entry we don't want
                            del existing_perf_list[perf_idx]
                            tie_same_name_managed = True
                            break
                        existing_source_score = source_pref_score(existing_perf.source)
                        this_source_score = source_pref_score(perf.source)
                        if existing_source_score >= this_source_score:
                            # E.g. don't add Runbritain score if already there from Po10
                            tie_same_name_managed = True
                            break
                        else:
                            # E.g. replace existing Runbritain one with Po10, or prefer newer file
                            # assuming provided in ascending date order
                            existing_perf_list[perf_idx] = perf
                            tie_same_name_managed = True
                            break
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


def calculate_ea_pb_score(ea_pb_obj, score, smaller_score_better):
    # The England Athletics PB Awards scheme gives an integer score to different
    # performance attainments, but we need a continuous (decimal) score for close
    # comparisons, so have to intepolate between the levels defined

    worst_defined = ea_pb_obj.level_scores[0]
    best_defined = ea_pb_obj.level_scores[num_ea_pb_levels - 1]

    if smaller_score_better:
        if score > worst_defined:
            # Below Level 1
            reciprocal_score = 1.0 / score # So long time gives small number
            reciprocal_worst = 1.0 / worst_defined # Normalised to 1.0 for Level 1
            ea_score = reciprocal_score / reciprocal_worst
        elif score <= best_defined:
            # At or above Level 9, ramp to "Level 10" at safe limit performance
            limit_score = best_defined * (1 - ea_pb_limit_perf_fraction)
            fraction_to_limit = (best_defined - score) / (best_defined - limit_score)
            ea_score = num_ea_pb_levels + fraction_to_limit
        else:
            for lo_idx in range(0, num_ea_pb_levels - 1):
                hi_idx = lo_idx + 1
                if (score <= ea_pb_obj.level_scores[lo_idx] and 
                       score > ea_pb_obj.level_scores[hi_idx]):
                    delta = ea_pb_obj.level_scores[lo_idx] - ea_pb_obj.level_scores[hi_idx]
                    diff = ea_pb_obj.level_scores[lo_idx] - score
                    ea_score = (lo_idx + 1) + diff / delta
    else:
        if score < worst_defined:
            # Below Level 1, do smooth ramp to "Level 0" at zero
            ea_score = score / worst_defined
        elif score >= best_defined:
            # At or above Level 9, ramp to "Level 10" at a safe limit performance
            limit_score = best_defined * (1 + ea_pb_limit_perf_fraction)
            fraction_to_limit = (score - best_defined) / (limit_score - best_defined)
            ea_score = num_ea_pb_levels + fraction_to_limit
        else:
            for lo_idx in range(0, num_ea_pb_levels - 1):
                hi_idx = lo_idx + 1
                if (score >= ea_pb_obj.level_scores[lo_idx] and 
                       score < ea_pb_obj.level_scores[hi_idx]):
                    delta = ea_pb_obj.level_scores[hi_idx] - ea_pb_obj.level_scores[lo_idx]
                    diff = score - ea_pb_obj.level_scores[lo_idx]
                    ea_score = (lo_idx + 1) + diff / delta

    return ea_score


def process_po10_wava(reqd_perf, performance_cache, types, rebuild_wava):
    """Add a performance to WAVA tables"""

    athlete_id_match = re.search(r'athleteid=([0-9]+)', reqd_perf.athlete_url)
    athlete_id = athlete_id_match.group(1)

    request_params = {'athleteid'   : athlete_id,
                      'viewby'      : 'agegraded'}

    url = powerof10_root_url + '/athletes/profile.aspx'
    cache_key = make_cache_key(url, request_params)

    if athlete_id in wava_athlete_ids_done:
        # We already have all performances for this athlete, even if cache rebuilt this time
        use_cache = True
    else:
        use_cache = not rebuild_wava
        wava_athlete_ids_done[athlete_id] = True

    if use_cache:
        perf_list = performance_cache.get(cache_key, None)
    else:
        perf_list = None

    report_string_base = f'PowerOf10 WAVA list for {reqd_perf.athlete_name} ID {athlete_id} '
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
            process_one_athlete_results_table(reqd_perf, rows, perf_list)

        performance_cache[cache_key] = perf_list

    for perf in perf_list:
        # Only match performance of interest this time, as athlete may have
        # performances logged when running for a different club
        if reqd_perf.date != perf.date:
            continue
        else:
            # Found performance we were looking for this time
            year = get_perf_year(perf.date)
            process_performance(perf, types, 'wava', 'ALL')
            process_performance(perf, types, 'wava', str(year))
            performance_count['Po10-WAVA'] += 1
            break


# PowerOf10 dates always have form "1 Jan 1980" or "11 Jan 1989"
regex_po10_date = re.compile(r'([0-9][0-9]?) ([A-Z][a-z][a-z]) ([0-9][0-9])')
regex_4digits = re.compile(r'([0-9]{4})')

def get_perf_year(perf_date_str):
    # Return useful numeric year from whatever string we have
    # Can't make this Performance method now because of cached objects on disk

    # Mostly Po10/RunBritain
    match_obj = regex_po10_date.match(perf_date_str)
    if match_obj:
        year_2digit = match_obj.group(3) # e.g. 24 for 2024
        return 2000 + int(year_2digit)
    # Manual records should at least have 4-digit number date
    match_obj = regex_4digits.search(perf_date_str)
    if match_obj:
        return int(match_obj.group(1))
    # If we get to here, we found nothing useful
    print(f"WARNING: unparseable date found: {perf_date_str}")
    return 1900


def process_one_athlete_results_table(example_perf, rows, perf_list):
    """Go through table of performances for a single athlete, especially intended
    to pick out age grades"""
    
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
        fixture_url = fixture_url.replace('..', '') # in these pages, starts with relative path
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
                                  rebuild_cache, first_claim_only, types):
    """Process gender/age category rankings from powerof10 for specified year
       for ALL events in one go (returned page has table per event, different
       from runbritain)"""
    
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
        process_performance_cat_and_all(perf, types, 'record')
        process_performance(perf, types, 'ea_pb', year='ALL')
        process_performance(perf, types, 'ea_pb', year=str(year))
        performance_count['Po10'] += 1


def make_cache_key(url, request_params):
    """Make unique string for web request to use as dict key for cache of previous
    web results"""

    cache_key = url + '?'
    for key, value in request_params.items(): # ordered as inserted in Python 3.6+
        cache_key += key + '=' + value + '&' # will leave trailing & but who cares
    return cache_key


def process_one_runbritain_year_gender(club_id, year, gender, category, event, performance_cache,
                                        rebuild_cache, first_claim_only, types, do_wava, rebuild_wava):
    """Get rankings for a single event, age group, gender etc from runbritain; this is
    where such detailed rankings tables are fetched from when requested from powerof10."""
    
    request_params = {'clubid'         : str(club_id),
                      'sex'            : gender,
                      'year'           : str(year),
                      'event'          : event,
                      'firstclaimonly' : 'y' if first_claim_only else 'n',
                      'limit'          : 'n'      } # Otherwise miss slower performances, undocumented option

    (min_age, max_age) = runbritain_category_lookup[category]
    if max_age <= 22:
        # Use category name for junior/youth categories where season age may not
        # match birthday age
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
        process_performance_cat_and_all(perf, types, 'record')
        process_performance(perf, types, 'ea_pb', year='ALL')
        process_performance(perf, types, 'ea_pb', year=year)
        performance_count['Runbritain'] += 1
        if do_wava and perf.event in wava_events:
            # Done in runbritain processing because po10 overall (all events)
            # rankings by year don't reliably include 5K
            process_po10_wava(perf, performance_cache, types, rebuild_wava)



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


def output_records(output_file, first_year, last_year, club_id, do_po10, do_runbritain,
                   input_files, club_name):

    stylesheet_ref_head="""
<head>
  <!-- This stylesheet font requested by Wing Wong to match C&C site 07Jun2025 -->
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Mulish:ital,wght@0,200..1000;1,200..1000" type="text/css" />
  <style>
    body {
      font-family: "Mulish", sans-serif;
    }
  </style>
</head>

"""

    header_part = [
                    '<html>\n',
                    stylesheet_ref_head,
                    '<body>\n']
                    # Title with club name could be put back here as option for standalone records esp for other clubs
    
    sources_part = [f'<h2><a name="sources" />Details and Sources for {club_name} Club Records</h2>\n',
                    f'<p>Autogenerated  on {datetime.date.today()} from:</p>\n',
                    f'<ul>\n']
    if do_po10:
        sources_part.append(f'<li><a href="{powerof10_root_url}/clubs/club.aspx?clubid={club_id}">PowerOf10 club page</a>  {first_year} - {last_year}</li>\n')
    if do_runbritain:
        sources_part.append(f'<li><a href="{runbritain_root_url}/rankings/rankinglist.aspx">runbritain rankings</a>  {first_year} - {last_year}</li>\n')
    for input_file in input_files:
        sources_part.append(f'<li>Local file: {input_file}</li>\n')
    sources_part.append('</ul>\n\n')
    sources_part.append(f'<p>Outputting maximum {max_records_all} places overall per event and {max_records_age_group} per age group.</p>\n')
    sources_part.append(f'<p>Outputting maximum {max_wavas_all} places for all time age graded and {max_wavas_year} per year, per event.</p>\n')
    sources_part.append(f'<p>Count of performances processed...')
    for type in performance_count.keys():
        sources_part.append(f' {type}: {performance_count[type]}')
    sources_part.append(f'</p>\n')

    main_contents_part = [] # "Contents" subheading could be put back here as option for standalone records
    
    main_contents_part.append(f'<table {common_table_attribs}>\n')

    complete_bulk_part = []

    new_records_last_year = [] # E.g. if running in 2025, best records obtained in 2024
    new_records_this_year = [] # E.g. if running in 2025, best records obtained in 2025
    last_complete_year = last_year - 1 # e.g. if running analysis in mid 2025, we have whole of 2024

    year_keys = ['ALL']
    for year in range(last_year, first_year - 1, -1):
        year_keys.append(str(year))

    for (category, _, _) in runbritain_categories:
        if category not in record: continue
        main_contents_part.append('<tr>\n')
        for gender in ['W', 'M']:
            section_bulk_part = []
            section_contents_part = []
            anchor = f'category_{gender}_{category}'
            subtitle = f'Category: {gender} {category}'
            main_contents_part.append(f'<td><center><b><a href="#{anchor}">{subtitle}</a></b></center></td>\n')
            section_contents_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
            section_contents_part.append('<p>Jump to: \n')
            for (event, _, _, _, _, _) in known_events:
                if event not in record[category]: continue
                record_list = record[category][event].get(gender)
                if not record_list: continue
                anchor = f'{event}_{gender}_{category}'.lower()
                subtitle = f'{event} {gender} {category}'
                section_contents_part.append(f'<em><a href="#{anchor}">...{subtitle}</a></em>\n')
                section_bulk_part.append(f'<h3><a name="{anchor}" />Records for {subtitle}</h3>\n\n')
                output_record_table(section_bulk_part, record_list, 'record')
                add_best_record_if_new_this_year(new_records_this_year, record_list, last_year, 'RECORD')
                add_best_record_if_new_this_year(new_records_last_year, record_list, last_complete_year, 'RECORD')
            section_contents_part.append('</p>\n\n')
            complete_bulk_part.extend(section_contents_part)
            complete_bulk_part.extend(section_bulk_part)
        main_contents_part.append('</tr>\n')

    for bucket in ea_pb.keys():
        section_bulk_part = []
        section_contents_part = []
        anchor = f'ea_pb_{bucket}'.lower()
        subtitle = f'England Athletics PB Awards scheme: {bucket} [experimental]'
        main_contents_part.append(f'<tr>\n<td colspan="2"><center><b><a href="#{anchor}">{subtitle}</a></b></center</td>\n</tr>\n')
        section_contents_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
        section_contents_part.append('<p>Jump to: \n')
        for year_key in year_keys:
            if year_key not in ea_pb[bucket]:
                continue
            record_list = ea_pb[bucket][year_key]
            anchor = f'ea_pb_{bucket}_{year_key}'.lower()
            subtitle = year_key
            section_contents_part.append(f'<em><a href="#{anchor}">...{subtitle}</a></em>\n')
            section_bulk_part.append(f'<h3><a name="{anchor}" />England Athletics PB Awards ({bucket}) year: {subtitle}</h3>\n\n')
            output_record_table(section_bulk_part, record_list, 'ea_pb')
            if year_key == 'ALL':
                add_best_record_if_new_this_year(new_records_this_year, record_list, last_year, 'EA PB')
                add_best_record_if_new_this_year(new_records_last_year, record_list, last_complete_year, 'EA PB')
        section_contents_part.append('</p>\n\n')
        complete_bulk_part.extend(section_contents_part)
        complete_bulk_part.extend(section_bulk_part)

    for event in wava_events:
        section_bulk_part = []
        section_contents_part = []
        if event not in wava:
            continue
        anchor = f'wava_{event}'.lower()
        subtitle = f'Age Grade: {event}'
        main_contents_part.append(f'<tr>\n<td colspan="2"><center><b><a href="#{anchor}">{subtitle}</a></b></center</td>\n</tr>\n')
        section_contents_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
        section_contents_part.append('<p>Jump to: \n')
        for year_key in year_keys:
            if year_key not in wava[event]:
                continue
            record_list = wava[event][year_key]
            anchor = f'wava_{event}_{year_key}'.lower()
            subtitle = year_key
            section_contents_part.append(f'<em><a href="#{anchor}">...{subtitle}</a></em>\n')
            section_bulk_part.append(f'<h3><a name="{anchor}" />Age Grade {event} year: {subtitle}</h3>\n\n')
            output_record_table(section_bulk_part, record_list, 'wava')
            if year_key == 'ALL':
                add_best_record_if_new_this_year(new_records_this_year, record_list, last_year, 'WAVA')
                add_best_record_if_new_this_year(new_records_last_year, record_list, last_complete_year, 'WAVA')
        section_contents_part.append('</p>\n\n')
        complete_bulk_part.extend(section_contents_part)
        complete_bulk_part.extend(section_bulk_part)

    if new_records_this_year:
        section_bulk_part = []
        anchor = f'new_best_{last_year}'
        subtitle = f'New (or equalled) records achieved so far this calendar year: {last_year}'
        main_contents_part.append(f'<tr>\n<td colspan="2"><center><b><a href="#{anchor}">{subtitle}</a></b></center</td>\n</tr>\n')
        section_bulk_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
        output_record_table(section_bulk_part, new_records_this_year, 'new_in_year')
        complete_bulk_part.extend(section_bulk_part)

    if new_records_last_year:
        section_bulk_part = []
        anchor = f'new_best_{last_complete_year}'
        subtitle = f'New (or equalled) records achieved last calendar year: {last_complete_year}'
        main_contents_part.append(f'<tr>\n<td colspan="2"><center><b><a href="#{anchor}">{subtitle}</a></b></center</td>\n</tr>\n')
        section_bulk_part.append(f'<h2><a name="{anchor}" />{subtitle}</h2>\n\n')
        section_bulk_part.append('<p><em>Note: skips records where same athlete has bettered record <b>this</b> year in same event.</em></p>')
        output_record_table(section_bulk_part, new_records_last_year, 'new_in_year')
        complete_bulk_part.extend(section_bulk_part)

    main_contents_part.append('</table>\n\n')

    tail_part = ['</body>\n',
                '</html>\n']

    with open(output_file, 'wt') as fd:
        for part in [header_part, main_contents_part, complete_bulk_part, sources_part, tail_part]:
            for line in part:
                fd.write(line)


def output_record_table(bulk_part, record_list, type):
    if len(record_list) < 1:
        return
    
    bulk_part.append(f'<table {common_table_attribs}>\n')
    bulk_part.append('<tr>\n')
    # Different columns depending on type of table
    show_rank_col     = (type != 'new_in_year')
    show_reason_col   = (type == 'new_in_year')
    show_wava_col     = (type in {'wava', 'new_in_year'})
    show_ea_pb_col    = (type in {'ea_pb', 'new_in_year'})
    show_category_col = (type in {'wava', 'ea_pb', 'new_in_year'})
    show_event_col    = (type in {'ea_pb', 'new_in_year'})
    if show_reason_col:
        bulk_part.append('<td><center><b>Type</b></center></td>')
    if show_rank_col:
        bulk_part.append('<td><center><b>Rank</b></center></td>')
    if show_wava_col:
        bulk_part.append('<td><center><b>Age Grade %</b></center></td>')
    if show_ea_pb_col:
        bulk_part.append('<td><center><b>EA PB Score</b></center></td>')
    if show_category_col:
        bulk_part.append('<td><center><b>Category</b></center></td>')
    if show_event_col:
        bulk_part.append('<td><center><b>Event</b></center></td>')
    bulk_part.append('<td><center><b>Performance</b></center></td><td><center><b>Athlete</b></center></td><td><center><b>Date</b></center></td><td><center><b>Fixture</b></center><td><center><b>Source</b></center></td>\n')
    bulk_part.append('</tr>\n')
    for idx, perf_list in enumerate(record_list):
        for perf_idx, perf in enumerate(perf_list): # May be ties with same score or different sources
            if type == 'new_in_year': # tuple for best last year entries
                reason = perf.reason
            else:
                reason = ''
            if perf.original_special:
                score_str = perf.original_special
            else:
                score_str = format_sexagesimal(perf.score, known_events_lookup[perf.event][1], perf.decimal_places)
            bulk_part.append('<tr>\n')
            if show_reason_col:
                bulk_part.append(f'  <td><center>{reason}</center></td>\n')
            if show_rank_col:
                rank_str = f'{idx+1}' if perf_idx == 0 else ''
                bulk_part.append(f'  <td><center>{rank_str}</center></td>\n')
            if show_wava_col:
                if reason == 'WAVA' or type == 'wava':
                    wava_str = '%.2f' % perf.wava
                else:
                    wava_str = '' # Not relevant for this one
                bulk_part.append(f'  <td><center>{wava_str}</td>\n')
            if show_ea_pb_col:
                if reason == 'EA PB' or type == 'ea_pb':
                    ea_pb_str = '%.3f' % perf.ea_pb_score
                else:
                    ea_pb_str = ''
                bulk_part.append(f'  <td><center>{ea_pb_str}</td>\n')
            if show_category_col:
                if reason == 'WAVA' or type == 'wava':
                    # Using category rather than age, for modesty, though age is public on po10!
                    category_str = f'{perf.gender} {age_category_lookup[perf.age]}'
                else:
                    category_str = perf.gender + " " + perf.category
                bulk_part.append(f'  <td><center>{category_str}</td>\n')
            if show_event_col:
                bulk_part.append(f'  <td><center>{perf.event}</td>\n')
            bulk_part.append(f'  <td><center>{score_str}</td>\n')
            if perf.athlete_url:
                bulk_part.append(f'  <td><a href="{perf.athlete_url}" target=”_blank”>{perf.athlete_name}</a></td>\n')
            else:
                bulk_part.append(f'  <td>{perf.athlete_name}</td>\n')
            bulk_part.append(f'  <td>{perf.date}</td>\n')
            if perf.fixture_url:
                bulk_part.append(f'  <td><a href="{perf.fixture_url}" target=”_blank”>{perf.fixture_name}</a></td>\n')
            else:
                bulk_part.append(f'  <td>{perf.fixture_name}</td>\n')
            bulk_part.append(f'  <td>{perf.source}</td>\n')
            bulk_part.append('</tr>\n')
    bulk_part.append('</table>\n\n')


def add_best_record_if_new_this_year(new_records_last_year, record_list, interest_year, reason):
    """Identify if the top of a record table was claimed anew in the year of interest,
     or has only been bettered in a later year if the year of interest is older"""
    
    for perf_list in record_list:
        interest_year_performances = []
        older_year_record_count = 0
        for perf in perf_list: # May be ties with same score or different sources
            perf_year = get_perf_year(perf.date)
            if perf_year == interest_year:
                interest_year_performances.append(perf)
            elif perf_year > interest_year:
                # OK if record obtained in 2025 when we're looking for records in 2024, ignore
                pass
            else:
                # Obtained in older year
                older_year_record_count += 1

        if not interest_year_performances and older_year_record_count > 0:
            # No new best in year of interest, bettered by an older record
            # so stop searching down the table
            return
        elif len(interest_year_performances) > 0:
            # New record achieved one or more times this year
            # Note: Pete Thompson preferred that if a historical record were equalled,
            # we show it "again" this time
            # Also showing historical one(s) alongside it
            for perf in perf_list:
                perf_copy = copy.copy(perf) # So same performance can be RECORD, WAVA etc
                perf_copy.reason = reason
                new_records_last_year.append([perf_copy])
            return
        else:
            # Only found newer records than year of interest so far, so continue searching
            pass


def process_one_club_record_input_file(input_file, types):

    print(f'Processing file: {input_file}')

    _, file_extension = os.path.splitext(input_file)
    if file_extension.lower() != '.xlsx':
        print(f'WARNING: ignoring input file, can only handle .xlsx currently: {input_file}')
        return
    
    workbook = openpyxl.load_workbook(filename=input_file)
    for worksheet in workbook.worksheets:
        process_one_club_record_excel_worksheet(input_file, worksheet, types)


def process_one_club_record_excel_worksheet(input_file, worksheet, types):

    reqd_headings = ['performance', 'date', 'name', 'po10 event', 'gender', 'age code']
    col_renames = {'year' : 'date', 'record holder' : 'name'}
    df = get_table_by_find_check_headings(worksheet, reqd_headings, col_renames)

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
        perf_str = str(perf).strip()
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
        source = 'Historical worksheet: ' + input_file + ':' + worksheet.title
        perf = construct_performance(event, gender, category, perf_str, name, name_url,
                            date, fixture, fixture_url, source)
        process_performance_cat_and_all(perf, types, 'record')
    
        performance_count['File(s)'] += 1


def get_table_by_find_check_headings(worksheet, reqd_headings, renames_dict=None):
    # Extract Pandas dataframe from Excel worksheet, if it has all the expected columns

    print(f'Processing worksheet: {worksheet.title}')
    df = pandas.DataFrame(worksheet.values)

    # We get integers as headings instead of the intended column headings, so find those
    headings_found = False
    for row_idx, row in df.iterrows():
        for col_idx, cell_value in enumerate(row.values):
            if cell_value is None: continue
            if cell_value.lower().strip() == reqd_headings[0]:
                headings_found = True
                break
        if headings_found: break

    if not headings_found:
        print(f'WARNING: could not find "{reqd_headings[0]}" heading, skipping sheet')
        return None
    
    df.columns = df.iloc[row_idx]
    df.drop(df.index[0:row_idx + 1], inplace=True)
    
    # Convert all headings to lower case and strip whitespace
    renames = {}
    for col_name in df.columns:
        if not col_name : continue
        new_col_name = col_name.lower().strip()
        renames[col_name] = new_col_name
    df.rename(columns=renames, inplace=True)
    
    if renames_dict:
        # C&C club records have some column headings that might not suit us for other
        # input lists
        df.rename(columns=renames_dict, inplace=True)

    col_name_list = df.columns.tolist()
    for reqd_heading in reqd_headings:
        if reqd_heading not in col_name_list:
            print(f'Required heading not found (case insensitive), skipping sheet: {reqd_heading}')
            if renames_dict:
                print(f'(That was after some column renames applied: {renames_dict})')
            return None

    return df


def get_po10_club_name(club_id):
    request_params = {'clubid'   : str(club_id)}

    url = powerof10_root_url + '/clubs/club.aspx'
    try:
        page_response = requests.get(url, request_params)
    except requests.exceptions.ConnectionError:
        print('WARNING: failed to get club name from powerof10')
        return 'n/a'

    h2_headings = get_html_content(page_response.text, 'h2')
    if len(h2_headings) != 1:
        print('WARNING: club page no longer has club name as only h2 heading, skipped')
        return 'n/a'

    return h2_headings[0].inner_text


def event_relevant_to_category(event, gender, category):
    # E.g. some hurdles events and throws only relevant to certain age groups/genders

    if event not in known_events_lookup:
        # May be historical event with no modern equivalent
        return False
    
    (_, _, _, _, categories) = known_events_lookup[event]
    if not categories:
        # Blank means relevant for all
        return True
    
    gender_wildcard = gender + ' ALL'
    if False and gender_wildcard in categories:
        # e.g. 2000SCW OK for any women
        return True
    
    category_specific = gender + ' ' + category
    if category_specific in categories:
        return True

    # Otherwise not relevant
    return False


def read_ea_pb_award_score_tables(ea_pb_award_file):
    print(f"Opening file for EA PB Award score tables: {ea_pb_award_file}")
    workbook = openpyxl.load_workbook(filename=ea_pb_award_file)
    if len(workbook.worksheets) > 1:
        raise ValueError(f"Expected an Excel workbook with only one worksheet for EA PB Awards tables")

    level_headings = [f'level {i}' for i in range(1, num_ea_pb_levels + 1)] # 1-based names
    reqd_headings = ['bucket', 'po10 event', 'gender', 'age code']
    reqd_headings.extend(level_headings)

    df = get_table_by_find_check_headings(workbook.worksheets[0], reqd_headings)
    if df is None:
        raise ValueError(f"Unable to read EA PB scores from {ea_pb_award_file}")
    
    rows_completed = 0
    for row_idx, row in df.iterrows():
        # Can get None obj references as well as empty strings
        excel_row_number = row_idx + 1
        bucket = row['bucket']
        bucket = str(bucket).strip() if bucket else '' # E.g. "Throw"
        event = row['po10 event']
        event = str(event).strip() if event else ''
        if not bucket and not event:
            # Assume blank row, quietly ignore
            continue
        if not bucket:
            print(f'WARNING: EA "bucket" missing at row {excel_row_number}')
            continue
        if not event:
            print(f'WARNING: Po10 event code missing at row {excel_row_number}')
            continue
        gender = row['gender']
        gender = str(gender).upper().strip() if gender else ''
        pb_genders = ['M', 'W', 'X']
        if gender not in pb_genders:
            print(f'WARNING: gender not one of {pb_genders} at row {excel_row_number}')
            continue
        age_code = row['age code']
        age_code = str(age_code).upper().strip() if age_code else ''
        if not age_code:
            print(f'WARNING: age code missing at row {excel_row_number}')
            continue
        category = gender + " " + age_code
        level_scores = [0] * num_ea_pb_levels
        for level_idx in range(0, num_ea_pb_levels):
            heading = f'level {level_idx + 1}'
            perf = row[heading]
            perf = str(perf).lower().strip() if perf else ''
            if not perf:
                print(f'WARNING: performance missing for {heading} at row {excel_row_number}')
                break
            score, original_dp, original_special, invalid = make_numeric_score_from_performance_string(perf)
            level_scores[level_idx] = score

        score_set = EaPbAwardScoreSet(bucket, event, category, level_scores)
        if event not in ea_pb_award_score:
            ea_pb_award_score[event] = {}
        ea_pb_award_score[event][category] = score_set
        rows_completed += 1
    
    print(f'... processed {rows_completed} rows from EA PB Awards tables')


def main(club_id=238, output_file='records.htm', first_year=2005, last_year=2024, 
         do_po10=False, do_runbritain=True, input_files=[],
         cache_file='cache.pkl', rebuild_last_year=False, first_claim_only=False,
         types=['T', 'F', 'R', 'M'], do_wava=True, rebuild_wava=False,
         ea_pb_award_file=None):

    # Retrieve cache of performances obtained from web trawl previously
    try:
        with open(cache_file, 'rb') as fd:
            performance_cache = pickle.load(fd)
            print(f'Cached web results retrieved from {cache_file}')
    except IOError:
        print(f"Cache file {cache_file} can't be opened, starting new cache")
        performance_cache = {}

    if ea_pb_award_file:
        read_ea_pb_award_score_tables(ea_pb_award_file)

    for year in range(first_year, last_year + 1):
        # E.g. to rebuild in Jan 2024 want last results from 2023 so year before too
        rebuild_cache = rebuild_last_year and (year == last_year or year == last_year - 1)
        for gender in ['W', 'M']:
            if do_po10:
                for category in powerof10_categories:
                    process_one_po10_year_gender(club_id, year, gender, category,
                                                 performance_cache, rebuild_cache, first_claim_only,
                                                 types)
            if do_runbritain:
                for (event, _, _, runbritain, type, categories) in known_events: # debug [('Mar', True, 3, True, 'R')]:
                    if not runbritain: continue
                    if type not in types: continue
                    for (category, _, _) in runbritain_categories: # debug [('ALL', 0, 0), ('V50', 50, 54)]
                        if event_relevant_to_category(event, gender, category):
                            process_one_runbritain_year_gender(club_id, year, gender, category, event,
                                                            performance_cache, rebuild_cache,
                                                            first_claim_only, types, do_wava, rebuild_wava)

    # Input files last so manual 'invalidate' entries will remove known anomalies from Po10
    for input_file in input_files:
        process_one_club_record_input_file(input_file, types)

    # Save updated cache for next time
    try:
        with open(cache_file, 'wb') as fd:
            pickle.dump(performance_cache, fd)
        print(f'Cached web results written to {cache_file}')
    except IOError:
        print(f"Cache file {cache_file} can't be written, any new web results this time not cached")

    club_name = get_po10_club_name(club_id)

    output_records(output_file, first_year, last_year, club_id, do_po10, do_runbritain, input_files, club_name)


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
    parser.add_argument('--rebuild-wava', dest='rebuild_wava',  choices=yes_no_choices, default='n')
    parser.add_argument('--first-claim-only', dest='first_claim_only',  choices=yes_no_choices, default='n')
    parser.add_argument('--track', dest='track',  choices=yes_no_choices, default='y')
    parser.add_argument('--field', dest='field',  choices=yes_no_choices, default='y')
    parser.add_argument('--road', dest='road',  choices=yes_no_choices, default='y')
    parser.add_argument('--multievent', dest='multievent',  choices=yes_no_choices, default='y')
    parser.add_argument('--wava', dest='wava',  choices=yes_no_choices, default='y')
    parser.add_argument('--ea-pb-award-file', dest='ea_pb_award_file', default=None)

    args = parser.parse_args()

    do_po10           = args.do_po10.lower().startswith('y')
    do_runbritain     = args.do_runbritain.lower().startswith('y')
    rebuild_last_year = args.rebuild_last_year.lower().startswith('y')
    rebuild_wava      = args.rebuild_wava.lower().startswith('y')
    first_claim_only  = args.first_claim_only.lower().startswith('y')
    do_wava           = args.wava.lower().startswith('y')
    ea_pb_award_file  = args.ea_pb_award_file
    types = []
    if args.track.lower().startswith('y'):      types.append('T')
    if args.field.lower().startswith('y'):      types.append('F')
    if args.road.lower().startswith('y'):       types.append('R')
    if args.multievent.lower().startswith('y'): types.append('M')

    main(club_id=args.club_id, output_file=args.output_filename, first_year=args.first_year, 
         last_year=args.last_year, do_po10=do_po10, do_runbritain=do_runbritain, 
         input_files=args.excel_file, cache_file=args.cache_filename, rebuild_last_year=rebuild_last_year,
         first_claim_only=first_claim_only, types=types, do_wava=do_wava, rebuild_wava=rebuild_wava,
         ea_pb_award_file=ea_pb_award_file)
