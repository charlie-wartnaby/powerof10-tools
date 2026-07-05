#!/bin/env python

"""
   Loads web cache file saved by these tools in Python .pkl format with pre-2026
   PowerOf10 and runbritain data, and resaves in long-term format. That data was
   disposable at the time because it could be re-acquired from the websites, but
   now they are gone it is important for historical club performances, so preserving
   it for the future matters.

   (.pkl not a good long-term format because it is not necessarily portable between
   Python versions.)

   First written by (c) Charlie Wartnaby 2026
   See https://github.com/charlie-wartnaby/powerof10-tools
"""

import argparse
import os
import pandas as pd
import pickle


class Performance():
    """Copy of class used for performance cache for pre-2026 processing, to allow reload from binary cache;
    all members now have defaults just to provide datatypes as per original usage"""

    def __init__(self, event='', score=0.0, category='', gender='', original_special='', decimal_places=0,
                 athlete_name='', athlete_url='', date='',
                 fixture_name='', fixture_url='', source='', wava=0.0, age=0, invalid=False, ea_pb_score=0.0):
        self.event = event
        self.score = score # could be time in sec, distance in m or multievent points, numeric
        self.category = category # e.g. U20 or ALL
        self.gender = gender # W or M
        self.original_special = original_special # for wind-assisted detail etc from club records, string version of original entry
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

        
def main(input_path, output_path):
    """Top-level processing entry point"""

    reqd_input_extension  = [".pkl"]
    reqd_output_extension = [".csv", ".zip"]
    _, input_extension = os.path.splitext(input_path)
    if input_extension.lower() not in reqd_input_extension:
        raise ValueError(f"Input file must be {reqd_input_extension} format")
    _, output_extension = os.path.splitext(output_path)
    if output_extension.lower() not in reqd_output_extension:
        raise ValueError(f"Output file must be {reqd_output_extension} format")

    # Retrieve cache of performances obtained from web trawl previously
    with open(input_path, 'rb') as fd:
        performance_cache = pickle.load(fd)
        print(f'Cached web results retrieved from {input_path}')

    # Assemble into pandas dataframe; probably woefully inefficient, but
    # only intending to do it once
    dataframe = None
    num_lists = len(performance_cache)
    progress_count = 0
    last_output_count = 0
    output_interval = 1000
    for url_key, perf_list in performance_cache.items():
        for perf in perf_list:
            perf_dataframe = pd.DataFrame.from_dict([perf.__dict__])
            perf_dataframe.insert(0, "url_key", url_key)
            if dataframe is None:
                dataframe = perf_dataframe
            else:
                dataframe = pd.concat([dataframe, perf_dataframe])
        progress_count += 1
        if progress_count - last_output_count >= output_interval:
            percentage = (progress_count / num_lists) * 100.0
            print(f"Progress {percentage:.1f}%", flush=True)
            last_output_count = progress_count

    print(f"Writing output file: {output_path}", flush=True)
    dataframe.to_csv(output_path, index=False)
    print("All done")


if __name__ == "__main__":
    """Script entry point"""

    parser = argparse.ArgumentParser(description='Convert legacy po10/runbritain cache data to long-term format')

    parser.add_argument("-i", "--input-file",  default="cnc_cache.pkl", 
                        help="Path of input cache file to read from")
    parser.add_argument("-o", "--output-file", default="cnc_pre_2026_po10_runbritain_perforances.zip", 
                        help="Path of output file to write to")

    args = parser.parse_args()
    main(args.input_file, args.output_file)
