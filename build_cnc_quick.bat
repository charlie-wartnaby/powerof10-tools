REM This one updates the output report without re-requesting web data for this year,
REM so it doesn't find new performances since the last 'update' run.
REM Files provided in descending date order so duplicates in older sheets don't overwrite entries.
python get_rankings.py --clubid 238 --cache cnc_cache.pkl 2022_CnC_records.xlsx 2021_CnC_records.xlsx 2009_CnC_records.xlsx CnC_known_historical.xlsx