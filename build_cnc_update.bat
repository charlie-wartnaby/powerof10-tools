REM This one re-requests data for the current year so it includes the lastest
REM public results, but is much slower to run. Older year results are still taken
REM from the cache file for speed.
REM Files provided in descending date order so duplicates in older sheets don't overwrite entries.
python get_rankings.py --clubid 238 --cache cnc_cache.pkl --rebuild-last-year y --rebuild-wava y --ea-pb-award-file EA_PB_Awards_tables.xlsx 2022_CnC_records.xlsx 2021_CnC_records.xlsx 2009_CnC_records.xlsx CnC_known_historical.xlsx