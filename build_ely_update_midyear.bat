REM This one re-requests data for the current year (only!) so it includes the lastest
REM public results, but is much slower to run. Older year results are still taken
REM from the cache file for speed.
REM Files provided in descending date order so duplicates in older sheets don't overwrite entries.
python get_rankings.py --clubid 480 --cache ely_cache.pkl --output ely_records.htm --rebuild-final-year y --rebuild-wava y 