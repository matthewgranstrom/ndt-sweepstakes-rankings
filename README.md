# ndt-sweepstakes-rankings

A python script to turn tournament results downloaded from tabroom.com into a report ranking schools by NDT sweepstakes points. 

## Usage:

To run, `python NDT-sweepstakes-2023-draft-3.py --year <year> --season <season>`. The script will take around five minutes to run, the primary culprit is the microsoft word libraries.
The `year` argument is the year in which the season starts (i.e. the 2021-2022 spring report is generated with `--year 2021 --season spring`).

For ADA points, `python ADA-sweepstakes-2023-draft-1.py --year 2023`. The ADA script takes the optional `-d` flag for debug output, but no others; the only output it provides is a members-only `ADA_members_only_output2023.csv` and an unfiltered `ADA_sweepstakes_output_2023.csv`.

The script supports the following optional arguments: `--debug` or `-d` for additional debug information, `--validate` or `-v` for validation information to help identify any schools that may have changed names, and `--no_report` or `-n` to disable both processing of new/moving schools and generation of the Word tables.

To load tournaments, modify `tournaments_<year>.csv`, which contains the following columns:

1. Tournament name
2. Number of prelim rounds in Varsity
3. Number of prelim rounds in Junior Varsity
4. Number of prelim rounds in Novice
5. Number of Round-Robin prelims
6. Season
7. ADA sanction (1/0 for yes/no)

When processing a tournament, this script looks in the `tournament_results/<year>` directory for a folder matching the name of the tournament.
For each division, there should be a `<name>-<division>-prelims.csv`, downloaded straight from tabroom.com's 'prelim records' page for the tournament, and any number of `<name>-<division>-elim-<x>.csv` files, each containing the results of one elimination round, again as downloaded from tabroom.com.
For ADA points, the ADA script requires a `name-division-speakers.csv`, unfortunately tabroom.com does not allow you to download straight from that page. The recommended solution is to copy and paste the table tabroom.com generates.

In the root directory, the script expects the following information about schools:

1. `community-colleges.csv`, indicating which schools (if any) are community colleges, 
2. `ndt-districts`.csv, listing the NDT district to which each school belongs,
3. `school_alias_map.csv`, listing (in the first column) the display name for each school (e.g. "University of Minnesota" or "United States Naval Academy") and in the remaining columns any names listed for these schools on tabroom.com (e.g. "Minnesota" or "Navy").

Also, this script expects two word documents, which will bookend the tables generated: 

1.`sweepstakes-table-template.docx` should contain any introduction or front matter. The first table style in this document will be used in each of the tables generated. In addition, any instance of `$YEAR` will be replaced with the chosen year, and `$SEASON_<FOO>` will be replaced with a formatted season string.
2.`sweepstakes-procedure.docx` contains any conclusion or appendices. 

## Future plans:

1. The table of contents does not know about the movers or new-schools reports, and is as a result very wrong in spring reports.
2. In the first elim round, sweepstakes procedure stipulates that a maximum of half the field can earn sweepstakes points in any elim. None of the tournaments listed in 2023 cleared more than half of its entries, but a tournament doing so would result in extra points being awarded for teams that would not clear at an ADA tournament.
3. Ideally, you would run this script and it would go download the results for you. There were some high schoolers who DDOS'ed Tabroom a few years ago, I wonder if their API is any good.
4. Ideally, the script would report the list of tourneys/divisions that counted and did not count, and the reason, at least for record-keeping.