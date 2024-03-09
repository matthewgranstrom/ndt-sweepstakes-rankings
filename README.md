# ndt-sweepstakes-rankings

A python script to turn tournament results downloaded from tabroom.com into a report ranking schools by NDT sweepstakes points. 

## Usage:

To run, `python NDT-sweepstakes-2023-draft-3.py --year <year> --season <season>`. The script will take around five minutes to run, the primary culprit is the microsoft word libraries.

The script supports the following optional arguments: `--debug` or `-d` for additional debug information, `--validate` or `-v` for validation information to help identify any schools that may have changed names, and `--no_report` or `-n` to disable both processing of new/moving schools and generation of the Word tables.

To load tournaments, modify `tournaments_<year>.csv`, which contains a tournament name and the number of rounds in varsity, junior varsity, novice, and round-round-robin competition. For tournaments without a particular division, enter `0`.
When processing a tournament, this script looks in the `tournament_results/<year>` directory for a folder matching the name of the tournament.
For each division, there should be a `<name>-<division>-prelims.csv`, downloaded straight from tabroom.com's 'prelim records' page for the tournament, and any number of `<name>-<division>-elim-<x>.csv` files, each containing the results of one elimination round, again as downloaded from tabroom.com.

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