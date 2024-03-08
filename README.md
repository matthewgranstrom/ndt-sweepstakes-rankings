# ndt-sweepstakes-rankings

A python script to turn tournament results downloaded from tabroom.com into a report ranking schools by NDT sweepstakes points. 

## Usage:

To run, `python NDT-sweepstakes-2023-draft-3.py --year <year> --season <season>`. The script will take around five minutes to run, the primary culprit is the microsoft word libraries.

To load tournaments, modify `tournaments_<year>.csv`, which contains a tournament name and the number of rounds in varsity, junior varsity, novice, and round-round-robin competition. For tournaments without a particular division, enter `0`.
When processing a tournament, this script looks in the `tournament_results/<year>` directory for a folder matching the name of the tournament.
For each division, there should be a `<name>-<division>-prelims.csv`, downloaded straight from tabroom.com's 'prelim records' page for the tournament, and any number of `<name>-<division>-elim-<x>.csv` files, each containing the results of one elimination round, again as downloaded from tabroom.com.

In the root directory, the script expects a `community-colleges-<year>.csv`, indicating which schools (if any) are community colleges, and a `ndt-districts-<year>`.csv, listing the NDT district to which each school belongs.

Also, this script expects two word documents, which will bookend the tables generated: 

1.`sweepstakes-table-template.docx` should contain any introduction or front matter. The first table style in this document will be used in each of the tables generated. In addition, any instance of `$YEAR` will be replaced with the chosen year, and `$SEASON_<FOO>` will be replaced with a formatted season string.
2.`sweepstakes-procedure.docx` contains any conclusion or appendices. 

Access debug mode with `-d`, output will describe each tournament and elim file processed. Disable both calculation of movers/new schools and docx report generation with `--no_report`. This option is most useful for generating the 'year zero' fall/spring reports that the year 1 reports you really want to generate will assume are present.

## Future plans:

1. The spring report should by rule only award points to NDT members. This is not currently implemented.
2. The table of contents does not know about the movers or new-schools reports, and is as a result very wrong in spring reports.
3. In the first elim round, sweepstakes procedure stipulates that a maximum of half the field can earn sweepstakes points in any elim. None of the tournaments listed in 2023 cleared more than half of its entries, but a tournament doing so would result in extra points being awarded for teams that would not clear at an ADA tournament.
4. Ideally, you would run this script and it would go download the results for you. There were some high schoolers who DDOS'ed Tabroom a few years ago, I wonder if their API is any good.
5. Right now, I process non-canonical names (CalBerkeley and UCBerkeley for Cal, for instance) right before I output the tables, this means I'm not properly storing this information. This check needs to be moved to before I save the reports, otherwise I'll think a school is new every time they change their name.