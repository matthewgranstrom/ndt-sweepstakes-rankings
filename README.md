# ndt-sweepstakes-rankings

A python script to turn tournament results downloaded from tabroom.com into a report ranking schools by NDT sweepstakes points. 

## Usage:

Runs as a jupyter notebook with hard-coded spring/fall and year, modify `YEAR_TO_PROCESS` and `REPORT_TO_GENERATE` to change these parameters.

To load tournaments, create a tournament object with a name, a reference to the year, a tuple of round counts in each division, and a tuple of divisions.
When processing a tournament, this script looks in the `tournament_results/<year>` directory for a folder matching the name of the tournament.
For each division, there should be a `<name>-<division>-prelims.csv`, downloaded straight from tabroom.com's 'prelim records' page for the tournament, and any number of `<name>-<division>-elim-<x>.csv` files, each containing the results of one elimination round, again as downloaded from tabroom.com.

In the root directory, the script expects a `community-colleges-<year>.csv`, indicating which schools (if any) are community colleges, and a `ndt-districts-<year>`.csv, listing the NDT district to which each school belongs.

Also, this script expects two word documents, which will bookend the tables generated: 

1.`sweepstakes-table-template.docx` should contain any introduction. The first table style in this document will be used in each of the tables generated. In addition, any instance of `$YEAR` will be replaced with the chosen year, and `$SEASON_<FOO>` will be replaced with a formatted season string.
2.`sweepstakes-procedure.docx` contains any conclusion or appendices. 

## Future plans:

1. Jupyter notebook is a great tool, but expect (February 2024) a major refactor to allow the script to be run with standard Python3, taking command-line input for the year and season.
2. I'd like to externalize the list of tournaments. I'd like to just read from another csv file, but right now I'm expecting the tournament definitions to include a tuple, and csvs don't play nice with extra bonus commas. 
3. The fall report works as intended, but the spring report should include a list of new schools and a list of 'movers'. This requires last year's reports be generated and loaded.
4. In the first elim round, sweepstakes procedure stipulates that a maximum of half the field can earn sweepstakes points in any elim. None of the tournaments listed in 2023 cleared more than half of its entries, but a tournament doing so would result in extra points being awarded for teams that would not clear at an ADA tournament.