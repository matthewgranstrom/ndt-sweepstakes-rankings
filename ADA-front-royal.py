### import statements
import pandas as pd
pd.options.mode.chained_assignment = None  #supress warnings, i've tested the code
import numpy as np
from enum import Enum
import os
import re
import docx
from docxcompose.composer import Composer
import argparse

#take input for year and season
command_line_argument_parser=argparse.ArgumentParser()
command_line_argument_parser.add_argument("-y","--year",help="Year to tabulate, default 2024",type=int,default=2024)
command_line_argument_parser.add_argument("-d","--debug",help="Debug mode",action='store_true')
arguments=command_line_argument_parser.parse_args()

YEAR_TO_PROCESS=arguments.year

##todos: combine all divs to 'grand sweepstakes'
#        limit to only ada members


### global definitions
def ada_points_column_from_prelims(prelim_wins_column,prelim_count): ## taken from the ranking procedure.
    if prelim_count == 4:
        return prelim_wins_column+1
    elif prelim_count == 5:
        points_dict = {0:1, 1:3, 2:5, 3:6, 4:7, 5:10}
    elif prelim_count == 6:
        points_dict = {0:1, 1:3, 2:4, 3:5, 4:6, 5:7, 6:10}
    elif prelim_count == 7:
        points_dict = {0:1, 1:2, 2:3, 3:4, 4:6, 5:7, 6:8, 7:9}
    elif prelim_count == 8:
        points_dict = {0:1, 1:2, 2:3, 3:4, 4:5, 5:6, 6:6, 7:8, 8:10}
    elif prelim_count == 9:
        return prelim_wins_column+1
    else:
        raise ValueError("No ADA points table defined for tournaments with the following number of rounds: ",prelim_count)
    return prelim_wins_column.replace(points_dict)

def print_if_debug(string):
    if arguments.debug:
        print(string)
    return

#I might be able to significantly speed this code up by attempting to vectorize, but do I really care?
def ada_winner_points_from_elims(loser_ballots): #No bonus points for unanimity in the ADA
    return 3
def ada_loser_points_from_elims(loser_ballots): #No consolation points for ballots in the ADA
    return 0

class Division(Enum):
    VARSITY = 'v'
    JUNIOR_VARSITY = 'jv'
    NOVICE = 'n'
    ROUND_ROBIN = 'rr'
    TOTAL = 'total' ##hack. TODO: Remove.

    
# I could probably get away without referencing the year, but I want to be able to process last year's results to generate
# this year's spring reports with 'movers' and 'new schools' so I plan for the future
class tournament():
        def __init__(self,tournament_name,tournament_year,prelim_count_vector,division_vector):
            self.prelim_counts = prelim_count_vector
            self.divisions = division_vector
            self.name = tournament_name
            self.year = tournament_year

#taken from the ranking procedure. all of these are checked inclusively.
MAXIMUM_TOURNAMENTS_COUNTED_PER_SCHOOL = 8
MAXIMUM_RECORDS_COUNTED_PER_SCHOOL = 8
MAX_RECORDS_FOR_SCHOOL_AT_TOURNAMENT = 2
FIVE_SPEAKER_THRESHOLD = 21 #20 debaters -> only three
TEN_SPEAKER_THRESHOLD = 31 # 30 debaters -> only five
POINTS_FOR_CLEARING = 3
POINTS_FOR_MISSING_ON_POINTS = 1
ADANATS_BONUS_FACTOR = 1.5

##Prepare to replace school names with 'pretty' school names for display: 'Minnesota' -> 'University of Minnesota'

school_alias_dataframe=pd.read_csv('school-alias-map.csv')
school_alias_dict_dataframe=pd.DataFrame()
for alias_item in school_alias_dataframe.columns[1:]:
    temp_school_alias_dataframe=pd.DataFrame()
    temp_school_alias_dataframe[['Display Name','Alias']]=school_alias_dataframe[['Display-School',alias_item]]
    temp_school_alias_dataframe.dropna(inplace=True)
    if school_alias_dict_dataframe.empty:
        school_alias_dict_dataframe = temp_school_alias_dataframe
    else:
        school_alias_dict_dataframe = pd.concat([school_alias_dict_dataframe,temp_school_alias_dataframe],ignore_index=True)

school_alias_dict=school_alias_dict_dataframe.set_index('Alias')['Display Name'].T.to_dict()

def apply_dictionary_to_results_dataframe(results_dataframe,school_dictionary):
    results_dataframe_test = results_dataframe['School'].map(school_dictionary)
    unmapped_schools=results_dataframe[results_dataframe_test.isna()]
    unmapped_school_count=len(unmapped_schools.index)
    #if unmapped_school_count>0:
    #    print_if_debug('The following rows contain unmapped schools:')
    #    print_if_debug(unmapped_schools['School'].to_string())
    results_dataframe['School'] = results_dataframe['School'].map(school_dictionary).fillna(results_dataframe['School'])
    return results_dataframe




# I use '3-0' to describe any unanimous win (bye, forfeit, walkover). If the scoring conditions are changed to award different
# points for a 3-0 win than a 5-0 win (or whatever), this code will break.

# Replaces elim rows with explicit walkovers, assigns a 3-0 win to the advancing entry.
def process_elim_walkovers(elim_record):
    elim_walkovers = pd.DataFrame()
    elim_walkovers = elim_record[elim_record['Win'].str.contains('advances')]
    if not elim_walkovers.empty:
        elim_walkovers['walkover_winner'] = elim_walkovers['Win'].str[:-9]
        elim_walkovers['aff_walks_over'] = elim_walkovers['walkover_winner'] == elim_walkovers['Aff']
        elim_walkovers['walkover_ballot'] = elim_walkovers['aff_walks_over'].replace({True: '3-0\tAFF', False: '3-0\tNEG'})
        elim_walkovers.drop(columns=['walkover_winner','aff_walks_over','Win'],axis=1,inplace=True)
        elim_walkovers['Win'] = elim_walkovers['walkover_ballot']
        elim_walkovers.drop('walkover_ballot',axis=1,inplace=True)
        elim_record.drop(elim_record.index[elim_record['Win'].str.contains('advances')],inplace=True)
        elim_record = pd.concat([elim_record,elim_walkovers[['Aff','Neg','Win']]])
    return elim_record

# Detects elim rounds where the tournament organizer has denoted an elim walkover by simply pairing two teams from the same
# school against each other and has not published a winner. Manual intervention is required to correct these elim rounds,
# because it's necessary to look at the next elim round to determine who actually advanced.
def detect_implied_walkovers(elim_record):
    elim_record['implied_walkover'] = (elim_record['Win'].isna()) & (~(elim_record['Neg'].isna())) & (~(elim_record['Aff'].isna()))
    return elim_record['implied_walkover'].any()

# replaces NaN and blanks with acceptable input for bye and walkover processing. 
# implicitly assumes that any team getting a bye is listed on the affirmative.
def eliminate_blank_teams(elim_record):
    elim_record
    na_values = {'Neg': 'BLANK_ENTRY', 'Win': 'Aff BYE'}
    elim_record.fillna(value=na_values,inplace=True)
    return elim_record

# awards a 0-3 loss for any team forfeiting an elimination round.
def process_elim_forfeits(elim_record):
    elim_forfeits = pd.DataFrame()
    elim_forfeits = elim_record[elim_record['Win'].str.contains('FFT')]
    if not elim_forfeits.empty:
        elim_forfeits[['aff_result','neg_result']] = elim_forfeits['Win'].str.split('\t+',expand=True)
        elim_forfeits['aff_result'] = elim_forfeits['aff_result'].str[4:]
        elim_forfeits['Win'] = elim_forfeits['aff_result'].replace({'FFT': '3-0\tNEG','BYE':'3-0\tAff'})
        elim_record = elim_record.drop(elim_record.index[elim_record['Win'].str.contains('FFT')])
        elim_record = pd.concat([elim_record,elim_forfeits])
    return elim_record

    
 # awards a 3-0 to the team getting a bye
def process_elim_byes(elim_record):
    elim_byes = pd.DataFrame()
    elim_byes = elim_record[elim_record['Win'].str.contains('BYE')]
    if not elim_byes.empty:
        elim_byes[['Win','bye']] = elim_byes['Win'].str.split(' ',expand=True)
        elim_byes['Win'] = elim_byes['Win'].str.upper()
        elim_byes[['bye']] = elim_byes[['bye']].replace('BYE','3-0\t')
        elim_byes['Win'] = elim_byes['bye'] + elim_byes['Win']
        elim_byes = elim_byes.drop('bye',axis=1)
        elim_record = elim_record.drop(elim_record.index[elim_record['Win'].str.contains('BYE')])
        elim_record = pd.concat([elim_record,elim_byes])
    return elim_record

# ensure hybrid entires are not awarded points. There are currently no 'ADA-recognized' hybrid teams, if any exist this function 
# will need to be modified to account for this. That change would also force me to assign a school to the hybrid team.
def ada_drop_hybrid_entries(tournament_points):
    tournament_points['is_team_hybrid'] = tournament_points['Code'].str.contains('/')
    tournament_points.drop(tournament_points[tournament_points.is_team_hybrid].index, inplace=True)
    tournament_points.drop(['is_team_hybrid'],axis=1,inplace=True)
    return tournament_points
    
    
def get_data_folder(tournament_name,year):
    return 'tournament_results/'+str(year)+'/'+tournament_name
    
def load_prelims_from_tournament_folder(tournament_name,year,division):
    data_folder = get_data_folder(tournament_name,year)
    prelimFilePath=data_folder+'/'+tournament_name+'-'+division.value+'-prelims.csv'
    tournament_prelims = pd.read_csv(prelimFilePath)
    return tournament_prelims

def load_speaker_points_from_tournament_folder(tournament_name,year,division):
    data_folder = get_data_folder(tournament_name,year)
    speakerFilePath=data_folder+'/'+tournament_name+'-'+division.value+'-speakers.csv'
    tournament_speakers = pd.read_csv(speakerFilePath)
    return tournament_speakers[['Place','Entry','School']]

def ada_apply_speaker_points(speaker_dataframe):
    debater_count = len(speaker_dataframe)
    speaker_point_dict = {1:3, 2:2, 3:1, 4:1, 5:1, 6:1, 7:1, 8:1, 9:1, 10:1}
    speaker_eligible_count = 3+2*(debater_count>=FIVE_SPEAKER_THRESHOLD)+5*(debater_count>=TEN_SPEAKER_THRESHOLD)
    speaker_dataframe = speaker_dataframe.head(speaker_eligible_count)#assumes the speaker point csv is sorted, which tabroom by default does.
    speaker_dataframe['speaker_points'] = speaker_dataframe['Place'].astype(int).map(speaker_point_dict).fillna(0)
    speaker_dataframe = speaker_dataframe[['Entry','speaker_points']]
    speaker_dataframe_merged = speaker_dataframe.groupby(speaker_dataframe['Entry'],as_index=False).aggregate(sum)
    return speaker_dataframe_merged[['Entry','speaker_points']]
    
    
def load_elims_from_tournament_folder(tournament_name,year,division,prelim_count): #returns a vector of dataframes
    data_folder = get_data_folder(tournament_name,year)
    dir_list = os.listdir(data_folder)
    search_string = '-'+division.value+'-elim'
    elims_to_process = filter(lambda x: re.search(search_string, x), dir_list)
    elim_index=0
    elim_return_vector = {}
    for elim_filename in list(elims_to_process):
        elim_index+=1
        print_if_debug('\t\t'+elim_filename)
        elim_record=pd.read_csv(data_folder+'/'+elim_filename)[['Aff','Neg','Win']]
        if detect_implied_walkovers(elim_record):
            raise ValueError("Human intervention needed: Implied walkover in ",elim_filename)
        elim_record = eliminate_blank_teams(elim_record)
        elim_record = process_elim_walkovers(elim_record)
        elim_record = process_elim_forfeits(elim_record)
        elim_record = process_elim_byes(elim_record)
        elim_record[['ballots','Win']] = elim_record['Win'].str.split('\t+',expand=True) # this is inelegant, but works
        elim_record[['winner_ballots','loser_ballots']] = elim_record['ballots'].str.split('-',expand=True) #breaks if there's not a dash in there, should be caught above
        elim_record[['loser_ballots']] = elim_record[['loser_ballots']].astype("int")
        elim_record['aff_win'] = elim_record['Win'].apply(lambda y: 1 if y=='AFF' else 0)## TODO: replace sad slow apply with vectorized happy fast 'replace'
        elim_record['neg_win'] = 1-elim_record['aff_win']
        elim_return_vector[elim_index]=elim_record
    return elim_return_vector
        
        
def ada_apply_elim_points(elim_record,elim_index):
    elim_record[['winner_points']] = elim_record[['loser_ballots']].apply(ada_winner_points_from_elims)
    elim_record[['loser_points']] = elim_record[['loser_ballots']].apply(ada_loser_points_from_elims)
    if elim_index == 1:
        elim_record[['winner_points']]+=(POINTS_FOR_CLEARING-POINTS_FOR_MISSING_ON_POINTS) #i'll give everyone who was eligible the 'missing' value later
        elim_record[['loser_points']]+=(POINTS_FOR_CLEARING-POINTS_FOR_MISSING_ON_POINTS)
    return elim_record
    
def merge_elim_affs_negs(elim_record,elim_index,points_column_header):
    elim_record['aff_points'] = elim_record['winner_points']*elim_record['aff_win']+elim_record['loser_points']*elim_record['neg_win']
    elim_record['neg_points'] = elim_record['winner_points']*elim_record['neg_win']+elim_record['loser_points']*elim_record['aff_win']
    temp_aff = pd.DataFrame()
    temp_neg = pd.DataFrame()
    temp_aff[['Code',points_column_header]] = elim_record[['Aff','aff_points']]
    temp_neg[['Code',points_column_header]] = elim_record[['Neg','neg_points']]
    elim_points = pd.concat([temp_aff,temp_neg])
    return elim_points

def ada_process_points_division(tournament_name,year,prelim_count,division):# returns a dataframe containing the school and the points earned by each of the school's top two entries
    school_division_points = pd.DataFrame()
    tournament_prelims = load_prelims_from_tournament_folder(tournament_name,year,division)
    tournament_prelims['prelim_points'] = ada_points_column_from_prelims(tournament_prelims['Wins'],prelim_count)
    tournament_points=tournament_prelims[['Code','Name','School','prelim_points','Wins']]
    tournament_speakers = load_speaker_points_from_tournament_folder(tournament_name,year,division)
    points_from_speakers = ada_apply_speaker_points(tournament_speakers)
    tournament_points = tournament_points.merge(points_from_speakers,how='left',left_on='Code',right_on='Entry').drop(columns=['Entry']).fillna(0)
    tournament_points['speaker_points'] = tournament_points['speaker_points'].astype(int)
    
    tournament_elim_results_vector = load_elims_from_tournament_folder(tournament_name,year,division,prelim_count)
    ran_elims=False #default, we'll set it to true if any of the elims had index 1
    for (elim_index,elim_dataframe) in zip(tournament_elim_results_vector.keys(),tournament_elim_results_vector.values()):
        points_column_header='elim_'+str(elim_index)+"_points"
        elim_dataframe = ada_apply_elim_points(elim_dataframe,elim_index)
        elim_dataframe = merge_elim_affs_negs(elim_dataframe,elim_index,points_column_header)
        tournament_points = tournament_points.merge(elim_dataframe,'left','Code')
        tournament_points[[points_column_header]] = tournament_points[[points_column_header]].fillna(0).astype(int)
        if elim_index == 1:
            ran_elims = True
            tournament_points['elim_participant'] = tournament_points['elim_1_points']>0
    tournament_points = ada_drop_hybrid_entries(tournament_points)
    if ran_elims:
        wins_to_clear = tournament_points[tournament_points['elim_participant']]['Wins'].min()
        tournament_points['should_have_cleared'] = tournament_points['Wins']>=wins_to_clear
        tournament_points['should_have_cleared'] = tournament_points['should_have_cleared'].map({True:POINTS_FOR_MISSING_ON_POINTS,False:0})
        tournament_points.drop(columns=['Wins','elim_participant'],inplace=True)
    points_column_name = tournament_name +'_' + division.value + '_points'
    tournament_points[points_column_name] = tournament_points.drop(['Code','Name','School'],axis=1).sum(axis=1)
    entry_division_points = tournament_points[['Name','School',points_column_name]].sort_values(points_column_name,ascending=False,ignore_index=True)
    #school_division_points = tournament_points[['School','total_points']].sort_values('total_points',ascending=False,ignore_index=True) #for front royal, we don't ever merge by school
    #school_division_points = school_division_points.groupby('School',as_index=False)
    #school_division_points = school_division_points.head(MAX_RECORDS_FOR_SCHOOL_AT_TOURNAMENT) 
    #apply_dictionary_to_results_dataframe(school_division_points,school_alias_dict)
    return entry_division_points
	
### Functions to split tournaments into divisions, and integrate tournaments into one Big Table
def process_points_tournament(tournament):
    tournament_name=tournament.name
    prelim_count_vector=tournament.prelim_counts
    division_vector=tournament.divisions
    year=tournament.year
    print_if_debug('\ttournament: '+tournament_name)
    entry_division_points=pd.DataFrame()
    for (division,prelim_count) in zip(division_vector, prelim_count_vector):
        if prelim_count==0:
            continue
        division_points = pd.DataFrame()
        division_points = ada_process_points_division(tournament_name,year,prelim_count,division)
        if entry_division_points.empty:
            entry_division_points = division_points #the merge will error out if there aren't any rounds in the division (and consequently the output dataframe is empty)
        elif division_points.empty:
            print_if_debug('no points added for '+tournament_name+' '+division.name)
        else:
            entry_division_points = entry_division_points.merge(division_points,how='outer',on='Name')
            print_if_debug(entry_division_points.to_string())
    entry_division_points.fillna(0,inplace=True)
    return entry_division_points

def tournament_merge(cumulative_list,new_tournament):
    if cumulative_list.empty:
        return new_tournament
    elif new_tournament.empty:
        return cumulative_list# the merge will error out if this is the first tourney processed because there's no 'school' column
    new_cumulative_list = cumulative_list.merge(new_tournament,how='outer',on='Name')
    possible_error_schools = cumulative_list.merge(new_tournament,how='inner',on='Name')[['Name','School_x','School_y']]
    actual_error_schools = possible_error_schools[possible_error_schools['School_x']!=possible_error_schools['School_y']]
    if not actual_error_schools.empty:
        print(actual_error_schools.to_string())
        raise ValueError("The teams above are listed with multiple different schools!") 
    new_cumulative_list['School'] = new_cumulative_list['School_x'].fillna(new_cumulative_list['School_y'])
    new_cumulative_list.drop(columns=['School_x','School_y'],inplace=True)
    new_cumulative_list.fillna(0,inplace=True)
    return new_cumulative_list
	
### define tournaments and execute
tournament_list = pd.read_csv('tournaments-'+str(YEAR_TO_PROCESS)+'.csv')
tournament_list = tournament_list[tournament_list['ada_sanctioned']==1]

cumulative_points_vector = {}
for division in [Division.VARSITY,Division.JUNIOR_VARSITY,Division.NOVICE]: # no sweepstakes points from RR rounds
    cumulative_points = pd.DataFrame()
    division_string = division.name.lower() + '_rounds'
    print_if_debug('Processing '+division.name+'...')   
    report_name = 'ADA_front_royal_cup_points_' + division.name + '_' + str(YEAR_TO_PROCESS)+'.csv'
    for tournament_index,tournament_data in tournament_list.iterrows():
        division_rounds = [tournament_data[division_string]]
        divisions = [division]
        tournament_to_process=tournament(tournament_data['tournament_name'],YEAR_TO_PROCESS,division_rounds,divisions)
        cumulative_points = tournament_merge(cumulative_points,process_points_tournament(tournament_to_process))
    cumulative_points['total_points'] = cumulative_points.drop(columns=['Name','School']).sum(axis=1).astype(int)
    cumulative_points.sort_values('total_points',ascending=False,ignore_index=True,inplace=True)
    cumulative_points.to_csv(index=False,path_or_buf=report_name)
    cumulative_points_vector[division.name] = cumulative_points
    
#check three-tourney status


three_division_merged = pd.DataFrame()
for division in cumulative_points_vector:
    division_result = cumulative_points_vector[division]
    three_division_merged = tournament_merge(three_division_merged,division_result)
three_division_merged = three_division_merged[three_division_merged.columns.drop(list(three_division_merged.filter(regex='total_points')))]
three_division_merged = three_division_merged.drop(columns='School').set_index('Name')
three_division_merged['divisions_entered'] = (three_division_merged > 0).sum(axis=1)
three_division_merged['FRC_eligible'] = three_division_merged['divisions_entered']>=3
frc_eligible = three_division_merged['FRC_eligible']

for division in cumulative_points_vector:
    division_result = cumulative_points_vector[division]
    division_result = division_result.merge(frc_eligible,how='left',left_on='Name', right_index = True)
    eligible_teams_list = division_result[division_result['FRC_eligible']]
    print(eligible_teams_list[['Name','School']].to_string())