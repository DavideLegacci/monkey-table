import pandas as pd
import os
import time
from datetime import datetime
import numpy as np
import xarray as xr
from tqdm import tqdm

# print(data)

####################################################################################################################################################################################
# START PARAMETERS TO EDIT
####################################################################################################################################################################################

empty_result = ''
positive_result = '1'
negative_result = ''
num_max_days = 23
sub_period_duration = 7
keep_kein_material = 'y'
reference_parameter = 'Albumin'
#reference_parameter = 'Proenkephalin'

#mode = 'test'
mode = 'full'

# Names
lab_results_raw_directory = 'a_lab_results_raw'     # data from software
directory_merged_results_per_patient = 'b_lab_results_per_patient'
directory_final_sheet = 'c_final_sheet'
patients_map_path = 'patients_map.xlsx'

####################################################################################################################################################################################
# END PARAMETERS TO EDIT
####################################################################################################################################################################################



####################################################################################################################################################################################
# START METADATA
####################################################################################################################################################################################



# if mode == 'full':
#   keep_kein_material = ''
#   while keep_kein_material not in ['y', 'n']:
#       keep_kein_material = input("\nShould I keep 'Kein Material' results? Type y or n: ")


current_date = datetime.utcfromtimestamp( int(time.time()) ).strftime('%Y-%m-%d-%H_%M_%S')


def find_nearest(items, pivot):
    return min(items, key = lambda x:abs(x - pivot ))

####################################################################################################################################################################################
# END METADATA
####################################################################################################################################################################################


####################################################################################################################################################################################
# START PATIENTS MAP
####################################################################################################################################################################################
patients_map = pd.read_excel(patients_map_path, keep_default_na = False)


def patNum2ID(num):
    return patients_map[ patients_map.LFDNR == num ].PATIFALLNR.iloc[0]

def patID2Num(ID):
    return patients_map[ patients_map.PATIFALLNR == ID ].LFDNR.iloc[0]


####################################################################################################################################################################################
# END PATIENTS MAP
####################################################################################################################################################################################



####################################################################################################################################################################################
# START DATA MERGING ROUTINE
# Goal: from multiple csv, each with data of multiple patients, get multiple excel files, each with all the data of a single patient
####################################################################################################################################################################################

perform_merging_routine = ''
while perform_merging_routine not in ['y', 'n']:
    perform_merging_routine = input("\nPerform merging routine? Type y or n: ")

if perform_merging_routine == 'y':

    lab_results_directory = f'{directory_merged_results_per_patient}/{current_date}'   # one file per patient

    def save_excel_patient_sheet(df, dirname, filename):

        filepath = f'{dirname}/{filename}'
        os.makedirs(dirname, exist_ok=True)

        with pd.ExcelWriter(filepath) as writer:
            df.to_excel(writer, index = False)

    raw_data = []
    for raw_result in os.listdir(lab_results_raw_directory):
        if raw_result.endswith(".csv") and not raw_result.startswith("~"):
            print(f'\n Extracting data from {raw_result}...')
            # encoding https://stackoverflow.com/questions/42339876/error-unicodedecodeerror-utf-8-codec-cant-decode-byte-0xff-in-position-0-in
            # separator https://stackoverflow.com/questions/18039057/python-pandas-error-tokenizing-data
            raw_data.append( pd.read_csv(f'{lab_results_raw_directory}/{raw_result}', encoding='cp1252', sep = ';', keep_default_na = False) )  

    # Merge into single
    raw_df = pd.concat( raw_data )
    print('\n All raw results merged into single result!')


    patient_IDs_in_current_labresults = set(raw_df.PATIFALLNR)
    patient_IDs_in_patients_map = set(patients_map.PATIFALLNR)
    #print(patient_IDs_in_current_labresults)
    #print(patient_IDs_in_patients_map)
    for p in patient_IDs_in_current_labresults:
        if p not in patient_IDs_in_patients_map:
            print('\nCurrent patient IDs:\n')
            [print(p) for p in patient_IDs_in_current_labresults ]
            raise Exception(f'PATIFALLNR {p} is NOT matched in patients map.')
        if patID2Num(p) == '':
            raise Exception(f'PATIFALLNR {p} is in patients map, but it is not associated to a number')

    print('\n Generating lab results file for each patient...\n')
    for patient in tqdm(set(raw_df.PATIFALLNR)):

        patient_number = patID2Num(patient)
        
        raw_df_patient = raw_df[raw_df.PATIFALLNR == patient]

        filename = f'{patient_number}-{patient}.xlsx'
        save_excel_patient_sheet(raw_df_patient, lab_results_directory, filename)

    print('\n-----------------------------------------------------------------------------------------------------------------------------------------')
    print('ALL GOOD :)\n Starting data manipulation.')
    print('-----------------------------------------------------------------------------------------------------------------------------------------\n')

if perform_merging_routine == 'n':

    name_of_directory_with_most_recent_results = os.listdir(directory_merged_results_per_patient)[-1]

    lab_results_directory = f'{directory_merged_results_per_patient}/{name_of_directory_with_most_recent_results}'   # one file per patient



####################################################################################################################################################################################
# END DATA MERGING ROUTINE
####################################################################################################################################################################################




####################################################################################################################################################################################
# START DATA MANIPULATION ROUTINE FOR EACH PATIENT
####################################################################################################################################################################################



def period_maker():
    '''Returns [7, 7, 7, 2] if num_max_days = 23 and sub_period_duration = 7'''
    if num_max_days < sub_period_duration:
        raise Exception('First input must be >= second')
    num_full_sub_periods = int(num_max_days/sub_period_duration)
    days_in_final_subperiod = num_max_days % sub_period_duration
    period_list = [sub_period_duration for _ in range(num_full_sub_periods)]
    if days_in_final_subperiod != 0: period_list = period_list +[days_in_final_subperiod]
    return period_list

def data_splitter(data):
    '''Given data of lenght num_max_days splits it according to period_maker

    e.g. num_max_days = 11
    sub_period_duration = 3
    period_list = period_maker() returns [3, 3, 3, 2]
    data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]

    data_splitter(data) returns [[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11]]

    '''
    period_list = period_maker()
    helper = [sum(period_list[:i]) for i in range(len(period_list)+1)]
    return [  data[helper[i]:helper[i+1]] for i in range(len(helper)-1) ]

all_days = [_ for _ in range(num_max_days)]
split_days = data_splitter(all_days)
#print(split_days)


def dictionary_values_splitter(dictionary_to_split):
    return { k:data_splitter(v) for k,v in dictionary_to_split.items() }

def dict_of_lists_to_list_of_dicts(dict_of_lists):
    return [ dict(zip(dict_of_lists, i)) for i in zip(*dict_of_lists.values()) ]

# <---------------------------------------------------- the parameters marked red in the numbers sheet must be added?
# This contains all 29 parameters in the order or Rebecca's work sheet; those marked red have -temp names, meaning I could not find them in lab sheet.
# Must find a lab sheet that contains them, and see how they are called

# TEST


# FIRST 29
#first_29_parameters  = ['tacro-temp', 'ciclo-temp', 'Natrium(ISE)', 'Kalium (ISE)', 'Calcium', 'Kreatinin', 'Proenkephalin', 'GFR, CKD-EPI', 'Harnstoff', 'Glucose', 'LDH', 'GOT/AST', 'GPT/ALT', 'AP', 'GGT', 'bili-temp', 'Phosphat', 'Ges.Eiweiss', 'Albumin quant.', 'CRP', 'Leukozyten', 'Hb', 'Thrombozyten', 'ntpro-temp', 'tnt-temp', 'INR - ber.', 'Quick', 'aPTT', 'ipth-temp']

# NEW 17
#new_17_parameters = ['pH/Tstr.', 'Glucose/Tstr.', 'Bili/Tstr.', 'Ketone /Tstr.', 'Erys /Tstr.', 'Eiweiß/Tstr.', 'Urobil /Tstr.', 'Nitrit /Tstr.', 'Leuko /Tstr.', 'U-Albumin', 'Protein/Urin', 'Eiweiss-temp', 'Erys/µl', 'Leuko/µl', 'platten-temp', 'Bakt./Sedu.', 'HyalZy./Sedu.']

testing_parameters = ['Albumin', 'C-reaktives Protein (CRP)', 'GFR nach CKD-EPI']
all_parameters_final = ['Tacrolimus ( MS )', 'Ciclosporin MC ( MS )', 'Natrium', 'Kalium', 'Calcium, korrigiert', 'Kreatinin', 'Proenkephalin', 'GFR nach CKD-EPI', 'Harnstoff', 'Glucose', 'Laktatdehydrogenase (LDH)', 'Glutamat-Oxalacetat-Transaminase/Aspartat-Aminotransferase (GOT/AST)', 'Glutamat-Pyruvat-Transaminase/Alanin-Aminotransferase (GPT/ALT)', 'Alkalische Phosphatase (AP)', 'gamma-Glutamyltransferase (GGT)', 'Bilirubin, gesamt', 'Phosphat', 'Gesamteiwei¤', 'Albumin', 'Procalcitonin (PCT), sensitiv', 'C-reaktives Protein (CRP)', 'Leukozyten', 'H_moglobin (Hb)', 'H_matokrit', 'Thrombozyten', 'n-terminal pro-brain natriuretic peptide (NT-ProBNP)', 'Troponin T (TNT), high sensitive im Plasma', 'International Normalized Ratio (INR)-berechnet', 'Quick', 'partielle Thromboplastinzeit, aktiviert (aPTT)', 'Parathormon (PTH), intakt', 'Triglyceride', 'Albumin im Urin / Kreatinin im Urin', 'pH-Wert im Urin (Teststreifen)', 'Nitrit im Urin (Teststreifen)', 'Eiwei¤ im Urin (Teststreifen)', 'Glucose im Urin (Teststreifen)', 'Ketonk_rper im Urin (Teststreifen)', 'Urobilinogen im Urin (Teststreifen)', 'Bilirubin im Urin (Teststreifen)', 'Bakterien im Urin/µl', 'Erythrozyten im Urin, absolut', 'Leukozyten im Urin, absolut', 'Granulierte Zylinder/µl', 'Plattenepithelien im Urin, absolut', 'Hyaline Zylinder/µl', 'Rundepithelien/µl', 'Eiweiß im Urin / d', 'pH-Wert, arteriell', 'Kohlendioxidpartialdruck (pCO2), arteriell', 'Sauerstoffpartialdruck (pO2), arteriell', 'Base Excess, arteriell', 'Standard-Bicarbonat, arteriell']

if mode == 'full':
    all_needed_parameters = all_parameters_final # first_29_parameters + new_17_parameters

elif mode == 'test':
    print('\n\n -------TEST MODE ------- \n\n')
    all_needed_parameters = testing_parameters


num_patients = 0
patient_identifier_PATIFALLNR = []

# Multiple sheets

number_of_sheets = int(num_max_days/sub_period_duration) + 1

big_data_multiple_sheets = [  [ ] for _ in range(number_of_sheets)  ]



# START PATIENT

# Read each excel file in lab_results_directory into dataframe and put it into data_list
print('\n Starting patients loop...')
for patient in tqdm(os.listdir(lab_results_directory)):
    # make sure to select only excel files; sometimes hidden files like ~$patient.xlsx are created, which must be excluded:
    if patient.endswith(".xlsx") and not patient.startswith("~"):

        num_patients +=1
        #patient_identifier_PATIFALLNR.append(patient)
        #print(f'processing patient {patient}...')

        # START MERGING DATAFRAME
        # Read data into two df
        data = pd.read_excel(f'{lab_results_directory}/{patient}') #, skiprows = 2, usecols = 'A, B, C, D')

        # Get rid of spaces in columns
        #data.columns = data.columns.str.replace(' ', '')

        # Get rid of » symbol in indiced
        # data.BESCHREIBUNG = data.BESCHREIBUNG.str.replace(' »', '')

        # IF INSTEAD PARAMETERS ARE VALUES OF COLUMN # <----------- MAIN DROP
        # https://stackoverflow.com/questions/18172851/deleting-dataframe-row-in-pandas-based-on-column-value
        for param in data.BESCHREIBUNG:
            if param not in all_needed_parameters:
                #print(f'{param} being dropped \n')
                data.drop(data.index[ data.BESCHREIBUNG == param ], inplace = True) 

                # Should be allright without this
                # try:
                #   data.drop(data.index[ data.BESCHREIBUNG == param ], inplace = True)
                # except:
                #   pass

        # Since some rows were dropped, not index looks like [0, 1, 5, 9, 15, ...]
        # Which is a mess because data.colum[i] refers to that index. So need to reset.
        data.reset_index(inplace = True, drop = True)

        # Now not needed parameters are dropped from data.BESCHREIBUNG. It may still happen that a needed parameter is not present in result. Fixed later. 

        # drop not needed columns
        #not_needed_columns = ['AUFTRAGNR', 'GEBDAT', 'SEX', 'EINSCODE', 'LABEINDAT']
        needed_columns = ['PATIFALLNR', 'BESCHREIBUNG', 'ERGEBNIST', 'VALIDIERTDAT']
        for col in data.columns:
            if col not in needed_columns:
                data.drop(col, axis = 1, inplace = True)


        # Fix dates format
        # for i in range(len(data.VALIDIERTDAT)):

        #   # Some dates were recognized by pandas already as datetimes; other are still strings because 1. they contain spaces and 2. they contain . rather than /
        #   if type( data.at[i, 'VALIDIERTDAT'] ) is str:
        #       data.at[i, 'VALIDIERTDAT'] = data.at[i, 'VALIDIERTDAT'].replace(' ', '')
        #       data.at[i, 'VALIDIERTDAT'] = data.at[i, 'VALIDIERTDAT'].replace('.', '/')   

        # Convert to datetime
        data['VALIDIERTDAT'] = pd.to_datetime(data['VALIDIERTDAT'], dayfirst = True)

        # Add column only with info about day
        data = data.assign(DAY=data['VALIDIERTDAT'].dt.strftime('%Y-%m-%d'))
        #data['DAY'] = pd.to_datetime(data['DAY'], dayfirst = True)

        #START GETTING RID OF kein material ROWS
        if keep_kein_material == 'n':
            strings_to_kill = ['Kein Material', 'K.Mat.']
            for s in strings_to_kill:
                index_to_kill = data.index[ data.ERGEBNIST == s ]
                [ print(f"---------------Dropping {s} for {data.at[ i , 'Parameter']}") for i in list(index_to_kill) ]
                data.drop(index_to_kill, inplace = True)

            data.reset_index(inplace = True, drop = True)

        # END GETTING RID OF kein material ROWS

        # NON STANDARD RESULTS
        # + and positive to 1
        # - and negative to empty
        # anything else remains text

        # Get rid of !L and !H flags
        for i in range(len(data.ERGEBNIST)):
            if type( data.at[i, 'ERGEBNIST'] ) is str:
                #data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace(' !L', '')
                #data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace(' !H', '')
                data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace('negativ', negative_result)
                data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace('-', negative_result)
                data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace('positiv', positive_result)
                data.at[i, 'ERGEBNIST'] = data.at[i, 'ERGEBNIST'].replace('+', positive_result)
        # END NON STANDARD RESULTS

        data = data.astype({"ERGEBNIST": str})

        # START GETTING RID OF DUPLICATE EXAM 
        # POSSIBILITIES:
        
        # 1. Restructure all so that parameters are column, not index, and use
        # https://stackoverflow.com/questions/50885093/how-do-i-remove-rows-with-duplicate-values-of-columns-in-pandas-data-frame

        # 2. Keep parameters as index, work on sub-frames for each parameter, drop duplicate date column
        # Then reconstruct big dataframe by composition

        ## IMPROVED DUPLICATED ALGORITH: KEEP WITH THIS PRIORITY
        # - the one done closest in time to penkid
        # - the first of the day

        # First make sure reference parameter does not appear more than once per day
        df_specific_for_reference_parameter = data.loc[ data.BESCHREIBUNG == reference_parameter ]
        boolean_duplicated_series_reference = df_specific_for_reference_parameter.duplicated( ['BESCHREIBUNG', 'DAY'], keep = False )
        if( boolean_duplicated_series_reference.any() ):
            df_duplicated_reference_parameter = df_specific_for_reference_parameter[boolean_duplicated_series_reference]
            
            for day in set(df_duplicated_reference_parameter.DAY):
                df_duplicated_reference_parameter_day = df_duplicated_reference_parameter[ df_duplicated_reference_parameter.DAY == day ]
                index_to_keep_reference = df_duplicated_reference_parameter_day['VALIDIERTDAT'].idxmin()
                #print()
                #print(df_duplicated_reference_parameter_day)
                for i in df_duplicated_reference_parameter_day.index:
                    if str(i) != str(index_to_keep_reference):
                        data.drop(i, inplace = True)
            data.reset_index(inplace = True, drop = True)
            #input(f'\n ---- ALERT----- \n\n Reference parameter {reference_parameter} appears more than once per day. Keeping first of each day. Enter to continue' )        
        ##############################################

        # OPTION 1
        #data.drop_duplicates( ['Parameter', 'VALIDIERTDAT'], keep = 'first', inplace = True, ignore_index = True  )
        for p in set(data.BESCHREIBUNG):
            #print(f'PERFORMING PARAMETER {p} WHILE REFERENCE IS {reference_parameter}\n')
            df_specific_for_p = data.loc[ data.BESCHREIBUNG == p ]

            # https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.duplicated.html
            boolean_duplicated_series = df_specific_for_p.duplicated( ['BESCHREIBUNG', 'DAY'], keep = False )

            if( boolean_duplicated_series.any() ):
                df_duplicated_p = df_specific_for_p[boolean_duplicated_series]
                for day in set(df_duplicated_p.DAY):
                    df_duplicated_p_day = df_duplicated_p[ df_duplicated_p.DAY == day ]
                    #print()
                    #print(df_duplicated_p_day)
                   

                    # Not reference parameter
                    current_day = list(df_duplicated_p_day.DAY)[0]
                    #####
                    reference_df = data[ data.BESCHREIBUNG == reference_parameter ]
                    reference_df = reference_df[ reference_df.DAY == current_day ]
                    #reference_time = list(reference_df.VALIDIERTDAT)[0]
                    # print()
                    # print('debug------------------')
                    # print(reference_df)
                    # print()
                    # print('debug------------------')
                    # print()
                    #####
                    try:
                        # reference_df = data[ data.BESCHREIBUNG == reference_parameter ]
                        # reference_df = reference_df[ reference_df.DAY == current_day ]
                        reference_time = list(reference_df.VALIDIERTDAT)[0]
                        closest_time = find_nearest( df_duplicated_p_day.VALIDIERTDAT, reference_time )
                        index_to_keep = df_duplicated_p_day.index[df_duplicated_p_day['VALIDIERTDAT'] == closest_time].tolist()[0]
                        # print('Comparison done')
                        # print(f'reference_time: {reference_time}')
                    except:
                        index_to_keep = df_duplicated_p_day['VALIDIERTDAT'].idxmin()
                    #     print('Just kept first')
                    # print(f'Index to keep: {index_to_keep}\n')

                    for i in df_duplicated_p_day.index:
                        if str(i) != str(index_to_keep):
                            data.drop(i, inplace = True)

        data.reset_index(inplace = True, drop = True)


        # END GETTING RID OF DUPLICATE EXAM # ----------

        # END MANIPULATING DATAFRAME


        # CAREFUL ACTUALLY DAY0 IS FROM EXTERNAL SOURCE, IT MAY BE THAT NO EXAM IS TAKEN ON DAY 0 <------------------------------------------------------------------------------------------------------------------- temporary
        # Get patient day0 = when she enters hospital

        day0 = min(data.DAY)


        # Get range of time in which exams are taken
        day_first_exam, day_last_exam = min(data.VALIDIERTDAT), max(data.VALIDIERTDAT)
        exam_period = day_last_exam - day_first_exam
        # print(exam_period)

        # Check time period
        if exam_period > pd.Timedelta(num_max_days, unit = 'd'):
            print(f'\n\n-----------SOMETHING WRONG---------\n\n Exam period lasts {exam_period}, longer than {num_max_days} days\n----------\n')
            raise Exception

        # Set maximal period of staying in the hospital; equal for everybody
        period = pd.date_range(start=day0, periods=num_max_days).strftime('%Y-%m-%d')
        # print(period)

        # PARAMETERS TO KEEP 

        # parameters actually present in lab results
        # parameters_needed_and_available_from_lab = list(set(data.BESCHREIBUNG)) # this works but CHANGES ORDER

        # this contains the parameters that are needed AND available from the lab results, in the same order of all_needed_parameters
        parameters_needed_and_available_from_lab = [ p for p in all_needed_parameters if p in list(data.BESCHREIBUNG )]

        # parameters_needed_and_available_from_lab is equal to data.BESCHREIBUNG, without repetitions
        parameters_needed_but_not_available_from_lab = [p for p in all_needed_parameters if p not in list(data.BESCHREIBUNG)]

        # for parameter in data.BESCHREIBUNG:
        #   if parameter not in parameters_needed_and_available_from_lab:
        #       data.drop(parameter, axis = 0, inplace = True)


        # data.sort_values(by=['VALIDIERTDAT'], inplace = True) 
        # print(f'\nSorted by date \n {data}\n')

        # Start building dictionary

        # This worked with parameter as index
        # def get_results(parameter):
        #   try:
        #       # if there is more than one result
        #       return list(data.loc[parameter].ERGEBNIST)
        #   except:
        #       # if there is only one result
        #       return [data.loc[parameter].ERGEBNIST]

        # def get_dates(parameter):
        #   try:
        #       # if there is more than one result
        #       return list(data.loc[parameter].VALIDIERTDAT)
        #   except:
        #       # if there is only one result
        #       return [data.loc[parameter].VALIDIERTDAT]

        def get_results(parameter):
            return list(data[data.BESCHREIBUNG == parameter].ERGEBNIST)

        def get_dates(parameter):
            return list(data[data.BESCHREIBUNG == parameter].DAY)


        final_dictionary = {}
        for p in all_needed_parameters:
            if p in parameters_needed_and_available_from_lab:
                final_dictionary[p] = [get_results(p), get_dates(p)]

            elif p in parameters_needed_but_not_available_from_lab:
                final_dictionary[p] = [ [empty_result for _ in range(num_max_days)] , 0 ]

            else:
                raise Exception('Something wrong')


        #print(final_dictionary)
        # Add empty in day when exam is not done
        for p in parameters_needed_and_available_from_lab:
            results, dates = final_dictionary[p]
            #print(dates)
            if len(results) < num_max_days:
                for i in range(num_max_days):
                    if period[i] not in dates:
                        #print(f'{period[i]} not in {dates}')
                        results.insert(i,empty_result)

        # Collect all results; here final dictionary still contains dates, and [0] gets rid of it
        # patient_results = [ final_dictionary[p][0] for p in all_needed_parameters ]
        # big_data.append(patient_results)



        # Make dictionaries for multiple sheets
        final_dictionary_without_dates = { k:v[0] for k,v in final_dictionary.items() }

        # This list contains as many dictionaries as number of sheets, each to be treated as final_dictionary
        list_of_patient_dictionaries = dict_of_lists_to_list_of_dicts( dictionary_values_splitter( final_dictionary_without_dates ) )

        for s in range(number_of_sheets):
            big_data_multiple_sheets[s].append( list(list_of_patient_dictionaries[s].values()) )

        # each element of big_data_multiple_sheets is to be treated as big_data

        patient_identifier_PATIFALLNR.append(data['PATIFALLNR'][0])


####################################################################################################################################################################################
# END DATA MANIPULATION ROUTINE FOR EACH PATIENT
####################################################################################################################################################################################
print('\n Patient loop completed!')
####################################################################################################################################################################################
# START WRITING TO FINAL SHEET
####################################################################################################################################################################################
print('\n Creating patient map, first step...')
patient_identifier_PATIFALLNR_last_digit_separated = [ f'{str(i)[:-1]}_{str(i)[-1:]}' for i in tqdm(patient_identifier_PATIFALLNR)  ]
print('\n Creating patient map, second step...')
patient_identifier_final = [ f'{patID2Num( patient_identifier_PATIFALLNR[i] )} - {patient_identifier_PATIFALLNR_last_digit_separated[i]}' for i in tqdm(range(len(patient_identifier_PATIFALLNR_last_digit_separated))) ]

dims = ['patient', 'parameter', 'day']

def save_excel_sheet(df, dirname, filename, sheetname):

    filepath = f'{dirname}/{filename}'
    os.makedirs(dirname, exist_ok=True)

    if not os.path.exists(filepath):
        with pd.ExcelWriter(filepath) as writer:
            df.to_excel(writer, sheet_name=sheetname)
    else:
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheetname)


print('\n Starting to write in Excel sheets...')
for s in tqdm(range(number_of_sheets)):
    #coords = {'patient':range(current_patient_one, current_patient_one+num_patients), 'parameter':all_needed_parameters, 'day':split_days[s] }
    coords = {'patient':patient_identifier_final, 'parameter':all_needed_parameters, 'day':split_days[s] }
    data = xr.DataArray(big_data_multiple_sheets[s], dims = dims, coords = coords )
    df = data.to_dataframe('value')
    df = data.to_series().unstack(level=[1,2])

    # filename = f'{current_patient_one}-{current_patient_one+num_patients-1}-{current_date}.xlsx'
    filename = f'{current_date}.xlsx'

    save_excel_sheet(df, directory_final_sheet, filename, f'Sheet{s+1}')


print('\nALL GOOD :)\n')






