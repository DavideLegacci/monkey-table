import pandas as pd
import os
from datetime import datetime
import numpy as np
import xarray as xr


mode = 'test'
# mode = 'full'

if mode == 'full':
	current_patient_one = int(input('\nWhich is the number of the first patient? Type a number and hit enter:  '))
	keep_kein_material = ''
	while keep_kein_material not in ['y', 'n']:
		keep_kein_material = input("\nShould I keep 'Kein Material' results? Type y or n: ")

elif mode == 'test':
	current_patient_one = 5
	keep_kein_material = 'n'

empty_result = ''
positive_result = '1'
negative_result = ''

####################################################################################################################################################################################

lab_results_directory = './lab_results'
num_max_days = 23
sub_period_duration = 7

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

def dictionary_values_splitter(dictionary_to_split):
	return { k:data_splitter(v) for k,v in dictionary_to_split.items() }

def dict_of_lists_to_list_of_dicts(dict_of_lists):
	return [ dict(zip(dict_of_lists, i)) for i in zip(*dict_of_lists.values()) ]

# <---------------------------------------------------- the parameters marked red in the numbers sheet must be added?
# This contains all 29 parameters in the order or Rebecca's work sheet; those marked red have -temp names, meaning I could not find them in lab sheet.
# Must find a lab sheet that contains them, and see how they are called

# TEST


# FIRST 29
first_29_parameters  = ['tacro-temp', 'ciclo-temp', 'Natrium(ISE)', 'Kalium (ISE)', 'Calcium', 'Kreatinin', 'Proenkephalin', 'GFR, CKD-EPI', 'Harnstoff', 'Glucose', 'LDH', 'GOT/AST', 'GPT/ALT', 'AP', 'GGT', 'bili-temp', 'Phosphat', 'Ges.Eiweiss', 'Albumin quant.', 'CRP', 'Leukozyten', 'Hb', 'Thrombozyten', 'ntpro-temp', 'tnt-temp', 'INR - ber.', 'Quick', 'aPTT', 'ipth-temp']

# NEW 17
new_17_parameters = ['pH/Tstr.', 'Glucose/Tstr.', 'Bili/Tstr.', 'Ketone /Tstr.', 'Erys /Tstr.', 'Eiweiß/Tstr.', 'Urobil /Tstr.', 'Nitrit /Tstr.', 'Leuko /Tstr.', 'U-Albumin', 'Protein/Urin', 'Eiweiss-temp', 'Erys/µl', 'Leuko/µl', 'platten-temp', 'Bakt./Sedu.', 'HyalZy./Sedu.']

if mode == 'full':
	all_needed_parameters = first_29_parameters + new_17_parameters

elif mode == 'test':
	all_needed_parameters =['a', 'b', 'c']

# big_data = [  ]
num_patients = 0

# Multiple sheets

number_of_sheets = int(num_max_days/sub_period_duration) + 1

big_data_multiple_sheets = [  [ ] for _ in range(number_of_sheets)  ]

# print

# START PATIENT

# Read each excel file in lab_results_directory into dataframe and put it into data_list
for patient in os.listdir(lab_results_directory):
	# make sure to select only excel files; sometimes hidden files like ~$patient.xlsx are created, which must be excluded:
	if patient.endswith(".xlsx") and not patient.startswith("~"):

		num_patients +=1
		print(patient)

		# Parse dates makes Datum column into date objects
		# correctly read time as dd.mm.yy to yy-mm-dd VERIFY
		# custom_date_parser = lambda x: datetime.strptime(x, "%d.%m.%y")
		# data = pd.read_excel(f'{lab_results_directory}/{patient}', index_col = 0, skiprows = lambda x: x in [0, 2], parse_dates=['Datum '], date_parser=custom_date_parser)

		# START MERGING DATAFRAME
		# Read data into two df
		data1 = pd.read_excel(f'{lab_results_directory}/{patient}', skiprows = 2, usecols = 'A, B, C, D')
		data2 = pd.read_excel(f'{lab_results_directory}/{patient}', skiprows = 2, usecols = 'F, G, H, I')

		# Rename columns of second to match columns of first
		for i in range(len(data1.columns)):
			data2.rename(columns={ data2.columns[i]: data1.columns[i] }, inplace = True)

		# Merge into single
		# ignore_index makes index range from 0 to n-1 rather than keeping original indice; dropna() drops rows with NaN
		data = pd.concat( [data1, data2], ignore_index=True ).dropna()


		# Print full dataframe
		pd.set_option("display.max_rows", None, "display.max_columns", None)
		# END MERGING DATAFRAME



		# THIS IS INCLUDED IN CYCLE BELOW: IF PARAM NOT IN ALL_NEEDED_PARAMATERS, DROP IT
		# Get rid of extra Parameter rows
		# try:
		# 	data.drop('Parameter', inplace = True)
		# except:
		# 	pass


		# Get rid of spaces in columns
		data.columns = data.columns.str.replace(' ', '')

		# Get rid of » symbol in indiced
		data.Parameter = data.Parameter.str.replace(' »', '')

		# Drop not needed parameters
		# THIS WORKS IS PARAMETERS ARE AXIS
		# for param in data.Parameter:
		# 	if param not in all_needed_parameters:
		# 		try:
		# 			data.drop(param, inplace = True) #<----------- MAIN DROP
		# 		except:
		# 			pass

		# IF INSTEAD PARAMETERS ARE VALUES OF COLUMN # <----------- MAIN DROP
		# https://stackoverflow.com/questions/18172851/deleting-dataframe-row-in-pandas-based-on-column-value
		for param in data.Parameter:
			if param not in all_needed_parameters:
				data.drop(data.index[ data.Parameter == param ], inplace = True) 

				# Should be allright without this
				# try:
				# 	data.drop(data.index[ data.Parameter == param ], inplace = True)
				# except:
				# 	pass

		# Since some rows were dropped, not index looks like [0, 1, 5, 9, 15, ...]
		# Which is a mess because data.colum[i] refers to that index. So need to reset.
		data.reset_index(inplace = True, drop = True)

		# Now not needed parameters are dropped from data.Parameter. It may still happen that a needed parameter is not present in result. Fixed later. 


		# drop norm column
		data.drop('Normb / Dimension', axis = 1, inplace = True)



		# Fix dates format
		for i in range(len(data.Datum)):

			# Some dates were recognized by pandas already as datetimes; other are still strings because 1. they contain spaces and 2. they contain . rather than /
			if type( data.at[i, 'Datum'] ) is str:
				data.at[i, 'Datum'] = data.at[i, 'Datum'].replace(' ', '')
				data.at[i, 'Datum'] = data.at[i, 'Datum'].replace('.', '/')	

		# Convert to datetime. Those that already are, are not affected
		data['Datum'] = pd.to_datetime(data['Datum'], dayfirst = True)



		# <------------------------------------------------------------------------- This raises error if something is left in results column which is not a number, as +, - poisitiv, negativ, kein material, ...
		# Ask Rebecca whether such results can appear in the parameters she needs, and start by getting rid of k.m.
		#print(data)
		#START GETTING RID OF kein material ROWS
		if keep_kein_material == 'n':
			strings_to_kill = ['Kein Material', 'K.Mat.']
			for s in strings_to_kill:
				index_to_kill = data.index[ data.Wert == s ]
				[ print(f"---------------Dropping {s} for {data.at[ i , 'Parameter']}") for i in list(index_to_kill) ]
				data.drop(index_to_kill, inplace = True)

			data.reset_index(inplace = True, drop = True)

		# END GETTING RID OF kein material ROWS

		# NON STANDARD RESULTS
		# + and positive to 1
		# - and negative to empty
		# anything else remains text

		# Get rid of !L and !H flags
		for i in range(len(data.Wert)):
			if type( data.at[i, 'Wert'] ) is str:
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace(' !L', '')
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace(' !H', '')
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace('negativ', negative_result)
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace('-', negative_result)
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace('positiv', positive_result)
				data.at[i, 'Wert'] = data.at[i, 'Wert'].replace('+', positive_result)
		# END NON STANDARD RESULTS


		# Set values as float; if not (i.e. if strinsg) the xarray is messed up bc/ interpretes 7.21 as '7', '.', '2', '1'
		# data = data.astype({"Wert": float}) # <-------------------------------------- this raises error if values are (incorrectly) interpreted by excel as dates. Temp fix by setting text type in excel
		data = data.astype({"Wert": str}) # <-------------------------------------- this raises error if values are (incorrectly) interpreted by excel as dates. Temp fix by setting text type in excel


		# START GETTING RID OF DUPLICATE EXAM <------------------------------------------------------------------------------------------------- WORKING HERE.
		# POSSIBILITIES:
		
		# 1. Restructure all so that parameters are column, not index, and use
		# https://stackoverflow.com/questions/50885093/how-do-i-remove-rows-with-duplicate-values-of-columns-in-pandas-data-frame

		# 2. Keep parameters as index, work on sub-frames for each parameter, drop duplicate date column
		# Then reconstruct big dataframe by composition

		# OPTION 1
		#data.drop_duplicates( ['Parameter', 'Datum'], keep = 'first', inplace = True, ignore_index = True  ) # <------------------------ to improve: allow choice of value to keep
		for p in set(data.Parameter):
			df_specific_for_p = data.loc[ data.Parameter == p ]

			# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.duplicated.html
			boolean_duplicated_series = df_specific_for_p.duplicated( ['Parameter', 'Datum'], keep = False )

			if( boolean_duplicated_series.any() ):
				df_duplicated_p = df_specific_for_p[boolean_duplicated_series]
				print()
				print(df_duplicated_p)
				s = f'\nThere are more exams for {p} taken on the same day. Type index to KEEP: '
				I = input(s)
				
				for i in df_specific_for_p.index:
					if str(i) != str(I):
						data.drop(i, inplace = True)
						
		data.reset_index(inplace = True, drop = True)

		# END GETTING RID OF DUPLICATE EXAM

		# END MANIPULATING DATAFRAME


		# CAREFUL ACTUALLY DAY0 IS FROM EXTERNAL SOURCE, IT MAY BE THAT NO EXAM IS TAKEN ON DAY 0 <------------------------------------------------ temporary
		# Get patient day0 = when she enters hospital

		day0 = min(data.Datum)

		# Get range of time in which exams are taken
		day_first_exam, day_last_exam = min(data.Datum), max(data.Datum)
		exam_period = day_last_exam - day_first_exam
		# print(exam_period)

		# Check time period
		if exam_period > pd.Timedelta(num_max_days, unit = 'd'):
			print(f'\n\n-----------SOMETHING WRONG---------\n\n Exam period lasts {exam_period}, longer than {num_max_days} days\n----------\n')
			raise Exception

		# Set maximal period of staying in the hospital; equal for everybody
		period = pd.date_range(start=day0, periods=num_max_days)
		# print(period)

		# PARAMETERS TO KEEP 

		# parameters actually present in lab results
		# parameters_needed_and_available_from_lab = list(set(data.Parameter)) # this works but CHANGES ORDER

		# this contains the parameters that are needed AND available from the lab results, in the same order of all_needed_parameters
		parameters_needed_and_available_from_lab = [ p for p in all_needed_parameters if p in list(data.Parameter )]

		# parameters_needed_and_available_from_lab is equal to data.Parameter, without repetitions
		parameters_needed_but_not_available_from_lab = [p for p in all_needed_parameters if p not in list(data.Parameter)]

		# for parameter in data.Parameter:
		# 	if parameter not in parameters_needed_and_available_from_lab:
		# 		data.drop(parameter, axis = 0, inplace = True)


		# data.sort_values(by=['Datum'], inplace = True) # <-------------------------------------------------------- this should not really be necessary, is it? 
		# print(f'\nSorted by date \n {data}\n')

		# Start building dictionary

		# This worked with parameter as index
		# def get_results(parameter):
		# 	try:
		# 		# if there is more than one result
		# 		return list(data.loc[parameter].Wert)
		# 	except:
		# 		# if there is only one result
		# 		return [data.loc[parameter].Wert]

		# def get_dates(parameter):
		# 	try:
		# 		# if there is more than one result
		# 		return list(data.loc[parameter].Datum)
		# 	except:
		# 		# if there is only one result
		# 		return [data.loc[parameter].Datum]

		def get_results(parameter):
			return list(data[data.Parameter == parameter].Wert)

		def get_dates(parameter):
			return list(data[data.Parameter == parameter].Datum)

		final_dictionary = {}
		for p in all_needed_parameters:
			if p in parameters_needed_and_available_from_lab:
				final_dictionary[p] = [get_results(p), get_dates(p)]

			elif p in parameters_needed_but_not_available_from_lab:
				final_dictionary[p] = [ [empty_result for _ in range(num_max_days)] , 0 ]

			else:
				raise Exception('Something wrong')


		# Add empty in day when exam is not done
		for p in parameters_needed_and_available_from_lab:
			results, dates = final_dictionary[p]
			if len(results) < num_max_days:
				for i in range(num_max_days):
					if period[i] not in dates:
							results.insert(i,empty_result)


		# Collect all results; here final dictionary still contains dates, and [0] gets rid of it
		patient_results = [ final_dictionary[p][0] for p in all_needed_parameters ]
		big_data.append(patient_results)

		print()
		print(data)
		print()
		
		print()
		print(big_data)

		# Make dictionaries for multiple sheets
		final_dictionary_without_dates = { k:v[0] for k,v in final_dictionary.items() }

		# This list contains as many dictionaries as number of sheets, each to be treated as final_dictionary
		list_of_patient_dictionaries = dict_of_lists_to_list_of_dicts( dictionary_values_splitter( final_dictionary_without_dates ) )

		for s in range(number_of_sheets):
			big_data_multiple_sheets[s].append(  )


# END PATIENT
# print(big_data)

dims = ['patient', 'parameter', 'day']
coords = {'patient':range(current_patient_one, current_patient_one+num_patients), 'parameter':all_needed_parameters, 'day':range(num_max_days)} # dict-like

data = xr.DataArray(big_data, dims = dims, coords = coords )
df = data.to_dataframe('value')
df = data.to_series().unstack(level=[1,2])
with pd.ExcelWriter("new.xlsx") as writer: 
    df.to_excel(writer)


print('\nALL GOOD :)\n')






