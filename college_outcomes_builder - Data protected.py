import pandas as pd 
import PySimpleGUI as sg 
import xlsxwriter


#Set dataframe viewing options
pd.set_option('display.max_columns', 999)
pd.set_option('display.max_rows', 999)



##########################


def get_ETO_data(_ETOpath):	#Retrieve and clean ETO data

	#Get data from excel file, fix it up
	df_data = pd.read_excel(_ETOpath, 'Post Secondary Planning', skiprows=3)



	# df_data = df_data.loc[:, ~df_data.columns.str.contains('^Unnamed')] # save this for future clean?
	
	df_data = df_data.rename(columns = {'Primary Org / Program Site': 'School', 'Applied to College?':'Applied?', 'On Time Graduation Year':'Cohort', 'Accepted to College?':'Accepted?', 'Committed College?':'Committed?', 'Post Secondary Plan': 'PSP', 'Service Date':'Service_Date', 'Committed College Name':'Committed_College_Name'})
	df_data['Name'] = df_data['Name'].str.upper()


	#Provide missing data. ex: if there is a comitted college, fill in other info like psp
	df_data.loc[(df_data['Committed_College_Name'].notnull()) & (df_data['Committed?'].isnull()), 'Committed?'] = 'Yes' # if there is a college name, put committed as yes
	df_data.loc[(df_data['Committed?'] == 'Yes'), 'Accepted?'] = 'Yes' # if committed is yes, make accepted yes
	df_data.loc[(df_data['Accepted?'] == 'Yes'), 'Applied?'] = 'Yes' # if accepted is yes, make applied yes
	df_data.loc[(df_data['Applied?'] == 'Yes') & (df_data['PSP'].isnull()), 'PSP'] = 'College' # if applied is yes, and there is no PSP, make it college

	# rename schools to match the dict
	df_data.loc[(df_data['School'] == 'Curtis High School'), 'School'] = 'CHS'
	df_data.loc[(df_data['School'] == 'Bronx Career and College Prep'), 'School'] = 'BCCP'
	df_data.loc[(df_data['School'] == 'Fannie Lou Hamer'), 'School'] = 'FLH'

	# print(df_data)
	# quit()

	return df_data

def build_demographic_breakdown(df, n, status, phrase1, phrase2, col1, col2, col3, col4): # build the breakdown of grads, applicants... etc by ethnicity and other demos

	def build_demographics(df, n, col1, col2, col3):
		# gender
		d = {col1: ['','Gender:'], 
			col2: ['','']
			}
		df_temp = pd.DataFrame(data=d)
		df_gen = df['Gender'].value_counts().rename_axis(col1).reset_index(name=col2)
		df_gen[col3] = round(df_gen[col2] / n, 3)
		df_gen = pd.concat([df_temp, df_gen])

		# SWD
		d = {col1: ['','SWD:'], 
			col2: ['','']
			}
		df_temp = pd.DataFrame(data=d)
		df_SWD = df['SWD'].value_counts().rename_axis(col1).reset_index(name=col2)
		df_SWD[col3] = df_SWD[col2] / n
		df_SWD = pd.concat([df_temp, df_SWD])

		# ML
		d = {col1: ['','ML:'], 
			col2: ['','']
			}
		df_temp = pd.DataFrame(data=d)
		df_ML = df['ML'].value_counts().rename_axis(col1).reset_index(name=col2)
		df_ML[col3] = df_ML[col2] / n
		df_ML = pd.concat([df_temp, df_ML])

		df = pd.concat([df_gen, df_SWD, df_ML]) # final df for c23 eth

		return df

	# get ethnicity by status (cohort, applied, accepted... etc)
	d = {col1: [phrase1, '', phrase2], # input column headers (phrases)
		col2: [n, '', ''], # n is the reqired number for phrase, num of cohort, num of applied...
		col3: ['', '', ''],
		col4: ['', '', '']
		}
	df_ethSummary = pd.DataFrame(data=d) # temp df

	if status == False: # if grads or cohort
		df_eth = df['Ethnicity / Race'].value_counts().rename_axis(col1).reset_index(name=col2) # summarize by eth
		df_eth[col3] = df_eth[col2] / n # add percentages

		df_demo = build_demographics(df, n, col1, col2, col3) # add to df_demo

	else: # if one of the applied to collges
		df_temp = df[(df[status] == 'Yes')] # filter to yes if its applied acc or comm

		df_eth = df_temp['Ethnicity / Race'].value_counts().rename_axis(col1).reset_index(name=col2) # summarize by eth
		df_eth[col3] = df_eth[col2] / n # create percentages

		df_demo = build_demographics(df_temp, n, col1, col2, col3)

	df_eth = pd.concat([df_ethSummary, df_eth, df_demo]) # final df for applied eth

	return df_eth

##################333main

def main(_ETOpath, _CHSpath, _BCCPpath, _FLHpath, _FDApath, _gspath):

	#Variables
	schools = ['CHS', 'BCCP', 'FLH', 'FDA'] # to use in file sav ename
	school_files = { # to store file paths
		'FDA':_FDApath,
		'CHS':_CHSpath,
		'BCCP':_BCCPpath,
		'FLH':_FLHpath
		
	}

	####################################################333 ETO MAKE AND MERGE

	df_eto_og = get_ETO_data(_ETOpath) # pull in eto file and clean

	# print(df_eto[(df_eto['Name'] == 'Coley, Justin')])
	# quit()

	#Set parameters get from input...
	cohort = 2023 # set cohort to be analyzed

	for school in school_files:

		# df_eto = df_eto_og[(df_eto_og['School'] == school)]

		nv_file = school_files[school]

		if len(nv_file) > 1: # if file path isnt empty
			if school == 'FDA':
				df_gs = pd.read_excel(_gspath, sheet_name=0)
				df_gs = df_gs.rename(columns = {'Last, First': 'Name'}) #rename columns to remove spaces
				df_gs = df_gs[['Name', 'PSP', 'Applied?', 'Committed?']]
				df_gs['Name'] = df_gs['Name'].str.upper()
				df_gs['School'] = 'FDA'

				df_gs['Committed?'][df_gs['Committed?'].notnull()] = 'Yes' # if there is a school, mark it as yes
				df_gs.loc[(df_gs['Committed?'] == 'Yes'), 'Accepted?'] = 'Yes' # if committed is yes, make accepted yes
				df_gs.loc[(df_gs['Accepted?'] == 'Yes'), 'Applied?'] = 'Yes' # if accepted is yes, make applied yes
				df_gs.loc[(df_gs['Applied?'] == 'Yes') & (df_gs['PSP'].isnull()), 'PSP'] = 'College' # if applied is yes, and there is no PSP, make it college

				df_eto_og = pd.concat([df_eto_og, df_gs])

			
			df_eto = df_eto_og[(df_eto_og['School'] == school)]


			print('started')
			df_nv = pd.read_excel(nv_file, sheet_name=0) # read first sheet



			df_og = df_nv # save original export

			df_nv = df_nv.rename(columns = {'On Track in Regents for Planned Diploma': 'Regents_OnTrack', 'Total Credits Earned': 'Total_Credits', 'Student Name': 'Name'}) #rename columns to remove spaces

			df_cohort = df_nv[(df_nv.Class == cohort)] #get correct cohort, save edit in cohort df

			df_cohort = df_cohort[['Student ID', 'Name', 'Status', 'Class', 'Grade', 'Gender', 'Ethnicity / Race', 'SWD', 'ML', 'Housing Status', 'DOE Attendance Risk Group', 'Admit Date', 'Discharge Code', 'Discharge Date', 'Total_Credits', 'Regents_OnTrack' ]] #use only needed columns

			# find existing names first. Pull in dates from  eto file to see who was even served. Doesnt matter what date
			df_cohort = pd.merge(df_cohort, df_eto.drop_duplicates(subset=['Name'])[['Name', 'Service_Date']], on='Name', how='left')

			# remove psp blanks from eto file, keep nearest date
			df_psp = df_eto[df_eto['PSP'].notna()]

			df_psp = (df_psp
					 .assign(cat=pd.Categorical(df_psp['PSP'], categories=['Other', 'College'], ordered=True))
					 .sort_values(by=['Service_Date', 'cat'], na_position='first')
					 .drop(columns='cat')
					 .groupby('Name', as_index=False).last()
					 )

			df_cohort = pd.merge(df_cohort, df_psp[['Name', 'PSP']], on='Name', how='left')

			def merge_appstati(col):
				# ask stack later if this is best way, i think there is better
				df = df_eto[(df_eto[col] == 'Yes')]

				df = df.drop_duplicates(subset=['Name'])

				return pd.merge(df_cohort, df[['Name', col]], on='Name', how='left')

			df_cohort = merge_appstati('Applied?')
			df_cohort = merge_appstati('Accepted?')
			df_cohort = merge_appstati('Committed?')

			# ask stack later if this is best way, i think there is better ^^^
			
			# print(df_cohort['Name'])#[(df_cohort['Name'] == 'Coley, Justin')])
			# quit()

			df_cohort = df_cohort.rename(columns = {'Service_Date' : 'Served_ETO?'}) #rename date col to served?
			df_cohort['Served_ETO?'][df_cohort['Served_ETO?'].notnull()] = 'Yes' # if there is a date, mark it as yes

			df_cohort['Expected_Grad?'] = None


			# if already graduated, they are expected grad...
			df_cohort.loc[(df_cohort['Status'] == 'D-grad'), 'Expected_Grad?'] = 'Yes'

			if school != 'FLH':
				df_cohort.loc[((df_cohort['Status'] == 'A') | (df_cohort['Status'] == 'A*')) & ((df_cohort['Regents_OnTrack'] == 'On Track') | (df_cohort['Regents_OnTrack'] == '1 Regent behind')) & (df_cohort['Total_Credits'] >= 37), 'Expected_Grad?'] = 'Yes'
				# df_cohort.loc[((df_cohort['Regents_OnTrack'] == 'On Track') | (df_cohort['Regents_OnTrack'] == '1 Regent behind')) & (df_cohort['Total_Credits'] >= 37), 'Expected_Grad?'] = 'Yes'
			else:
				df_cohort.loc[((df_cohort['Status'] == 'A') | (df_cohort['Status'] == 'A*')) & (df_cohort['Total_Credits'] >= 37), 'Expected_Grad?'] = 'Yes'

			df_cohort = df_cohort.fillna({'Served_ETO?': 'No', 'PSP':'No Plan', 'Applied?':'No', 'Accepted?':'No', 'Committed?':'No', 'Expected_Grad?':'No'}) # fill in na's, cohort df done


			df_cohort = add_manual_edits(df_cohort, school)

			df_grads = df_cohort #[((df_cohort.Status == 'A') | (df_cohort.Status == 'D-grad'))]# & df_cohort['Expected_Grad?'] == 'Yes'] # get active or grads only
			df_grads = df_grads[(df_grads['Expected_Grad?'] == 'Yes')]


			############################################################################################## Summaries

			df_served = df_cohort[(df_cohort['Served_ETO?'] != 'No')]

			n_cohort = len(df_cohort)
			expected_grads = len(df_grads)

			# n_psps = df_served['PSP'][(df_served['PSP'] != 'No Plan')].value_counts().sum()
			# n_gradsServed = df_grads['Served_ETO?'][(df_grads['Served_ETO?'] == 'Yes')].count()
			n_served = len(df_served)
			#df_cohort['Served_ETO?'].value_counts()['Yes']

			n_psps = df_grads['PSP'][(df_grads['PSP'] != 'No Plan')].value_counts().sum()
			n_app = df_grads['Applied?'].value_counts()['Yes']
			n_acc = df_grads['Accepted?'].value_counts()['Yes']
			n_comm = df_grads['Committed?'].value_counts()['Yes']

			df_collegePSP = df_grads[(df_grads['PSP'] == 'College')]

			n_collegePSPs = len(df_collegePSP)
			n_Capp = df_collegePSP['Applied?'].value_counts()['Yes']
			n_Cacc = df_collegePSP['Accepted?'].value_counts()['Yes']
			n_Ccomm = df_collegePSP['Committed?'].value_counts()['Yes']


			d = {'col1': ['Total Cohort 23:', 'Expected Grads:', 'Served Cohort (ETO):', 'Students with PSP:', '', 'Post Secondary Plans (of Grads):'], 
				'col2': [n_cohort, expected_grads, n_served, n_psps, '', ''], 
				'col3': ['', round(expected_grads/n_cohort, 3), round(n_served/n_cohort, 3), round(n_psps/expected_grads, 3), '',''], 
				'col4': ['','of cohort expected to graduate','of cohort were served with college access', 'of grads have a psp','','']
				}

			df_summary = pd.DataFrame(data=d)

			# df_plans = df_joined['PSP'][(df_joined['PSP'] != 'No Plan')].value_counts().rename_axis('col1').reset_index(name='col2')
			df_plans = df_grads['PSP'].value_counts().rename_axis('col1').reset_index(name='col2')
			df_plans['col3'] = df_plans['col2'] / expected_grads

			df_summary = pd.concat([df_summary, df_plans])

			d = {'col1': ['', 'RobinHood Criteria', 'Applied:', 'Accepted:', 'Committed:', '', "Normal Criteria (Expected Grads with 'College' as a PSP)", 'Applied:', 'Accepted:', 'Committed:'], 
				'col2': ['', '', n_app, n_acc, n_comm, '', '', n_Capp, n_Cacc, n_Ccomm], 
				'col3': ['', '', round(n_app/expected_grads, 3), round(n_acc/expected_grads, 3), round(n_comm/n_acc, 3), '', '', round(n_Capp/n_collegePSPs, 3), round(n_Cacc/n_Capp, 3), round(n_Ccomm/n_Cacc, 3)], 
				'col4': ['', '', 'of expected grads applied to college' ,'of expected grads accepted to college','of accepted committed to college', '', '', 'of grads with college PSP applied to college' ,'of previous applied accepted to college','of previous accepted committed to college']
				}

			df_percents = pd.DataFrame(data=d)

			df_summary = pd.concat([df_summary, df_percents])

			################################################################-- RH racial summary 

			df_cohort_eth = build_demographic_breakdown(df_cohort, n_cohort, False, 'Total 2023 Cohort:', 'Cohort by Ethnicity:', 'col1', 'col2', 'col3', 'col4')

			df_grads_eth = build_demographic_breakdown(df_grads, expected_grads, False, 'Total 2023 Expected Grads:', 'Grads by Ethnicity:', 'col5', 'col6', 'col7', 'col8')

			df_app_eth = build_demographic_breakdown(df_grads, n_app, 'Applied?', 'Total Students who Applied to College:', 'Students Who Applied by Ethnicity:', 'col9', 'col10', 'col11', 'col12' )

			df_acc_eth = build_demographic_breakdown(df_grads, n_acc, 'Accepted?', 'Total Students who were Accepted to College:', 'Students Who were accepted by Ethnicity:', 'col13', 'col14', 'col15', 'col16')

			df_comm_eth = build_demographic_breakdown(df_grads, n_comm, 'Committed?', 'Total Students who Committed to College:', 'Students Who Committed by Ethnicity:', 'col17', 'col18', 'col19', 'col20')

			df_psp_temp = df_grads[(df_grads['PSP'] != 'No Plan')]

			df_psp_eth = build_demographic_breakdown(df_psp_temp, n_psps, False, 'Total Students with PSP:', 'Grads by Ethnicity:', 'col21', 'col22', 'col23', 'col24')



			df_ethSummary = pd.concat([df_cohort_eth.reset_index(drop=True), df_grads_eth.reset_index(drop=True), df_app_eth.reset_index(drop=True), df_acc_eth.reset_index(drop=True), df_comm_eth.reset_index(drop=True), df_psp_eth.reset_index(drop=True)], axis=1)

			############################################################## - target students, not served, no plan, not applied

			'''
			not served in ETO
			not expected to graduate
			served but no PSP
			served but not applied



			'''

			# get not served
			d = {'col1': ['Students Not Served in ETO:']}

			# cols = ['col1', 'Student ID', 'Name', 'Status', 'Class', 'Grade', 'Gender', 'Ethnicity / Race', 'SWD', 'ML', 'Housing Status', 'DOE Attendance Risk Group', 'Admit Date', 'Discharge Code', 'Discharge Date', 'Total_Credits', 'Regents_OnTrack', 'Served_ETO_SY22-23?', 'PSP', 'Applied?', 'Accepted?', 'Committed?' ]

			df_temp = pd.DataFrame(data=d) # temp df

			df_target_notserved = df_cohort[(df_cohort['Served_ETO?'] == 'No')]
			# df_target_notserved.insert(loc=0, column='col1', value='')
			df_target_notserved = pd.concat([df_temp, df_target_notserved])

			# grads, but no psp plan
			d = {'col1': ['', 'Expected Grads without a Post-Secondary Plan:']}
			df_temp = pd.DataFrame(data=d) # temp df

			df_target_noplan = df_grads[(df_grads['PSP'] == 'No Plan')]

			df_target_noplan = pd.concat([df_temp, df_target_noplan])

			# grads, but not applied
			d = {'col1': ['', 'Expected Grads who have not applied to college:']}
			df_temp = pd.DataFrame(data=d) # temp df

			df_target_notapp = df_grads[(df_grads['Applied?'] == 'No')]

			df_target_notapp = pd.concat([df_temp, df_target_notapp])


			df_targets = pd.concat([df_target_notserved, df_target_noplan, df_target_notapp])


			########################################################## - Attrition data



			df_dneg = df_nv[(df_nv.Status == 'D-neg')]

			# print(df_dneg['Discharge Date'].str[:-33])
			# quit()

			df_dneg['Discharge Date'] = pd.to_datetime(df_dneg['Discharge Date'].str[:-33], format='%a %b %d %Y %H:%M:%S') # Fri Jul 01 2022 00:00:00 GMT-0400 (Eastern Daylight Time)

			df_dneg = df_dneg[(df_dneg['Discharge Date'] >= '2022-09-01') & (df_dneg['Discharge Date'] <= '2023-09-01')]

			# write title for dneg
			d = {'col1': ['Total Negative Disharges:', '', 'Discharge Code'],
				'col2': [len(df_dneg), '', 'n'],
				'col3': ['', '', '%']
			}

			df_temp = pd.DataFrame(data=d) # temp df

			df_attrn = df_dneg['Discharge Code'].value_counts().rename_axis('col1').reset_index(name='col2')

			df_attrn['col3'] = df_attrn['col2'] / len(df_dneg)

			df_attrn = pd.concat([df_temp, df_attrn])






			###########################################################

			filesave = str(school) + '_CollegeSummary.xlsx'

			with pd.ExcelWriter(filesave) as writer:  

				# xlsx write vars
				workbook = writer.book
				bold = workbook.add_format({'bold': True})
				bold_and_percent = workbook.add_format({'num_format': '0.0%', 'bold': True})

				# tab for overall summary
				df_summary.to_excel(writer, sheet_name='Summary', index=False, header=False) # created summary, first tab

				bolded_cols = ['col2']
				percent_col = ['col3']

				for column in df_summary: # Auto-adjust columns' width 
					column_width = max(df_summary[column].astype(str).map(len).max(), len(column))
					col_idx = df_summary.columns.get_loc(column)

					if column in bolded_cols:
						writer.sheets['Summary'].set_column(col_idx, col_idx, column_width, bold)
					elif column in percent_col:
						writer.sheets['Summary'].set_column(col_idx, col_idx, 6.22, bold_and_percent)
					else:
						writer.sheets['Summary'].set_column(col_idx, col_idx, column_width)

				# tab for rh ethnic breakdown
				df_ethSummary.to_excel(writer, sheet_name='RH Ethnicity Breakdown', index=False, header=False)

				bolded_cols = ['col2', 'col6', 'col10', 'col14', 'col18', 'col22']
				bold_and_percent_cols = ['col3', 'col7', 'col11', 'col15', 'col19', 'col23']

				for column in df_ethSummary: # Auto-adjust columns' width 
					column_width = max(df_ethSummary[column].astype(str).map(len).max(), len(column))
					col_idx = df_ethSummary.columns.get_loc(column)
					if column in bolded_cols:
						writer.sheets['RH Ethnicity Breakdown'].set_column(col_idx, col_idx, column_width, bold)
					elif column in bold_and_percent_cols:
						writer.sheets['RH Ethnicity Breakdown'].set_column(col_idx, col_idx, 5.22, bold_and_percent) 
					else:				
						writer.sheets['RH Ethnicity Breakdown'].set_column(col_idx, col_idx, column_width)

				# tab for target students
				df_targets.to_excel(writer, sheet_name='Target Students', index=False)

				for column in df_targets: # Auto-adjust columns' width
					column_width = max(df_targets[column].astype(str).map(len).max(), len(column))
					col_idx = df_targets.columns.get_loc(column)
					if column == 'col1':
						writer.sheets['Target Students'].set_column(col_idx, col_idx, 5, bold)
					else:
						writer.sheets['Target Students'].set_column(col_idx, col_idx, column_width)

				# tab for attrition data
				df_attrn.to_excel(writer, sheet_name='Student Attrition SY22-23', index=False, header=False)

				for column in df_attrn: # Auto-adjust columns' width
					column_width = max(df_attrn[column].astype(str).map(len).max(), len(column))
					col_idx = df_attrn.columns.get_loc(column)
					if column == 'col2':
						writer.sheets['Student Attrition SY22-23'].set_column(col_idx, col_idx, column_width, bold)
					elif column == 'col3':
						writer.sheets['Student Attrition SY22-23'].set_column(col_idx, col_idx, 5.22, bold_and_percent)
					else:
						writer.sheets['Student Attrition SY22-23'].set_column(col_idx, col_idx, column_width)

				df_dneg.to_excel(writer, sheet_name='Student Attrition SY22-23', startrow=len(df_attrn)+1, index=False)


				# tab for those meeting expected graduates criteria
				df_cohort.to_excel(writer, sheet_name='Cohort Data', index=False) # expected grads tab

				df = df_cohort
				worksheet = writer.sheets['Cohort Data']
				(max_row, max_col) = df.shape
				column_settings = [{"header": column} for column in df.columns]
				worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
				worksheet.set_column(0, max_col - 1, 12)

				# tab for original data
				df_og.to_excel(writer, sheet_name='Original Export', index=False) # original export

				df = df_og
				worksheet = writer.sheets['Original Export']
				(max_row, max_col) = df.shape
				column_settings = [{"header": column} for column in df.columns]
				worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
				worksheet.set_column(0, max_col - 1, 12)
				writer.close()

		print('finished')


#-------------------------------- gui

sg.theme('BlueMono') # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('-Please make sure files are saved as .xlsx and not .csv')],
			[sg.Text('-Please make sure the first tab of the excel is the data grid')],
			[sg.Text('-In order to get file path, SHIFT + RIGHT CLICK on a file and select "Copy as Path"')],
			[sg.Text('Enter ETO File Path'), sg.InputText(key='-IN1-')],
            [sg.Text('Enter CHS NV Export'), sg.InputText(key='-IN2-')], 
            [sg.Text('Enter BCCP NV Export'), sg.InputText(key='-IN3-')], 
            [sg.Text('Enter FLH NV Export'), sg.InputText(key='-IN4-')], 
            [sg.Text('Enter FDA NV Export'), sg.InputText(key='-IN5-')], 
            [sg.Text('Enter FDA GS Export'), sg.InputText(key='-IN6-')],
            [sg.Button('Run Program'), sg.Push(), sg.Button('Cancel')] ]

# Create the Window
window = sg.Window('High School Outcomes', layout)

# Event Loop to process "events" and get the "values" of the inputs
while True:
	event, values = window.read()

	if event == 'Run Program':
		# real file path entries
		eto_file = values['-IN1-'].strip('"')
		CHS_file = values['-IN2-'].strip('"')
		BCCP_file = values['-IN3-'].strip('"')
		FLH_file = values['-IN4-'].strip('"')
		FDA_file = values['-IN5-'].strip('"')
		gs_file = values['-IN6-'].strip('"')

		main(eto_file, CHS_file, BCCP_file, FLH_file, FDA_file, gs_file)

		window.close()

	if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
		break

window.close()





