# Importing libraries used
from openpyxl import Workbook as wk
import csv
import re

# Function to find which students did not fill feedback
def feedback_not_submitted():
	
	ltp_mapping_feedback_type = {1: 'lecture', 2: 'tutorial', 3:'practical'}
	# Dictionaries for quick access of information
	# Dictionary rollno: list of subcode registered
	rollno_to_reg_subno={}
	# Dictionary rollno: list of {subcode registered:frequency }
	rollno_to_feedback={}
	# Dictionary rollno: semester info (registered_sem,scheduled_sem,subcode) 
	rollno_to_seminfo={}
	# Dictionary rollno: student info (name,email,aemail,contact) 
	rollno_to_info={}
	# Dictionary subno: no of non zero ltp
	subno_to_ltp={}

	# output sheet	
	wb = wk()
	ws1 = wb.active
	# Add column names in the output sheet
	ws1.append(["rollno","register_sem","schedule_sem","subno","Name","email","aemail","contact"])

	# Files used to create dictionaries
	output_file_name = "course_feedback_remaining.xlsx" 
	reg_subno_file="course_registered_by_all_students.csv"
	course_master_file="course_master_dont_open_in_excel.csv"
	feedback_sub="course_feedback_submitted_by_students.csv"
	inf="studentinfo.csv"

	# creating dictionary rollno: list of {subcode feedback submitted:frequency }
	# Read the file
	with open(feedback_sub, 'r') as file:
		rd=csv.reader(file)
		temp=1
		for details in rd:
			# If it's the first row(columns) then skip
			if temp:
				temp=0
				continue
			# check if the rollno , subcode is present in dict
			# then increase the frequency if it does exist already 
			if details[3] in rollno_to_feedback.keys() and details[4] in rollno_to_feedback[details[3]].keys():
				rollno_to_feedback[details[3]][details[4]]+=1
			# if it shows an error means that it is the
			# first encounter
			else:
				# if rollno is present in dictionary
				# add the subcode with freq 1
				if details[3] in rollno_to_feedback.keys():
						rollno_to_feedback[details[3]][details[4]]=1
				# Otherwise add the rollno and subcode
				else:
					rollno_to_feedback[details[3]]={}
					rollno_to_feedback[details[3]][details[4]]=1

	# creating dictionary rollno: list of {subcode registered}
	# and also dictionary rollno: semester info (registered_sem,scheduled_sem,subcode) 
	# Read the file
	with open(reg_subno_file, 'r') as file:
		rd=csv.reader(file)
		for details in rd:
			# if the rollno is already present in the keys
			# add subno in the list of subnos for rollno
			if details[0] in rollno_to_reg_subno.keys():
				rollno_to_reg_subno[details[0]].append(details[3])
			# Otherwise create a new list for subno for the rollno as key
			else:
				rollno_to_reg_subno[details[0]]=[details[3]]
			# if the rollno is already present in the keys
			# add semester info for particular rollno and subno
			if details[0] in rollno_to_seminfo.keys():
				rollno_to_seminfo[details[0]][details[3]]=[details[1],details[2]]
			# if the key is not present then create a dict
			# for the rollono then add {subno: seminfo(list)}
			else:
				rollno_to_seminfo[details[0]]={}
				rollno_to_seminfo[details[0]][details[3]]=[details[1],details[2]]

	# creating dictionary rollno: student info list (name,email,aemail,contact)
	# Read the file 
	f=open(inf,'r')
	rd=csv.DictReader(f)
	for reader1 in rd:
		rollno_to_info[reader1["Roll No"]]=[reader1["Name"],reader1["email"],reader1["aemail"],reader1["contact"]]

	# creating dictionary subno: no of non zero ltp
	# Read the file 
	with open(course_master_file, 'r') as file:
			rd=csv.reader(file)
			for details in rd:
					str=details[2]
					# split l-t-p using separator "-"
					# we get a list [l,t,p]
					ltp = re.split("-", str)
					non_zero=0
					# count number of non-zero terms in list
					for i in range(0,len(ltp)):
						if (ltp[i])!='0':
							non_zero+=1
					# add non-zero count to dict
					subno_to_ltp[details[0]]=non_zero

	temp=1
	# Traversing through all rollno registered
	for rollno in rollno_to_reg_subno.keys():
		# traverse through all subnos for the rollno
		for subno in rollno_to_reg_subno[rollno]:
			# skip it if it is column name 
			if temp :
				temp=0
				continue
			# if nonzero count is 0 then skip
			# i.e., ignore feedback for l-t-p = 0-0-0
			if ((subno in subno_to_ltp) and subno_to_ltp[subno]==0):
				continue
			# if the rollno is not present in feedback file
			# or if the subno is not present in file for the rollno
			# or if number of feedback submitted is less than the non zero
			fk_not_submitted=(rollno not in rollno_to_feedback) or (subno not in rollno_to_feedback[rollno]) or (subno_to_ltp[subno]>rollno_to_feedback[rollno][subno])
			if fk_not_submitted:
				
				# add these rollno with all info in the output file
				output_row=[rollno]
				output_row.extend(rollno_to_seminfo[rollno][subno])
				output_row.append(subno)

				# Check if the rollno info is present in the file
				if rollno in rollno_to_info:
					output_row.extend(rollno_to_info[rollno])
				else:
					output_row.extend(["NA_IN_STUDENTINFO","NA_IN_STUDENTINFO","NA_IN_STUDENTINFO","NA_IN_STUDENTINFO"])

				# Add the row in the output sheet
				ws1.append(output_row)

	# save changes in the file
	wb.save(output_file_name)

feedback_not_submitted()
