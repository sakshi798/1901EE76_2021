# Importing all the neccessary libraries and module

import os
import shutil
from os import listdir
from os.path import isfile, join

# Clear the output screen

os.system("cls")

# Dictionary to store number:name of the web series

diction = {

	1: "Breaking Bad",
	2: "Game of Thrones",
	3: "Lucifer"

}

# Function to rename the web series

def helper(webseries_num,season_padding,episode_padding,source,destn,st):
		webseries_name = diction[webseries_num]

		# Finding the files in the Web series folder
		# File names are stored in the list FileNames

		FileNames = [f for f in listdir(source) if isfile(join(source, f))]
		FileNames_crr = [f for f in listdir(destn) if isfile(join(destn, f))]


		# for using regex 

		import re

		# Traversing list FileNames containing original filenames

		index=-1

		for orig_file_name in FileNames:

			index = index+1
			# Creating new name for the files

			new_name=""
			new_name+=webseries_name+" - "+"Season "

			# Using regex to extract season number
			# and also to get episode number
			# \d+ gives list containing digits until next 
			# non digit is encountered in string

			pattern= re.compile(r'\d+')
			number=re.findall(pattern, orig_file_name)
			
			# First such encounter happens for season
			# and next happens for episode 

			season=int(number[0])
			episode=int(number[1])

			# Calculating number of digits in season

			x = season
			season_digits=0
			while(x):
				x//=10
				season_digits=season_digits+1

			# Calculating number of digits in episode

			y = episode
			episode_digits=0
			while(y):
				y//=10
				episode_digits=episode_digits+1

			# if number padding is greater than number
			# of digits then add zeroes in the beginning 
	
			if season_padding>season_digits:
				new_name+=("0"*(season_padding-season_digits))
			
			# add season number in new name

			new_name+=str(season)
			new_name+=" Episode "

			# if number padding is greater than number
			# of digits then add zeroes in the beginning 

			if episode_padding>episode_digits:
				new_name+=("0"*(episode_padding-episode_digits))

			# add season number in new name

			new_name+=str(episode)

			if(st!="-1"):
				
				new_name+=" - "

				# Split with separator st and " - " 
				# store the required part 
				
				split_1=re.split(st,orig_file_name)
				split_2=split_1[0]
				ep_list=re.split(" - ",split_2)
				episode_name = ep_list[2]

				# Adding episode name extracted using split

				new_name+=episode_name

			# Identifying the file extension from the original list

			length=len(orig_file_name)

			# Add file extension to the file name

			new_name+=orig_file_name[length-4:length]
			add = "./corrected_srt/"+webseries_name+"/"

			# Rename the file in the corrected_srt folder
			
			os.rename(add+FileNames_crr[index],add+new_name)
	

def regex_renamer():

	# Taking input from the user

	print("1. Breaking Bad")
	print("2. Game of Thrones")
	print("3. Lucifer")
	
	# user enters the number of the web series

	webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
	season_padding = int(input("Enter the Season Number Padding: "))
	episode_padding = int(input("Enter the Episode Number Padding: "))

	# Copying the content of wrong_srt (source)
	# to corrected_srt (Destination)

	source = "./wrong_srt"
	destn = "./corrected_srt"

	# Create copy only if it does not exist already
	
	try:
		shutil.copytree(source,destn)
	except:
		pass
	
	webseries_name = diction[webseries_num]

	# Path for the Files in the given web series

	destn=destn+ "/" +webseries_name
	source=source+ "/" +webseries_name

	# If the webseries is Breaking Bad
	# Then rename as follows

	if (webseries_num==1):

		st="-1"
		helper(webseries_num,season_padding,episode_padding,source,destn,st)

	# If the webseries is Game Of Thrones
	# Then rename as follows
	
	if (webseries_num==2):
		st=".WEB"
		helper(webseries_num,season_padding,episode_padding,source,destn,st)

	# If the webseries is Lucifer
	# Then rename as follows
	
	if(webseries_num==3):
		st=".HDTV"
		helper(webseries_num,season_padding,episode_padding,source,destn,st)
		
regex_renamer()