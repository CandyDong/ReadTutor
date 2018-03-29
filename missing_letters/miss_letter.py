from xlrd import open_workbook
import xlwt
import sys, os, random

#temporaty placeholder
def read_spreadsheet(excel_path):
	with open_workbook(excel_path,'r') as excel_descriptor:
		for sheet in excel_descriptor.sheets():
			row_num = sheet.nrows
			col_num = sheet.ncols
			level_list = []
			part_class = []

			for row in range(row_num):
				#dic for a specific column(a number writing file)
				col_dic = {} 

				for col in range(col_num):

					if row == 0: 
						part_class.append(str(sheet.cell(row,col).value))
						#['Name', 'Level', '# Missing Letters', 'Word Length', 
						#'Position', 'Pronunciation', 'Word Frequency', 
						#'Stimulus Representation', 'Answer Representation']
						continue

					col_dic[part_class[col]] = str(sheet.cell(row,col).value)

				if col_dic != {}: level_list.append(col_dic)
			
			break

	return level_list



#read file names in narration path and put them into a txt file 
#which is under the same direction as this python script
def read_narration(audio_path, word_data_path):
	#store all the names of narration folders
	narration_folder_names = []

	for narration_folder in os.listdir(audio_path):
		if narration_folder == '.DS_Store':
			continue
		narration_folder_names.append(narration_folder)
	narration_folder_names.sort() #sorted increasingly (starts from prior_1)

	#loop through all the narration files and 
	#put names of the files into a single txt file
	#specified by word_data_path
	with open(word_data_path, 'w') as word_data_file:
		result_str = ""
		for folder in narration_folder_names:
			folder_path = audio_path + "/" + folder

			print("scanning recorded words in %s" % folder_path)

			for filename in os.listdir(folder_path):

				#eliminate files that are not mp3 or wav
				if ((not filename.endswith(".mp3")) and
				   (not filename.endswith(".wav"))):
					continue

				#delete suffix like ".mp3" and ".wav"
				suffix_index = filename.find(".")
				filename = filename[:suffix_index]

				#delete "copy of" if it's in the filename
				if filename.startswith("Copy of"):
					filename_list = filename.split(" ")
					filename = filename_list[2]

				#delete (1) if it's in the filename
				if filename.endswith("(1)"):
					num_index = filename.find("(1)")
					filename = filename[:num_index]

				#eliminate phrase
				if len(filename.split(" ")) > 1:
					continue

				#convert all the letters to lowercase
				filename = filename.lower()

				#eliminate words that have already been included
				
				if filename[:suffix_index] in result_str:
					continue

				#print(filename)

				result_str += filename
				result_str += "\n"
			
		word_data_file.write(result_str)


#loop through story word
#calculate total appearance of words that contain the consonant
#filter out those have frequence less than or equal to 5
def filter_low_freq(part_path, new_part_path, story_path):
	new_str = ""
	with open(new_part_path, 'w') as new_part_file:
		with open(part_path, 'r') as part_file:

			print("filter low frequency parts in %s" % part_path)

			for part_line in part_file:
				part = part_line[:-1] #"\n"

				freq_sum = 0
				with open(story_path, 'r') as story_file:
					for story_line in story_file:
						storyword_line_list = story_line.split(" ")
						word = storyword_line_list[0]
						freq = int(storyword_line_list[1])

						if part in word:
							freq_sum += freq
							if freq_sum > 5:
								#write this word into the new file
								break

					if freq_sum > 5:
						new_str += str(part_line)

		new_part_file.write(new_str)



def write_info_data(info_path, data_path):
	with open_workbook(info_path, 'r') as info_file:
		#find the spreadsheet for
		#1. syllable data
		#2. consonant data
		#3. story word info

		for sheet in info_file.sheets():
			row_num = sheet.nrows
			col_num = sheet.ncols

			if "syllable" in sheet.name:

				print("write data in %s" % data_path[0])

				syllable_data = ""
				for row in range(row_num):
					if row < 2: 
						continue
					syllable_value = sheet.cell(row,7).value #197 swahili word starts from column 7
					if syllable_value == "":
						continue
					syllable_data += str(syllable_value) 
					syllable_data += "\n"
				syllable_file = open(data_path[0], 'w')
				syllable_file.write(syllable_data)
				syllable_file.close()

			if "consonant" in sheet.name:

				print("write data in %s" % data_path[1])

				consonant_data = ""
				for row in range(row_num):
					if row < 1:
						continue
					consonant_value = sheet.cell(row,0).value
					consonant_data += str(consonant_value)
					consonant_data += "\n"
				consonant_file = open(data_path[1], 'w')
				consonant_file.write(consonant_data)
				consonant_file.close()

			if "story words" in sheet.name:

				print("write data in %s" % data_path[2])

				story_data = ""
				for row in range(row_num):
					if row < 1 :
						continue

					story_word_value = sheet.cell(row,0).value
					story_freq_value = int(sheet.cell(row,2).value)

					#do not include story words that have frequence <= 3 
					#(might be English words)
					if story_freq_value <= 3:
						continue
					
					#Swahili words end with a vowel virtually always.
					#filter out those end with a consonant
					#eg. of, oh.
					end_letter = story_word_value[-1];
					if end_letter not in ['a','e','i','o','u']:
						continue 

					story_data += str(story_word_value)
					story_data += " "
					story_data += str(story_freq_value)
					story_data += " "

					#temprary "common/rare" cutoff 20 can be changed later
					if story_freq_value > 20:
						story_data += "common"
					else:
						story_data += "rare"
					story_data += "\n"
				story_file = open(data_path[2], 'w')
				story_file.write(story_data)
				story_file.close()



def make_blank(word, part, index_start):
	index_end = index_start + len(part)
	word_blank = word[:index_start] + "_" * len(part) + word[index_end:] 
	return word_blank



def make_string(word, part, index_start, level, token, freq_level):
	word_blank = make_blank(word, part, index_start)
	return word + " " + word_blank + " " + str(len(part)) + " " + token + " " + freq_level + "\n"



#difficulty factors:
#(1). Word Length
#(2). Word frequency (common / rare but >2x)
#(3). #missing letters: 1 < 2 < 3 < 4
# Type:  vowel < consonant; syllable < cluster (ends with vowel < ends with consonant)
#(4). Position of missing letter(s): initial / final / medial
def generate_problems(word_path, problem_path, data_path):
	word_file = open(word_path, 'r')
	syllable_file = open(data_path[0], 'r')
	consonant_file = open(data_path[1], 'r')
	storyword_file = open(data_path[2], 'r')
	problem_file = open(problem_path, 'w')

	result_str = ""
	word_list = []
	syllable_list = []
	consonant_list = []
	storyword_list = []
	level = 0

	#all data sorted by length
	for line in syllable_file:
		syllable_list.append(line[:-1])
	syllable_file.close()
	syllable_list.sort(key=len)

	#all data sorted by length
	for line in consonant_file:
		consonant_list.append(line[:-1])
	consonant_file.close()
	consonant_list.sort(key=len)

	#all data sorted by frequency originally 
	storyword_line_list_list = []
	for line in storyword_file:
		storyword_line_list = line.split(" ")
		storyword_line_list_list.append(storyword_line_list)
		storyword_list.append(storyword_line_list[0])
	storyword_file.close()

	#all data sorted by length and then frequency as specified in storyword
	for line in word_file:
		word = line[:-1]

		#eliminate those that are not in the storyword list
		if not (word in storyword_list):
			continue

		word_list.append(word)
	word_file.close()
	storyword_list_reverse = list(reversed(storyword_list))
	word_list.sort(key=lambda w: [len(w),storyword_list_reverse.index(w)])
	
	for word in word_list:
		#number of missing letters
		for missing_num in range(1, len(word)+1):
			for start_index in range(0, len(word)):
				if (start_index + missing_num) > len(word):
					break
				if missing_num == len(word):
					break
				part = word[start_index : (start_index+missing_num)]

				#word_freq
				#eliminate '\n'
				freq_level = storyword_line_list_list[storyword_list.index(word)][2][:-1] 

				#loop through consonant_list and syllables_list
				#for part of same length consonants have higher difficulty levels
				for syllable in syllable_list:
					if syllable == part:
						level += 1
						result_str += make_string(word, part, start_index, level, 'vowel', freq_level)
				for consonant in consonant_list:
					if consonant == part:
						level += 1
						result_str += make_string(word, part, start_index, level, 'consonant', freq_level)

	problem_file.write(result_str)
	problem_file.close()
	return
			


def generate_data(level_list, problem_path):
	with open(problem_path, "r") as problem_file:
		#['Name', 'Level', '# Missing Letters', 'Word Length', 
		#'Position', 'Pronunciation', 'Word Frequency', 
		#'Stimulus Representation', 'Answer Representation']

		#list for all selected words and their file path (2D list)
		word_list = []
		#list for all problems in the file
		problem_list = []
		
		for line in problem_file:
				problem_list.append(line[:-1])

		for level in level_list:
			print("\n")
			
			json_path = level['Name']
			num_miss = int(float(level['# Missing Letters']))
			word_len = int(float(level['Word Length']))
			part_pos = level['Position']
			pronu = level['Pronunciation']
			freq = level['Word Frequency']

			print("Finding problem for level %s" % level['Level'])
			print(level)

			#loop through problem.txt to find an appropriate word
			#word word_blank len(part) token freq_level
			for problem in problem_list:
				print("Checking problem %s" % problem)

				problem_part = problem.split(" ")
				#check word length
				prob_intact = problem_part[0]
				if len(prob_intact) != word_len:
					if len(prob_intact) > word_len:
						print("no appropriate word exists!!")
						break
					print("invalid word length")
					continue
				print("Word length checked")

				#check number of missing letters
				prob_num_miss = problem_part[2]
				if int(float(prob_num_miss)) != num_miss:
					print("invalid number of missing letters")
					continue
				print("number of missing letters checked")

				#check position of the missing part
				prob_part = problem_part[1]
				if part_pos == "initial":
					if not prob_part.startswith("_"):
						print("invalid part position")
						continue
				elif part_pos == "final":
					if not prob_part.endswith("_"):
						print("invalid part position")
						continue
				else:
					if prob_part.startswith("_") or prob_part.endswith("_"):
						print("invalid part position")
						continue
				print("part position checked")

				#check pronunciation 
				prob_pro = problem_part[3]
				if pronu != prob_pro:
					print("invalid pronunciation")
					continue
				print("pronunciation checked")

				#check word frequency
				prob_freq = problem_part[4]
				if freq != prob_freq:
					print("invalid word frequency ")
					continue
				print("word frequency checked")

				#check if word is repeated
				for info_pair in word_list:
					if prob_intact in info_pair:
						print("repeated word")
						continue

				print("word passed!!")
				word_list.append([prob_intact, json_path])
				break


		print(word_list)



				





				
#sys.argv[1] = path of txt file which contains all data source names
def main():
	#there should be only one command line argument 
	#(not counting the program itself)
	if len(sys.argv) != 2:
		print("Wrong number of cmdline args!!")

	audio_path = "./narration"
	word_data_path = "./word_data.txt"
	read_narration(audio_path, word_data_path)

	info_path = "./7181_swahili_story_words_Swahili_syllable_and_consonant_statistics.xlsx"
	syllable_path = "./syllable.txt" 
	consonant_path = "./consonant.txt"
	storyword_path = "./storyword.txt"
	filtered_consonant_path = "./consonant_filtered.txt"
	filtered_syllable_path = "./syllable_filtered.txt"
	data_path = [syllable_path, consonant_path, storyword_path]
	write_info_data(info_path, data_path)
	filter_low_freq(consonant_path, filtered_consonant_path, storyword_path)
	filter_low_freq(syllable_path, filtered_syllable_path, storyword_path)
	data_path = [filtered_syllable_path, filtered_consonant_path, storyword_path]

	problem_path = "./problems.txt"
	generate_problems(word_data_path, problem_path, data_path)

	###############generate JSON files###################### 
	excel_path = sys.argv[1]
	level_list = read_spreadsheet(excel_path) 
	generate_data(level_list, problem_path)



if __name__ == '__main__':
	main()