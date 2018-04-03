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



def make_string(word, part, index_start, level, token):
	word_blank = make_blank(word, part, index_start)
	# 5 mjini 2 i mj_ni 1
	return (str(level) + " " + word + " " + str(index_start) + " " 
			+ part + " " + word_blank + " " + str(len(part)) + 
			" " + token + "\n")



#difficulty factors:
#(1). Word Length
#(2). Word frequency (common / rare but >2x)
#(3). #missing letters: 1 < 2 < 3 < 4
# Type:  vowel < consonant; syllable < cluster (ends with vowel < ends with consonant)
#(4). Position of missing letter(s): initial / final / medial
def generate_problems(problem_path, data_path):
	
	syllable_file = open(data_path[0], 'r')
	consonant_file = open(data_path[1], 'r')
	storyword_file = open(data_path[2], 'r')
	problem_file = open(problem_path, 'w')

	result_str = ""
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

	#all data sorted by wordlength first then
	#frequency
	storyword_line_list_list = []
	for line in storyword_file:
		storyword_line_list = line.split(" ")
		storyword_line_list_list.append(storyword_line_list)
		storyword_list.append(storyword_line_list[0])
	storyword_list_reverse = list(reversed(storyword_list))
	storyword_list.sort(key=lambda w: [len(w),storyword_list_reverse.index(w)])
	storyword_file.close()

	#store all syllables of a word in a list
	problem_dic = {}
	for storyword in storyword_list:
		problem_dic[storyword] = []

	for storyword in storyword_list:
		for start_index in range(0, len(storyword)):
			storyword_slice = storyword[start_index:] 

			#loop through syllables_list
			for syllable in syllable_list:
				if storyword_slice.startswith(syllable):
					if len(syllable) == 1:
						problem_dic[storyword].append((syllable, start_index, "vowel"))
					else:
						problem_dic[storyword].append((syllable, start_index, "syllable"))

			#loop through consonant_list:
			for consonant in consonant_list:
				if len(consonant) > 1:
					continue
				if storyword_slice.startswith(consonant):
					problem_dic[storyword].append((consonant, start_index, "consonant"))

	
	result_str = ""
	for storyword in problem_dic:
		level = len(problem_dic[storyword])
	 	for syllable_info in problem_dic[storyword]:
			syllable = syllable_info[0]
			start_index = syllable_info[1]
			token = syllable_info[2]
			result_str += make_string(storyword, syllable, start_index, level, token)
	problem_file.write(result_str)
	problem_file.close()
	return
			


def generate_data(level_list, problem_path):
	with open(problem_path, "r") as problem_file:
		# ['Name', 'Level', '# Missing Letters', '# Missing Syllables', 
		# '# Syllables', 'Position', 'Pronunciation', 'Stimulus Representation', 
		# 'Answer Representation']

		#dictionary for all selected words and their file path (2D list)
		word_list = {}

		#list for word already picked
		picked_word = []

		#list for all problems in the file
		problem_list = []
		for line in problem_file:
			problem_list.append(line[:-1])

		for level in level_list:
			
			json_path = level['Name']
			prob_level = level['Level']
			num_miss = int(float(level['# Missing Letters']))
			syllable_miss = int(float(level['# Missing Syllables']))
			num_syllable = int(float(level['# Syllables']))
			part_pos = level['Position']
			pronu = level['Pronunciation']

			#loop through problem.txt to find an appropriate word
			#word word_blank len(part) token freq_level
			#ni _i n 1 consonant common
			for problem in problem_list:
				# 8 njema 0 n _jema 1 consonant
				
				problem_part = problem.split(" ")
				prob_num_syllable = int(float(problem_part[0]))
				prob_intact = problem_part[1]
				prob_quest = problem_part[3]
				prob_blank = problem_part[4]
				prob_pro = problem_part[6]
				# print("prob_intact %s" % (prob_intact))
				# print("prob_quest %s" % (prob_quest))
				# print("prob_blank %s" % (prob_blank))
				# print("prob_pro %s" % (prob_pro))

				#check repeated word
				if prob_intact in picked_word:
					# print("WRONG: check repeated word")
					continue
				# print("check repeated word done..................")

				#check pronunciation 
				if prob_pro != pronu:
					# print("WRONG: check pronunciation")
					continue
				# print("check pronunciation done..................")

				#check number of missing syllables
				if prob_num_syllable != num_syllable:
					if ((prob_pro == 'vowel') or 
				        (prob_pro == 'consonant')):
						#for vowel/consonant questions choose word with shorter length
						if len(prob_intact) > 5:
							# print("WRONG: check number of missing syllables")
							continue
				# print("check number of missing syllables done..................")

				#check position of the missing part
				if part_pos == "initial":
					if not prob_blank.startswith("_"):
						# print("WRONG: check position of the missing part")
						continue
				elif part_pos == "final":
					if not prob_blank.endswith("_"):
						# print("WRONG: check position of the missing part")
						continue
				else:
					if prob_blank.startswith("_") or prob_blank.endswith("_"):
						# print("WRONG: check position of the missing part")
						continue
				# print("check position of the missing part done..................")

				#only need 20 questions for each
				if len(word_list.setdefault(prob_level, [])) > 20:
					break
				picked_word.append(prob_intact)
				word_list[prob_level].append([prob_intact, prob_blank, prob_quest, json_path])
		
		#print all problems collected
		# for level in word_list:
		# 	print(level)
		# 	for info in word_list[level]:
		# 		print(info)
		# 	print("\n")

		#write into JSON files
		for level in word_list:
			info_list = word_list[level]
			prob_str = ""
			prob_intact = ""
			prob_blank = ""
			prob_quest = ""
			json_path = ""

			# {
			#   "bootFeatures":"FTR_DEMO_MISSING_LTR",
			#   "random": false,
			#   "singleStimulus": true,
			#   "dataSource": [
			prob_str += '{\n\t\"bootFeatures\": \"FTR_DEMO_MISSING_LTR\",'
			prob_str += '\n\t\"random\": false,\n\t\"singleStimulus\": true,\n'
			prob_str += '\t\"dataSource\": [\n\t\t'

			for info in info_list:
				prob_intact = info[0]
				prob_blank = info[1]
				prob_quest = info[2]
				json_path = info[3]
				
				#     {
				#       "stimulus": "s_ba",
				#       "audioStimulus": ["saba"],
				#       "answer": "a"
				#     },
				#     {
				#       "stimulus": "c_a",
				#       "audioStimulus": ["cha"],
				#       "answer": "h"
				#     }
				#   ]
				# }

				prob_str += '{\n\t\t\t'
				prob_str += '\"stimulus\": \"' + prob_blank + '\",\n\t\t\t'
				prob_str += '\"audioStimulus\": [\"' + prob_intact + '\"],\n\t\t\t'
				prob_str += '\"answer\": \"' + prob_quest + '\"\n\t\t},\n\t\t'

			#delete the last comma
			prob_str = prob_str[:-4]
			prob_str += "\n\t]\n}"
			json_path += ".json"
			prob_file = open(json_path, 'w')
			prob_file.write(prob_str)
			prob_file.close()

		return


				
#sys.argv[1] = path of txt file which contains all data source names
def main():
	#there should be only one command line argument 
	#(not counting the program itself)
	if len(sys.argv) != 2:
		print("Wrong number of cmdline args!!")

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
	generate_problems(problem_path, data_path)

	###############generate JSON files###################### 
	excel_path = sys.argv[1]
	level_list = read_spreadsheet(excel_path)

	generate_data(level_list, problem_path)
	return


if __name__ == '__main__':
	main()