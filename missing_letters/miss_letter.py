from xlrd import open_workbook
import xlwt
import sys, os, random

def read_spreadsheet(excel_path):
    with open_workbook(excel_path,'r') as excel_descriptor:
        for sheet in excel_descriptor.sheets():

            if sheet.name == "Candy - Add Subtract": 
                row_num = sheet.nrows
                col_num = sheet.ncols
                content_list = []
                content_class = []

                for row in range(row_num):
                    #dic for a specific column(a number writing file)
                    col_dic = {} 

                    for col in range(col_num):

                        if row == 0: 
                            content_class.append(str(sheet.cell(row,col).value))
                            #['Level', 'MinValue', 'MaxValue', 'Offset', 'Domain', 
                            #'KC', 'Increasing/Decreasing/Random', 'Shape', 'Demo', 
                            #'Add/subtract', 'Description', 'Name', '# questions']
                            continue

                        col_dic[content_class[col]] = str(sheet.cell(row,col).value)

                    if col_dic != {}: content_list.append(col_dic)
                # print(content_list)
                break

    return content_list



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
			print(folder_path)
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
				result_str += filename
				result_str += "\n"
			
		word_data_file.write(result_str)



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
				syllable_data = ""
				for row in range(row_num):
					if row < 2: 
						continue
					syllable_value = sheet.cell(row,7).value
					if syllable_value == "":
						continue
					syllable_data += str(syllable_value) #197 swahili word starts from column 7
					syllable_data += "\n"
				syllable_file = open(data_path[0], 'w')
				syllable_file.write(syllable_data)
				syllable_file.close()

			if "consonant" in sheet.name:
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
				story_data = ""
				for row in range(row_num):
					if row < 1 :
						continue
					story_word_value = sheet.cell(row,0).value
					story_freq_value = int(sheet.cell(row,2).value)
					story_data += str(story_word_value)
					story_data += " "
					story_data += str(story_freq_value)
					story_data += " "

					if story_freq_value > 4:
						story_data += "common"
					else:
						story_data += "rare"
					story_data += "\n"
				story_file = open(data_path[2], 'w')
				story_file.write(story_data)
				story_file.close()




def make_blank(word, content):
	index_start = word.find(content)
	index_end = index_start + len(content)
	word_blank = word[:index_start] + "_" * len(content) + word[index_end:] 
	return word_blank




def make_string(word, content, level):
	word_blank = make_blank(word, letter)
	return "level" + str(level) + "   " + word + "   " + word_blank + "\n"


#difficulty factors:
#(1). Word Length
#(2). #missing letters: 1 < 2 < 3 < 4
# Type:  vowel < consonant; syllable < cluster (ends with vowel < ends with consonant)
#(3). Position of missing letter(s): initial / final / medial
#(4). Word frequency (common / rare but >2x)
def generate_problems(word_path, problem_path, data_path):
	#loop through all the words in word_data_path
	with open(word_path, "r") as word_file:
		syllable_file = open(data_path[0], 'r')
		consonant_file = open(data_path[1], 'r')
		storyword_file = open(data_path[2], 'r')

		result_str = ""
		word_list = []
		syllable_list = []
		consonant_list = []
		storyword_list = []
		level = 0

		#all data sorted by length
		for line in word_file:
			word_list.append(line[:-1])
		word_list.sort(key=len)

		#all data sorted by length
		for line in syllable_file:
			syllable_list.append(line[:-1])
		syllable_list.sort(key=len)

		#all data sorted by length
		for line in consonant_file:
			consonant_list.append(line[:-1])
		consonant_list.sort(key=len)

		#all data sorted by frequency
		for line in storyword_file:
			storyword_line_list = line.split(" ")
			storyword_list.append(storyword_line_list[0])
		

		for word in word_list:
			print(word)
			#number of missing letters
			for missing_num in range(1, len(word)+1):
				for start_index in range(0, len(word)):
					if (start_index + missing_num) > len(word):
						break
					phrase = word[start_index : (start_index+missing_num)]
					print(phrase)
		return
					


			# for letter in word:
			# 	for consonant in consonant_file:
			# 		if len(consonant) > 1:
			# 			break
			# 		if consonant in word:
			# 			level += 1
			# 			result_str += make_string(word, letter, level)

			# 	if letter in vowels:
			# 		level += 1
			# 		result_str += make_string(word, letter, level)

		


		return
					

				

				
		
#sys.argv[1] = path of txt file which contains all data source names
def main():
    #there should be only one command line argument 
    #(not counting the program itself)
    if len(sys.argv) != 2:
        print("Wrong number of cmdline args!!")

    excel_path = sys.argv[1]
    content_list = read_spreadsheet(excel_path)

    audio_path = "./narration"
    word_data_path = ".word_data.txt"
    read_narration(audio_path, word_data_path)

    info_path = "./7181_swahili_story_words_Swahili_syllable_and_consonant_statistics.xlsx"
    syllable_path = "./syllable.txt" 
    consonant_path = "./consonant.txt"
    storyword_path = "./storyword.txt"
    data_path = [syllable_path, consonant_path, storyword_path]
    write_info_data(info_path, data_path)

    problem_path = "./problems.xlsx"
    generate_problems(word_data_path, problem_path, data_path)

if __name__ == '__main__':
    main()