from xlrd import open_workbook
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
def read_narration(audio_path, txt_path):
	#store all the names of narration folders
	narration_folder_names = []
	for narration_folder in os.listdir(audio_path):
		if narration_folder == '.DS_Store':
			continue
		narration_folder_names.append(narration_folder)
	narration_folder_names.sort() #sorted increasingly (starts from prior_1)

	#loop through all the narration files and 
	#put names of the files into a single txt file
	#specified by txt_path
	with open(txt_path, 'w') as txt_file:
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
			
		txt_file.write(result_str)



		
#sys.argv[1] = path of txt file which contains all data source names
def main():
    #there should be only one command line argument 
    #(not counting the program itself)
    if len(sys.argv) != 2:
        print("Wrong number of cmdline args!!")

    excel_path = sys.argv[1]
    content_list = read_spreadsheet(excel_path)

    audio_path = "./narration"
    txt_path = "./word_data.txt"
    read_narration(audio_path, txt_path)

    return

    for content_dic in content_list:
        filename = content_dic['Name']
        write_file(filename, content_dic, narration_txt_path)



if __name__ == '__main__':
    main()