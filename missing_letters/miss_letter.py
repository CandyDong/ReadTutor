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
def read_narration(narration_folder_path, narration_txt_path):
	with open(narration_txt_path, 'w') as narration_txt:
		result_str = ""
		for filename in os.listdir(narration_folder_path):
			result_str += filename
			result_str += "\n"
		narration_txt.write(result_str)


		
#sys.argv[1] = path of txt file which contains all data source names
def main():
    #there should be only one command line argument 
    #(not counting the program itself)
    if len(sys.argv) != 2:
        print("Wrong number of cmdline args!!")

    excel_path = sys.argv[1]
    content_list = read_spreadsheet(excel_path)

    narration_folder_path = "./Initial_syllable_narration"
    narration_txt_path = "./Initial_syllable_narration.txt"
    read_narration(narration_folder_path, narration_txt_path)

    return



    for content_dic in content_list:
        filename = content_dic['Name']
        write_file(filename, content_dic)



if __name__ == '__main__':
    main()