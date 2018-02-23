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
							#'Add/subtract', 'Description', '# questions']
							continue

						col_dic[content_class[col]] = str(sheet.cell(row,col).value)

					if col_dic != {}: content_list.append(col_dic)
				# print(content_list)
				break

	return content_list



def write_file(filename, content_dic):
	with open(filename, 'w') as file_descriptor:

		result_str = "{\n\t"#newline, tab
		fixed_str = '\n\t\t{\"type\": \"Asm_Data\", \"level\": ' #newline, tab, tab

		if "TRUE" in content_dic['Demo']:
			result_str += '\"bootFeatures\": \"FTR_DEMO_'
			if Add in content_dic['Add/subtract']:
				result_str += 'ADD\",\n\n\t'
				operation_str = "+"
			else:
				result_str += 'SUB\",\n\n\t'
				operation_str = "-"
			fixed_str += '\"demo\", '
		else:
			level_str = content_dic['Level']
			fixed_str += '\"' + level_str + '\"'

		result_str += '\"datasource\": [\n\n\t'

		#get dataset
		minValue = content_dic['MinValue']
		maxValue = content_dic['MaxValue']
		offset = content_dic['Offset']
		if offset == "within":
			data_array = range(minValue, maxValue+1)
			random.shuffle(data_array)
		else:
			data_array = range(minValue, maxValue+1, int(offset))

		#get the task field
			task_str = '\"' + content_dic['Description'] + '\"'
		
		#get the image field
			image_str = '\"' + content_dic['Shape'] + '\"' 

		#construct each row in datasouce
		quest_num = int(float(content_dic['# questions']))
		for quest_index in range(0, quest_num):
			result_str += fixed_str

			#task
			result_str +='\"task\": ' + task_str + ', '

			#dataset
			result_str += '\"dataset\": ' + str(data_array) + ', '

			#operation
			result_str += '\"operation\": ' + operation_str + ', '

			#image
			result_str += '\"image\": ' + image_str + '}'

			if quest_index == (quest_num-1): 
				result_str += ']\n\n}'
			result_str += ',\n'




#sys.argv[1] = path of txt file which contains all data source names
def main():
	#there should be only one command line argument 
	#(not counting the program itself)
	if len(sys.argv) != 3:
		print("Wrong number of cmdline args!")

	excel_path = sys.argv[2]
	content_list = read_spreadsheet(excel_path)

	txt_path = sys.argv[1]
	#loop through the file names
	with open(txt_path, 'r') as txt_descriptor:
		for line in txt_descriptor:
			filename = line[:-1] #eliminate '\n'

			#find the attribute dic for this specific file
			for content_dic in content_list:
				radix_index = content_dic['Level'].index('.')
				level_num = content_dic['Level'][:radix_index]
				level_str = "level:"+level_num

				if level_str in filename:
					write_file(filename, content_dic)
				break


if __name__ == '__main__':
    main()
