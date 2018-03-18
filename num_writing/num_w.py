from xlrd import open_workbook
import sys, os

def read_spreadsheet(excel_descriptor):

	for sheet in excel_descriptor.sheets():

		if sheet.name == "Candy - Number Writing": 
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
						#'KC', 'Stimulus', 'Increasing/Decreasing/Random', 'Stimulus Representation']
						continue

					col_dic[content_class[col]] = str(sheet.cell(row,col).value)

				if col_dic != {}: content_list.append(col_dic)
			# print(content_list)
	return content_list



def write_file(filename, content_dic):
	with open(filename, 'w') as file_descriptor:

		result_str = "{\n\t"#newline, tab

		if 'Random' not in content_dic['Increasing/Decreasing/Random']:
			result_str += '\"random\": false,\n\t'
		else:
			result_str += '\"random\": true,\n\t'

		result_str += '\"dataSource\": ['
		for num in range(0,10):
			result_str += '\"'
			result_str += str(num)
			result_str += '\"'
			if num == 9: break
			result_str += ','
		result_str += "]\n"
		result_str += "}"
		file_descriptor.write(result_str)
		file_descriptor.close()



#sys.argv[1] = path of txt file which contains all data source names
def main():
	#there should be only one command line argument 
	#(not counting the program itself)
	if len(sys.argv) != 2:
		print("Wrong number of cmd args!")

	excel_path = sys.argv[1]
	with open_workbook(excel_path,'r') as excel_file:
		content_list = read_spreadsheet(excel_file)

	for content_dic in content_list:
		filename = content_dic['Name']
		write_file(filename, content_dic)

	# txt_path = sys.argv[1]
	# #loop through the file names
	# with open(txt_path, 'r') as txt_file:
	# 	for line in txt_file:
	# 		filename = line[:-1] #eliminate '\n'

	# 		#find the attribute dic for this specific file
	# 		for content_dic in content_list:
	# 			radix_index = content_dic['Level'].index('.')
	# 			level_str = content_dic['Level'][:radix_index]

	# 			if level_str in filename:
	# 				write_file(filename, content_dic)
	# 				break



if __name__ == '__main__':
	main()
