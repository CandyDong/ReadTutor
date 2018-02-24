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



def num_digit(number):
	if number == 0: return 1
	digit_num = 0
	while number > 0:
		digit_num += 1
		number //= 10
	return digit_num



#check no carry/no borrow
def isValidOperation(operand1, operand2, operation):
	digit_num = num_digit(operand1)
	operand1 = str(operand1)
	operand2 = str(operand2)
	if operation == "+":
		for index in range(digit_num):
			digit1 = int(operand1[index])
			digit2 = int(operand2[index])
			if (digit1 + digit2) > 9:
				return False
		return True
	else:
		for index in range(digit_num):
			digit1 = int(operand1[index])
			digit2 = int(operand2[index])
			if digit1 < digit2:
				return False
		return True



def write_file(filename, content_dic):
	with open(filename, 'w') as file_descriptor:

		result_str = "{\n\t"#newline, tab
		fixed_str = '\n\t\t{\"type\": \"Asm_Data\", \"level\": ' #newline, tab, tab

		if content_dic['Demo']:
			result_str += '\"bootFeatures\": \"FTR_DEMO_'
			if 'Add' in content_dic['Add/subtract']:
				result_str += 'ADD\",\n\n\t'
			else:
				result_str += 'SUB\",\n\n\t'
			fixed_str += '\"demo\", '
		else:
			level_str = content_dic['Level']
			fixed_str += '\"' + level_str + '\"'

		#get operation
		if 'Add' in content_dic['Add/subtract']:
			operation_str = '\"' + "+" '\"'
		else:
			operation_str = '\"' + "-" '\"'

		result_str += '\"datasource\": [\n\t'

		#get the task field
		task_str = '\"' + content_dic['Description'] + '\"'
		
		#get the image field
		image_str = '\"' + content_dic['Shape'] + '\"' 

		#get dataset
		minValue = int(float(content_dic['MinValue']))
		maxValue = int(float(content_dic['MaxValue']))
		if content_dic['Offset'] != "within":
			offset = int(float(int(float(content_dic['Offset']))))

		#increasing/decreasing
		seq_str = content_dic['Increasing/Decreasing/Random']
		if seq_str == "Increasing":
			data_array = range(minValue, maxValue+1, offset)
		elif seq_str == "Decreasing":
			data_array = range(maxValue, minValue-1, -offset)
		else:
			data_array = []
			#minValue is a 3-digit number
			if minValue//100 != 0: 
				offset = 100
			#minValue is a 2-digit number
			elif minValue//10 != 0: 
				offset = 10
			else: 
				offset = 1
			for limit in range(minValue, maxValue+1, offset):
				data_array.append(limit)
			print("random data_array is: ", data_array)

		#construct data set from all data array
		data_set = []
		#number of data set needed
		quest_num = int(float(content_dic['# questions']))
		array_length = len(data_array)
		if "Count" in task_str: 
			if "up" in task_str:
				operand2 = int(float(offset))
			else:
				operand2 = -int(float(offset))
			for data in data_array:
				operand1 = data
				operand3 = operand1 + operand2
			 	data_set.append([str(operand1), str(abs(operand2)), str(operand3)])
			print("ordered data_set is: ", data_set)
		else:
			#find two operands with which addition/subtraction does not involve carry
			#and in different range
			while(True):
				if (len(data_set) >= quest_num):
					print("length of data set is: ", len(data_set))
					break
				i = random.randint(0, array_length-2)
				operand1 = random.randint(data_array[i], data_array[i+1])
				j = random.randint(0, array_length-2)
				operand2 = random.randint(data_array[j], data_array[j+1])
				if 'Add' in content_dic['Add/subtract']:
					if isValidOperation(operand1, operand2, "+"):
						print("valid add")
						operand3 = operand1 + operand2
						data_set.append([str(operand1), str(operand2), str(operand3)])
				else:
					if isValidOperation(operand1, operand2, "-"):
						print("valid sub")
						operand3 = operand1 - operand2
		 				data_set.append([str(operand1), str(operand2), str(operand3)])

			print("random data_set is: ", data_set)

		#construct each row in datasouce
		for quest_index in range(0, quest_num):
			print("quest_index is: ", quest_index)
			result_str += fixed_str

			#task
			result_str +='\"task\": ' + task_str + ', '


			#dataset
			cur_data_set = data_set[quest_index]
			print("cur_data_set is: ", cur_data_set, "\n")
			result_str += '\"dataset\": [' + ','.join(cur_data_set) + '], '

			#operation
			result_str += '\"operation\": ' + operation_str + ', '

			#image
			result_str += '\"image\": ' + image_str + '}'

			if quest_index == (quest_num-1): 
				result_str += ']\n\n}'
				break
			result_str += ',\n'

		file_descriptor.write(result_str)
		file_descriptor.close()




#sys.argv[1] = path of txt file which contains all data source names
def main():
	#there should be only one command line argument 
	#(not counting the program itself)
	if len(sys.argv) != 3:
		print("Wrong number of cmdline args!!")

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
				level_str = "level" + level_num + ":"

				if level_str in filename:
					print(filename)
					write_file(filename, content_dic)
					break

if __name__ == '__main__':
    main()
