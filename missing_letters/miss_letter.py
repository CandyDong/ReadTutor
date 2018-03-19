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



#sys.argv[1] = path of txt file which contains all data source names
def main():
    #there should be only one command line argument 
    #(not counting the program itself)
    if len(sys.argv) != 2:
        print("Wrong number of cmdline args!!")

    excel_path = sys.argv[1]
    content_list = read_spreadsheet(excel_path)

    for content_dic in content_list:
        filename = content_dic['Name']
        write_file(filename, content_dic)



if __name__ == '__main__':
    main()