import xlsxwriter
import os
import copy


class Py_Line_Reader():
    def __init__(self):
        self.line = ''
        self.cur_class = ''

    def set_line(self, line):
        self.line = line

    def get_line(self):
        return self.py_file

    def get_class_name(self):
        start = 0
        end = 0
        class_name = ''
        found = False

        for i in range(len(self.line)):
            if self.line[i] != ' ':
                if self.line[i:i + 5] == 'class':
                    start = i + 6
                    found = True

        for i in range(len(self.line)):
            if line[i] == ':' and found is True:
                end = i
                break

        class_name = self.line[start:end]

        return class_name

    def get_function_name(self):
        start = 0
        end = 0
        function_name = ''
        found = False

        for i in range(len(self.line)):
            if self.line[i] != ' ':
                if self.line[i:i + 3] == 'def':
                    start = i + 4
                    found = True

        for i in range(len(self.line)):
            if line[i] == ':' and found is True:
                end = i
                break

        function_name = self.line[start:end]

        return function_name

    def get_functions_len(self):
        return len(self.functions)


class Excel_File_Writer():
    def __init__(self, classes):
        self.classes = classes
        self.workbook = xlsxwriter.Workbook('class_tester_test.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.bold = self.workbook.add_format({'bold': True})

        self.row = 0
        self.col = 0

        self.not_passed = self.workbook.add_format()
        self.not_passed.set_bg_color('red')

        self.title_format = self.workbook.add_format({'bold': True})
        self.title_format.set_bg_color('orange')

        self.data_format = self.workbook.add_format({'text_wrap': True})

        self.col_widths = [17, 30, 15, 20, 20, 5, 40]
        self.headings = ['Class:', 'Unit Being Tested:', 'Test Case:',
                         'Test Data:', 'Expected Result:', 'Result:', 'Comments:']

    def fill_sheet(self):

        for i in range(len(self.headings)):
            self.worksheet.write(self.row, self.col, self.headings[i],
                                 self.title_format)
            self.col += 1

        self.col = 0
        self.row += 1

        for i in range(len(self.classes)):
            self.worksheet.write(self.row, self.col, self.classes[i].get_name())

            self.worksheet.set_column(self.row, self.col,
                                      self.col_widths[self.col])
            self.col = 1

            for x in range(len(self.classes[i].get_functions())):
                self.worksheet.write(self.row, self.col,
                                     self.classes[i].get_functions()[x])

                self.worksheet.set_column(self.row, self.col,
                                          self.col_widths[self.col])
                self.col = 3

                final_input = ''

                for y in range(len(self.classes[i].get_data_input()[x])):
                    final_input += self.classes[i].get_data_input()[x][y] + " =" + "\n"

                self.worksheet.write(self.row, self.col,
                                     final_input, self.data_format)
                self.worksheet.set_column(self.row, self.col,
                                          self.col_widths[self.col])

                self.row += 1
                self.col = 1

            self.row += 1
            self.col = 0

        self.workbook.close()


class Class_Infomation():
    def __init__(self):
        self.name = ''
        self.functions = []
        self.data_input = []

    def set_name(self, name):
        self.name = name

    def get_name(self):
        return self.name

    def set_functions(self, functions):
        self.functions = functions

    def get_functions(self):
        return self.functions

    def set_data_input(self):
        data_input = []
        data = ''
        start = 0
        end = 0

        #print(self.functions)
        for i in range(len(self.functions)):

            for x in range(len(self.functions[i])):
                if self.functions[i][x] == '(':
                    start = x
                    #print(start)

                if self.functions[i][x] == ')':
                    end = x
                    #print(end)

            data = copy.deepcopy(self.functions[i][start + 1:end + 1])
            #print(i)
            print(data)
            count = 0

            start = 0

            for d in range(len(data)):

                #print('woo')

                if data[d] == ',' or data[d] == ')':
                    print(data[d])
                    #print(data[start:d])
                    data_input.append(copy.deepcopy(data[start:d]))
                    #print(data_input[count])
                    start = d + 2
                    count += 1
            for d in range(len(data_input)):
                #print(i)
                #print(data_input[d])
                pass

            self.data_input.append(copy.deepcopy(data_input))
            data_input = []

    def get_data_input(self):
        return self.data_input


if __name__ == "__main__":
    py_file = open(r"C:\Users\duncan\Desktop\AI course work\A Star Search Files\A_star_Euclidean_Heuristic.py")
    py_reader = Py_Line_Reader()
    class_array = []

    class_names = []
    class_functions = []
    class_data_input = []
    cur_name = ''
    cur_function = ''
    cur_data_input = ''
    class_number = 0

    for line in py_file:
        py_reader.set_line(line)

        cur_name = py_reader.get_class_name()
        cur_function = py_reader.get_function_name()
        #class_functions = py_reader.get_function_name()

        if cur_name != '':
            class_names.append(cur_name)
            #print(cur_name)
            class_number += 1
            class_functions.append([])


        if cur_function != '':
            class_functions[class_number - 1].append(cur_function)
            #print(cur_function)

    for i in range(len(class_names)):

        class_array.append(Class_Infomation())
        class_array[i].set_name(class_names[i])
        class_array[i].set_functions(class_functions[i])
        class_array[i].set_data_input()

        print(class_array[i].get_name() + ':')
        print('')
        print('')
        print('')
        print('{')


        for x in range(len(class_array[i].functions)):
            print(class_array[i].functions[x] + ":")
            print('')


            print(class_array[i].data_input[x])



            print('')
            print('')
            print('')
            print('')
        print('}')

    writer = Excel_File_Writer(class_array)
    writer.fill_sheet()
