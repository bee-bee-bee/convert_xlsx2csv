import codecs
import glob
import os
import re
import time
from asyncio.log import logger

import xlrd
import csv
import sys


# reload(sys)
# sys.setdefaultencoding('UTF-8')


class XLSX_CSV():
    def find_case_id(self, table):
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            for j, column in enumerate(row_value):
                if column == 'Case ID':
                    column_num = j
                    return row_num, column_num

    def align_result_to_step(self, expect_result, step):
        step_len = len(step)
        result_len = len(expect_result)
        result_after_align = [''] * step_len
        index = {}
        sum = 0
        for i, item in enumerate(step):
            if '检查' in item or '检测' in item:
                index[i] = item.count('检查') + item.count('检测')
                sum += index[i]
        if not sum:
            for j, result in enumerate(expect_result):
                result_after_align[-1] += '\n' + result
        elif sum <= len(expect_result):
            for k, v in index.items():
                for j in range(v):
                    result_after_align[k] += expect_result.pop(0) + '\n'
            while expect_result:
                result_after_align[-1] += expect_result.pop(0) + '\n'
        else:
            for j, result in enumerate(expect_result):
                result_after_align[-1] += '\n' + result
        return result_after_align

    def xlsx_to_csv(self, xlsxfile, sheet):
        workbook = xlrd.open_workbook(xlsxfile)
        if re.match(r'\d', sheet):
            table = workbook.sheet_by_index(sheet + 1)
        else:
            table = workbook.sheet_by_name(sheet)
        with codecs.open('output.csv', 'w', encoding='utf-8') as csvfile:
            spamwriter = csv.writer(csvfile, delimiter=';')
            spamwriter.writerow(['Test Case Identifier', 'Component/s', 'Test Repository Path', 'Summary', 'Priority', 'Status', 'Step', 'Expected Result'])
            row, column = self.find_case_id(table)
            while row < table.nrows - 1:
                row += 1
                data = table.row_values(row)
                if not data[column]:
                    if not data[column + 3]:
                        break
                    else:
                        data[column] = str(int(time.time()))
                summary = data[column + 3] + ' > ' + data[column + 4]
                component = data[column + 2]
                priority = data[column + 8] if data[column + 8] else 'P3'
                status = 'valid'
                description = data[column + 5].strip().split("\n")
                step = []
                expect_result = []
                i = 0
                while i < len(description) - 1:
                    if 'Precondition' in description[i] or 'Test Steps' in description[i]:
                        i += 1
                        continue
                    elif re.match(r'^[0-9]', description[i]):
                        temp_str = re.sub(r'^[0-9]*\. *', '', description[i])
                        while i < len(description) - 1:
                            if not re.match(r'^[0-9]', description[i + 1]) and 'Test Steps' not in description[i + 1] and 'Expected Result' not in description[i + 1]:
                                temp_str += description[i + 1]
                                i += 1
                            else:
                                break
                        i += 1
                        step.append(temp_str)
                    elif 'Expected Result' in description[i]:
                        i = i + 1
                        while i < len(description):
                            temp_str = description[i]
                            while i < len(description) - 1:
                                if not re.match(r'^[0-9]', description[i + 1]):
                                    temp_str += '\n' + description[i + 1]
                                    i += 1
                                else:
                                    break
                            i += 1
                            expect_result.append(temp_str)
                if len(step) == 0:
                    logger.info(f"Case {data[column]} has no steps!")
                result = self.align_result_to_step(expect_result, step)
                for i in range(len(step)):
                    if i:
                        summary, component, priority, status = '', '', '', ''
                    spamwriter.writerow([data[column], component, component, summary, priority, status, step[i], result[i]])
            csvfile.close()


if __name__ == '__main__':
    text = XLSX_CSV()
    for xlsx_file in glob.glob(os.curdir + '/*.xlsx'):
        sheet = input("Please input the number or the name of sheet to convert:\n")
        text.xlsx_to_csv(xlsx_file, sheet)
