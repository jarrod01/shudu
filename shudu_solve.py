import xlrd
from prettytable import PrettyTable
from copy import deepcopy


all_numbers = list(i+1 for i in range(9))
shudu_table = []
guesses = {'level': 0, 'first_guess_cell_detail': {}, 'order_now': 0}   # detail里面是每个级别对应的第一个猜测的单元格索引以及猜测的数字
# 每一次开始猜测都需要一次deep_copy


def shudu_print(shudu_table=shudu_table):
    # print(shudu_table)
    print_table = PrettyTable()
    for i in range(9):
        row = list(cell['num'] if cell['num'] else '--' for cell in shudu_table[i])
        print_table.add_row(row)
    print(print_table)


def load_orginal_table(excel_path):
    workbook = xlrd.open_workbook(excel_path)
    sheet = workbook.sheet_by_index(0)
    for i in range(0, 9):
        shudu_row = []
        row_values = sheet.row_values(i)
        # print(row_values)
        for j in range(0, 9):
            cell = {'row': i, 'column': j, 'num': None, 'possible_numbers': all_numbers, 'guess_level': 0, 'guess_order': 0}
            if type(row_values[i]) == float:
                try:
                    num = int(float(row_values[j]))
                    cell['num'] = num
                    cell['possible_numbers'] = [num]
                except ValueError:
                    # print('当前行：', i+1, '当前列：', j+1, '数据：', row_values[j], '数独初始化失败， 请检查excel数据！')
                    # return None
                    pass
            shudu_row.append(cell)
        shudu_table.append(shudu_row)
    # print(shudu_table)
    return shudu_table


def find_block_index(cell):
    return (cell['row'] // 3, cell['column'] // 3)


def exclude_possible_numbers(existing_numbers, possbile_numbers):
    temp_nums = deepcopy(possbile_numbers)
    for num in temp_nums:
        if num in existing_numbers:
            possbile_numbers.remove(num)
    return possbile_numbers


def update_cell(cell):
    # 只有一个可能的数字，就直接填充
    if len(cell['possible_numbers']) == 1:
        guesses['order_now'] += 1
        cell['num'] = cell['possible_numbers'][0]
        cell['guess_level'] = guesses['level']
        cell['guess_order'] = guesses['order_now']
    # 没有可能的数字，则表示前面的计算有误，在第一级函数里判断吧
    return cell


# 按行找到某个单元格可能的数字
def find_possible_number_by_row(cell):
    if cell['num']:
        print('cell已经被填充了！')
        return
    row_existing_numbers = list(c['num'] for c in shudu_table[cell['row']])
    # print(row_existing_numbers)
    cell['possible_numbers'] = exclude_possible_numbers(row_existing_numbers, cell['possible_numbers'])
    print(cell)
    return cell


# 按列找到某个单元格可能的数字
def find_possible_number_by_column(cell):
    column_existing_numbers = []
    for i in range(9):
        column_existing_numbers.append(shudu_table[i][cell['column']]['num'])
    # print(column_existing_numbers)
    cell['possible_numbers'] = exclude_possible_numbers(column_existing_numbers, cell['possible_numbers'])
    print(cell)
    return cell


# 按块找到某个单元格可能的数字
def find_possible_number_by_block(cell):
    block_index = find_block_index(cell)
    block_existing_numbers = []
    for i in range(3):
        for j in range(3):
            temp_cell_index = [block_index[0]*3+i, block_index[j]*3+j]
            block_existing_numbers.append(shudu_table[temp_cell_index[0]][temp_cell_index[1]]['num'])
    # print(column_existing_numbers)
    cell['possible_numbers'] = exclude_possible_numbers(block_existing_numbers, cell['possible_numbers'])
    print(cell)
    return cell


def guess_cells():
    guessed_numbers = 0
    for i in range(9):
        for j in range(9):
            pass
    return guessed_numbers



if __name__ == '__main__':
    path = './数独表格.xlsx'
    load_orginal_table(path)
    print('读取原始表格，数据如下：')
    shudu_print()
    cell = shudu_table[0][5]
    find_possible_number_by_row(cell)
    find_possible_number_by_column(cell)
