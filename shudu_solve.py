import xlrd
from prettytable import PrettyTable
from copy import deepcopy


all_numbers = list(i+1 for i in range(9))
shudu_table = []  # 数独表格
shudu_table_by_column = []  # 按每列重排的数独表格，方便按列查找
shudu_table_by_block = []  # 按块重排的数独表格，方便按块查找
guesses = {'level': 0, 'first_guess_cell_detail': {}, 'guessed_num_cnt': 0}   # detail里面是每个级别对应的第一个猜测的单元格索引以及猜测的数字
number_possible_cells_by_block = {}
# 每一次开始猜测都需要一次deep_copy
error = {'status': False, 'position': None, 'description': None}


def concat_str(*args):
    s = ''
    for arg in args:
        s += str(arg)
    return s


def shudu_print(blank_cell_format=None):
    # print(shudu_table)
    print_table = PrettyTable()
    for i in range(9):
        if blank_cell_format == 'detail':
            row = list(concat_str(cell['num'], '(', cell['guess_level'], ', ', cell['guess_order'], ')') if cell[
                'num'] and cell['guess_order'] != 0 else (str(cell['possible_numbers']) if not cell[
                'num'] else cell['num']) for cell in shudu_table[i])
        else:
            row = list(cell['num'] if cell['num'] else '--' for cell in shudu_table[i])
        print_table.add_row(row)
    print(print_table)


def update_column_and_block_table():
    for i in range(9):
        shudu_table_by_column.append([])
        for j in range(9):
            shudu_table_by_column[i].append(shudu_table[j][i])

            # 行和列的下标都是3的倍数的话，说明到了一个新的块
            if i % 3 == 0 and j % 3 == 0:
                shudu_table_by_block.append([])
            shudu_table_by_block[(i//3)*3 + j//3].append(shudu_table[i][j])
    # print(shudu_table_by_column)
    # print(shudu_table_by_block)


def load_orginal_table(excel_path):
    workbook = xlrd.open_workbook(excel_path)
    sheet = workbook.sheet_by_index(0)
    for i in range(0, 9):
        shudu_row = []
        row_values = sheet.row_values(i)
        # print(row_values)
        for j in range(0, 9):
            cell = {'row': i, 'column': j, 'num': None, 'possible_numbers': deepcopy(all_numbers), 'guess_level': 0,
                    'guess_order': 0}
            value = str(row_values[j])
            if value:
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
    # 更新按行和按块的数独表格
    update_column_and_block_table()
    return shudu_table


def exclude_possible_numbers(existing_numbers, possbile_numbers):
    temp_nums = deepcopy(possbile_numbers)
    for num in temp_nums:
        if num in existing_numbers:
            possbile_numbers.remove(num)
    return possbile_numbers


def update_cell(cell):
    if cell['num']:
        # print('cell已经被填充了！')
        return cell
    # 只有一个可能的数字，就直接填充
    if len(cell['possible_numbers']) == 1:
        guesses['guessed_num_cnt'] += 1
        cell['num'] = cell['possible_numbers'][0]
        cell['guess_level'] = guesses['level']
        cell['guess_order'] = guesses['guessed_num_cnt']
    # 没有可能的数字，则表明猜测有误或数独有误
    if len(cell['possible_numbers']) == 0:
        error['status'] = True
        error['position'] = (cell['row'], cell['column'])
        error['description'] = '这个单元格没有找到可能的数字！'
    return cell


def find_cell_possible_nums():
    for table in [shudu_table, shudu_table_by_column, shudu_table_by_block]:
        for i in range(9):
            for j in range(9):
                cell = table[i][j]
                if cell['num']:
                    continue
                else:
                    existing_numbers = list(c['num'] for c in table[i])
                    cell['possible_numbers'] = exclude_possible_numbers(existing_numbers, cell['possible_numbers'])


# 每一行、列、块挨个看每个数字，某个数字只看在一个位置，就填充它
def find_one_possible_place_numbers():
    for table in [shudu_table, shudu_table_by_column, shudu_table_by_block]:
        for i in range(9):
            existing_numbers = list(c['num'] for c in table[i])
            for num in all_numbers:
                if num in existing_numbers:
                    continue
                num_possible_count = 0
                last_num_possible_cell = None
                for cell in table[i]:
                    if num in cell['possible_numbers']:
                        num_possible_count += 1
                        last_num_possible_cell = cell
                if num_possible_count == 0:
                    error['status'] = True
                    error['position'] = (i)
                    error['description'] = str(num) + '在' + str(i) + '没有找到可能的位置！'
                    break
                elif num_possible_count == 1:
                    last_num_possible_cell['possible_numbers'] = [num]
                    update_cell(last_num_possible_cell)
            if error['status']:
                break
        if error['status']:
            break


def find_one_possible_num_cells():
    for i in range(9):
        for j in range(9):
            update_cell(shudu_table[i][j])


def check_shudu_table():
    for table in [shudu_table, shudu_table_by_column, shudu_table_by_block]:
        for i in range(9):
            if sum(c['num'] for c in table[i])


def main():
    path = './数独表格.xlsx'
    load_orginal_table(path)
    print('读取原始表格，数据如下：')
    shudu_print()

    cycle = 1
    guessed_num_cnt = -1
    while guesses['guessed_num_cnt'] > guessed_num_cnt:
        print('\n\n', '第', cycle, '轮猜测')
        guessed_num_cnt = guesses['guessed_num_cnt']
        find_cell_possible_nums()
        find_one_possible_num_cells()
        find_one_possible_place_numbers()


        print('本轮猜出', guesses['guessed_num_cnt']-guessed_num_cnt, '个数字。')
        shudu_print('detail')
        cycle += 1

    print(guesses)


if __name__ == '__main__':
    main()
