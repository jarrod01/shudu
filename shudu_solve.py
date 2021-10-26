import xlrd
from prettytable import PrettyTable
from copy import deepcopy
from hashlib import md5


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
            row = list(str(cell['possible_numbers']) if not cell[
                'num'] else (concat_str(cell['num'], '(', cell['guess_level'], ', ', cell['guess_order'], ')') if cell[
                'guess_order'] != 0 else  cell['num']) for cell in shudu_table[i])
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


# 根据每个单元格所在的行、列、块，排除已填充的数字，查找单元格可能的数字
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


# 根据同一行或列的其他两个块的可能的数字，排除本块的可能的数组
# 当前块内缺的数字，挨个看其在其他同一行块的可能位置，如果在对应行的其他某块内，这个数字只能在某两行，那么这一行的当前块，该两行可能的数字去掉当前的数字
def exclude_cell_possible_numbers_by_other_block_possible_numbers():
    for i in range(9):
        block = shudu_table_by_block[i]
        existing_numbers = list(c['num'] for c in block)

        for num in all_numbers:
            if num in existing_numbers:
                continue
            # 找到同一行的两个块
            block_column_indexes = [0, 1, 2]
            block_column_indexes.remove(i % 3)
            same_row_other_two_blocks = [shudu_table_by_block[3*(i//3)+block_column_indexes[0]],
                                         shudu_table_by_block[3*(i//3)+block_column_indexes[1]]]
            # 找到当前数字在其他两个块的行
            num_possible_rows = []
            for srb in same_row_other_two_blocks:
                srb_existing_numbers = list(c['num'] for c in srb)
                if num in srb_existing_numbers:
                    num_possible_rows.append(srb[srb_existing_numbers.index(num)]['row'])
                else:
                    for cell in srb:
                        if num in cell['possible_numbers']:
                            num_possible_rows.append(cell['row'])
            num_possible_rows = set(num_possible_rows)
            if len(num_possible_rows) == 2:
                for cell in block:
                    if cell['row'] in num_possible_rows and not cell['num'] and num in cell['possible_numbers']:
                        cell['possible_numbers'].remove(num)

            # 找到同一列的两个块
            block_row_indexes = [0, 1, 2]
            block_row_indexes.remove(i // 3)
            same_column_other_two_blocks = [shudu_table_by_block[3 * (block_row_indexes[0]) + i % 3],
                                            shudu_table_by_block[3 * (block_row_indexes[1]) + i % 3]]
            # 找到当前数字在其他两个块的列
            num_possible_columns = []
            for scb in same_column_other_two_blocks:
                scb_existing_numbers = list(c['num'] for c in scb)
                if num in scb_existing_numbers:
                    num_possible_columns.append(scb[scb_existing_numbers.index(num)]['column'])
                else:
                    for cell in scb:
                        if num in cell['possible_numbers']:
                            num_possible_columns.append(cell['column'])
            num_possible_columns = set(num_possible_columns)
            if len(num_possible_columns) == 2:
                for cell in block:
                    if cell['column'] in num_possible_columns and not cell['num'] and num in cell['possible_numbers']:
                        cell['possible_numbers'].remove(num)


# 同一个块或行、列内，某几个数字可能在的位置是重复的，那么这几个格也只能是这几个数字
def exclude_cell_possible_numbers_by_number_possible_cells():
    pass


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


# 查找只有一个可能的数字的单元格，并填充
def find_one_possible_num_cells():
    for i in range(9):
        for j in range(9):
            update_cell(shudu_table[i][j])


# 每行每列每块检查求和是否等于45
def check_shudu_table():
    for table in [shudu_table, shudu_table_by_column, shudu_table_by_block]:
        for i in range(9):
            existing_numbers = list(c['num'] for c in table[i])
            if None in existing_numbers:
                return False
            if sum(existing_numbers) != 45:
                return False
    return True


def main():
    path = './数独表格.xlsx'
    load_orginal_table(path)
    print('读取原始表格，数据如下：')
    shudu_print()

    cycle = 1
    cycle_before_shudu_md5 = ''
    cycle_after_shudu_md5 = None
    while cycle_before_shudu_md5 != cycle_after_shudu_md5:
        print('\n\n', '第', cycle, '轮猜测')
        guessed_num_cnt = guesses['guessed_num_cnt']
        cycle_before_shudu_md5 = md5(str(shudu_table).encode('utf-8')).hexdigest()

        find_cell_possible_nums()
        find_one_possible_num_cells()
        exclude_cell_possible_numbers_by_other_block_possible_numbers()
        find_one_possible_place_numbers()


        cycle_after_shudu_md5 = md5(str(shudu_table).encode('utf-8')).hexdigest()

        if error['status']:
            print('本轮猜测发现错误：', error['description'])
            break
        else:
            print('本轮猜出', guesses['guessed_num_cnt']-guessed_num_cnt, '个数字。')
            shudu_print('detail')
            cycle += 1

    print(guesses)
    checked = check_shudu_table()
    if checked:
        print('数独求解完成')


if __name__ == '__main__':
    main()
