from random import randint
import xlwt, xlrd

def shudu_generate():
    num_list = [1, 2, 3, 4, 5, 6, 7, 8, 9]
    shudu_num = []
    success = True
    for i in range(0, 9):
        shudu_num.append([])
        for j in range(0, 9):
            shudu_num[i].append(0)
    for i in range(0, 9):
        for j in range(0, 9):
            i_tmp = i
            j_tmp = j
            tmp_num_list = num_list.copy()
            while i_tmp != 0:
                i_tmp -= 1
                try:
                    tmp_num_list.remove(shudu_num[i_tmp][j])
                except ValueError:
                    continue
            while j_tmp != 0:
                j_tmp -= 1
                try:
                    tmp_num_list.remove(shudu_num[i][j_tmp])
                except ValueError:
                    continue
            i_tmp = i
            while i_tmp % 3 != 0:
                i_tmp -= 1
                j_tmp = j
                while j_tmp % 3 != 0:
                    j_tmp -= 1
                    try:
                        tmp_num_list.remove(shudu_num[i_tmp][j_tmp])
                    except ValueError:
                        continue
                j_tmp = j
                while (j_tmp+1) % 3 != 0 and j < 9:
                    j_tmp += 1
                    try:
                        tmp_num_list.remove(shudu_num[i_tmp][j_tmp])
                    except ValueError:
                        continue
            n = len(tmp_num_list)
            if n == 0:
                success = False
                return {'success': success, 'shudu': []}
            if n == 1:
                shudu_num[i][j] = tmp_num_list[0]
            else:
                t = randint(0, n-1)
                shudu_num[i][j] = tmp_num_list[t]
    if success:
        return {'success': success, 'shudu': shudu_num}

def print_result(shudu):
    for i in range(0, 9):
        line = ''
        for j in range(0, 9):
            if j % 3 == 2:
                line += str(shudu[i][j]) + '| '
            else:
                line += str(shudu[i][j]) + '|'
        print(line)
        if i % 3 == 2:
            print('--------------------')

def get_result():
    result = shudu_generate()
    t = 1
    while not result['success']:
        result = shudu_generate()
        if not result['success']:
            pass
            # print('第' + str(t) + '次尝试失败！')
        else:
            print('第' + str(t) + '次尝试成功！')
            shudu = result['shudu']
        t += 1
    return shudu

def put_result_to_excel(shudu, result_file):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('shudu')
    border_left = xlwt.Borders()
    border_left.left = xlwt.Borders.THIN
    style_left = xlwt.XFStyle()
    style_left.borders = border_left
    border_top = xlwt.Borders()
    border_top.top = xlwt.Borders.THIN
    style_top = xlwt.XFStyle()
    style_top.borders = border_top
    border_left_top = xlwt.Borders()
    border_left_top.left = xlwt.Borders.THIN
    border_left_top.top = xlwt.Borders.THIN
    style_left_top = xlwt.XFStyle()
    style_left_top.borders = border_left_top
    for i in range(0, 9):
        for j in range(0, 9):
            num = shudu[i][j]
            if i%3 == 0:
                if j%3 == 0:
                    sheet.write(i, j, num, style_left_top)
                else:
                    sheet.write(i, j, num, style_top)
            else:
                if j%3 == 0:
                    sheet.write(i, j, num, style_left)
                else:
                    sheet.write(i, j, num)
    for i in range(0, 9):
        sheet.write(i, 9, '', style_left)
        sheet.write(9, i, '', style_top)
    caculate = workbook.add_sheet('result')
    caculate.write(0, 0, '每行')
    caculate.write(0, 1, '每列')
    caculate.write(0, 2, '每块')
    columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    for i in range(0, 9):
        caculate.write(i+1, 0, xlwt.Formula('IF(SUM(shudu!A%s:I%s)=45,1,0)'% (str(i+1), str(i+1))))
        caculate.write(i+1, 1, xlwt.Formula('IF(SUM(shudu!%s1:%s9)=45,1,0)'% (columns[i], columns[i])))
        if i%3 == 0:
            caculate.write(i+1, 2, xlwt.Formula('IF(SUM(shudu!%s%s:%s%s)=45,1,0)'% (columns[0], str(i+1), columns[2], str(i+3))))
            caculate.write(i+2, 2, xlwt.Formula('IF(SUM(shudu!%s%s:%s%s)=45,1,0)' % (columns[3], str(i + 1), columns[5],str(i + 3))))
            caculate.write(i+3, 2, xlwt.Formula('IF(SUM(shudu!%s%s:%s%s)=45,1,0)' % (columns[6], str(i + 1), columns[8], str(i + 3))))
    sheet.write(11, 0, '结果：')
    sheet.write(11, 1, xlwt.Formula('IF(SUM(result!A2:C10)=27, "成功", "继续努力")'))
    workbook.save(result_file + '.xls')


def play_shudu():
    shudu = get_result()
    hard = 1
    hard = int(input('请选择难度(1-4)：'))
    level = [6, 5, 4, 3]
    hard = level[hard-1]
    shudu_out = shudu
    put_result_to_excel(shudu, 'shudu_answer')
    for i in range(0, 9):
        for j in range(0, 9):
            r = randint(1, 10)
            if r > hard:
                shudu_out[i][j] = ''
    put_result_to_excel(shudu_out, 'shudu_test')
    print('请打开shudu.xls作答，填完后会自动显示结果')


    '''
    line = []
    for i in range(0, 9):
        line[i] = ('请输入第' + str(i) + '行')
        line[i] = line[i].split(' ')
        while len(line[i]) != 9:
            line[i] = ('请重新输入第' + str(i) + '行')
            line[i] = line[i].split(' ')
    print_result(shudu)
    '''

if __name__ == '__main__':
    play_shudu()
