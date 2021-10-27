'''
组合是一个基本的数学问题，本程序的目标是输出从n个元素中取m个的所有组合。
例如从[1,2,3]中取出2个数，一共有3中组合：[1,2],[1,3],[2,3]。（组合不考虑顺序，即[1,2]和[2,1]属同一个组合）

本程序的思路（来自网上其他大神）：
（1）创建有n个元素数组，数组元素的值为1表示选中，为0则没选中。
（2）初始化，将数组前m个元素置1，表示第一个组合为前m个数。
（3）从左到右扫描数组元素值的“10”组合，找到第一个“10”组合后将其变为“01”组合，同时将其左边的所有“1”全部移动到数组的最左端。
（4）当某次循环没有找到“10“组合时，说明得到了最后一个组合，循环结束。
例如求5中选3的组合：
1 1 1 0 0 //1,2,3
1 1 0 1 0 //1,2,4
1 0 1 1 0 //1,3,4
0 1 1 1 0 //2,3,4
1 1 0 0 1 //1,2,5
1 0 1 0 1 //1,3,5
0 1 1 0 1 //2,3,5
1 0 0 1 1 //1,4,5
0 1 0 1 1 //2,4,5
0 0 1 1 1 //3,4,5
————————————————
版权声明：本文为CSDN博主「books1958」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/books1958/article/details/46861341
'''

'''
python itertools 实现
'''
import itertools

#从n个数中取出m个数，并将其列表打印出来

t_list = ["1", "2", "3", "4", "5"]
#随机取两个数，允许重复
print("product")
for i in itertools.product(t_list, repeat=2):
    print(i)
#取3个数的排列
print("permutations")
for i in itertools.permutations(t_list, 3):
    print(i)
#取n= 5，分别取m = 1,2,3,4,5的组合
print("combinations")
for x in range(len(t_list)):
    for i in itertools.combinations(t_list, x + 1):
        print(i)
#取m = 2,允许重复的组合
print("combinations_with_replacement")
for i in itertools.combinations_with_replacement(t_list, 2):
    print(i)
