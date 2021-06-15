import math
import  random
# 冒泡排序
alist=[5,8,6,9,3,4,1,2,7,0]

def maopao_sort():
    for a in range(len(alist)-1):
        for b in range(len(alist)-1):
            if alist[b]>alist[b+1]:
                alist[b],alist[b+1]=alist[b+1],alist[b]
    print(alist)

# 选择排序
def select_sort(alist):
    """选择排序
    外层只循环数组的长度次，内存每次循环找出最大值或最小值，然后通过外层循环放在最前面，内循环逐渐减少
     因为每次都会排序好一个最大值或最小值，之后只循环剩下的，一次类推"""
    # 外层循环控制循环次数
    n=len(alist)
    for j in range(0,  n- 1):
        # 假设找到的最小元素下标为j
        min_index = j
        # 寻找最小元素的过程
        for i in range(j + 1, n):
            # 假设最小下标的值，大于循环中一个元素，那么就改变最小值的下标,找出最小值下标
            if alist[min_index] > alist[i]:
                min_index = i
        # 为了容错处理，因为循环一开始就假设把最小值的下标j赋值给变量min_index
        if j != min_index:
            # 在不停的循环中，不停的交换两个不一样大小的值
            alist[min_index], alist[j] = alist[j], alist[min_index]
    print(alist)

alist=[5,8,6,9,3,4,1,2,7,0]


for a in range(len(alist)-1):
    mind=a
    for b in range(a+1,len(alist)):
        if alist[mind]>alist[b]:
            mind=b
    if a!=mind:
         alist[mind],alist[a]=alist[a],alist[mind]
print(alist)


# 插入排序
""" 插入排序
    外层只循环数组的长度次，然后内循环先排列好【0-1】第二循环排列好【0-2】第三次循环排列好【0-3】的以此类推"""
def charu_sort():
    count = len(alist)
    for i in range(0, count):
        key = alist[i]
        j = i
        while j >= 0:
            if alist[j] > key:
                alist[j + 1] = alist[j]
                alist[j] = key
            j -= 1
    print(alist)

# 快速排序
def quick_sort(lists, left, right):
    # 快速排序
    if left >= right:
        return lists
    key = lists[left]
    low = left
    high = right
    while left < right:
        while left < right and lists[right] >= key:
            right -= 1
        lists[left] = lists[right]
        while left < right and lists[left] <= key:
            left += 1
        lists[right] = lists[left]
    lists[right] = key
    quick_sort(lists, low, left - 1)
    quick_sort(lists, left + 1, high)
    return lists

# 堆排序是借助堆来实现的选择排序s
# 希尔排序
def ShellInsetSort(array, len_array, dk):  # 直接插入排序
      for i in range(dk, len_array):  # 从下标为dk的数进行插入排序
          position = i
          current_val = array[position]  # 要插入的数

          index = i
          j = int(index / dk)  # index与dk的商

          index = index - j * dk

         # while True:  # 找到第一个的下标，在增量为dk中，第一个的下标index必然 0<=index<dk
         #     index = index - dk
         #     if 0<=index and index <dk:
         #         break
         # position>index,要插入的数的下标必须得大于第一个下标
          while position > index and current_val < array[position-dk]:
              array[position] = array[position-dk]  # 往后移动
              position = position-dk
          else:
              array[position] = current_val

def ShellSort(array, len_array):  # 希尔排序
     dk = int(len_array/2)  # 增量
     while(dk >= 1):
         ShellInsetSort(array, len_array, dk)
         print(">>:",array)
         dk = int(dk/2)

if __name__ == "__main__":
     array = [49, 38, 65, 97, 76, 13, 27, 49, 55, 4]
     #print(">:", array)
     #ShellSort(array, len(array))
     select_sort(array)










