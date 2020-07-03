import numpy as np
import math
import xlwt

if __name__ == '__main__':
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('Sheet1')
    with open("life.txt", "r") as f:  # 打开文件
        data = f.read()  # 读取文件
    tmp = data.split("\n")
    database = []
    # 将文本格式保存的人数数据转化为数字格式
    for i in range(len(tmp)):
        if tmp[i]:
            ttltmp = tmp[i].split(" ")
            ttltmp = list(map(eval, ttltmp))
            database.append(ttltmp)
    database = np.array(database)
    qx = []

    # 计算qx死亡率并且存入数组
    for i in range(len(database)):
        qx.append(database[i, 2] / database[i, 1])
    qx = np.array(qx)

    Ax = []
    Ax_2 = []
    var_ax = []
    dingqi = []
    liangquan = []
    v = 1 / (1 + 0.05)

    # 计算各年龄的数据，i为年龄
    for i in range(100):
        sum_1 = 0
        sum_2 = 0

        # 计算精算现值和二阶矩
        for k in range(100 - i):
            t_P_i = database[k + i, 1] / database[i, 1]
            tmp = math.pow(v, k + 1) * t_P_i * qx[i + k]
            sum_1 += tmp  # 累加计算Ax
            sum_2 += tmp * math.pow(v, k + 1)  # 累加计算二阶矩
        var_ax.append(sum_2 - sum_1 ** 2)  # 计算方差
        Ax.append(sum_1 * 1000)
        Ax_2.append(sum_2 * 1000)

        # 计算5、10、15、20、25、30年的定期和两全保险的精算现值
        temp_x = []
        temp_2 = []
        sum_3 = 0
        sum_4 = 0
        # x为时间段
        for x in [5, 10, 15, 20, 25, 30]:
            # 如果时间段加年龄超过一百则判为0
            if x + i >= 100:
                temp_x.append(0)
                temp_2.append(0)
            else:
                for n in range(x):
                    t_P_i = database[n + i, 1] / database[i, 1]
                    tmp = math.pow(v, n + 1) * t_P_i * qx[i + n]
                    sum_3 += tmp  # 累加计算定期
                temp_x.append(sum_3)
                # 计算生存
                shengcun = math.pow(v, x) * database[x + i, 1] / database[i, 1]
                # 求和放入两全
                temp_2.append(sum_3 + shengcun)
        dingqi.append(temp_x)
        liangquan.append(temp_2)

    # 数组格式转换
    var_ax = np.array(var_ax)
    Ax = np.array(Ax)
    Ax_2 = np.array(Ax_2)
    dingqi = np.array(dingqi)
    liangquan = np.array(liangquan)
    qx = qx * 1000

    # 合并数据集
    database = np.c_[database, qx.T]
    database = np.c_[database, Ax]
    database = np.c_[database, Ax_2]
    database = np.c_[database, var_ax]
    database = np.c_[database, dingqi]
    database = np.c_[database, liangquan]

    # 制作表头
    txt = np.array(["x", "lx", "dx", "1000qx", "1000Ax", "1000*2Ax", "Var",
                    "定期5年", "10年", "15年", "20年", "25年", "30年",
                    "两全5年", "10年", "15年", "20年", "25年", "30年"])

    # 写入excel
    for i in range(len(txt)):
        worksheet.write(0, i, txt[i])  # 表头填充
    for i in range(100):
        for j in range(len(txt)):
            temp = database[i, j]
            if temp == 0:
                worksheet.write(i + 1, j, "0")  # “0”的填充
            elif j >= 3:
                worksheet.write(i + 1, j, ("%.4f" % temp))  # 文本格式
            else:
                worksheet.write(i + 1, j, round(temp, 2))  # 数字格式
    workbook.save('data_test.xls')  # 保存
    print(np.average(Ax))  # 计算Ax平均值
