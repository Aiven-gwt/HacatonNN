import openpyxl as xl




file = open("DATA_2022.txt", 'w')


def take_part(start_pos, end_pos, check=[]):

    DIR = {
        'Номер:': 1,
        'Испол.': 2,
        'Вид:': 3
    }
    str = ''
    res = [sheet['A4'].value]

    counter = 0
    flag = False

    for line in sheet:
        for cell in line:
            if cell.value is not None and cell.value in DIR.keys():
                counter = DIR[cell.value]
                flag = True

            if counter:
                counter -= 1
            elif flag:
                res.append(cell.value)
                flag = False

    #res = np.array(res)
    #res.reshape(len(res) // 2, 2)

    return res


for i in range(1, 6):
    if (len(str(i)) > 1):
        str_ = str(i)
    else:
        str_ = '0' + str(i)
    bock = xl.load_workbook(f"2022/{str_}-2022.xlsx")
    sheet = bock["Print_Povt"]

    res = take_part(3, 60)

    print(res)

    for i in range(0, len(res), 3):
        file.write(str(res[i]) + ' ' + str(res[i+1]) + ' ' + str(res[i+2]) + '\n')

file.close()


dict = {}
file = open("DATA_2022.txt", 'r')

for line in file:
    ITEM = line.split(' ')[0]
    if ITEM not in dict.keys():
        dict[ITEM] = 1
    else:
        dict[ITEM] += 1

file.close()

file = open("freq_2022.txt", 'w')

for key in dict.keys():
    file.write(str(key) + ' ' + str(dict[key]) + '\n')

file.close()