import xlrd
import os

def find(list,a):
    cnt=[]
    for i in range(0,len(list)):
        if list[i]==a:
            cnt.append(i)
    return cnt

while(1):
    ipt = input("学号：")
    book0 = xlrd.open_workbook(r'C:\Users\76778\Desktop\【信通】开具四六级成绩证明统计.xlsx')
    filepath = r"C:\Users\76778\Desktop\test"
    sheet0 = book0.sheet_by_index(0)
    xh = sheet0.col_values(2)
    idx0=find(xh,ipt)
    need_seq=['语言级别','身份证件号','准考证号','成绩单号','总分']
    print(idx0)
    print(sheet0.row_values(idx0[0]))
    print(need_seq)
    my = os.listdir(filepath)
    for dir in my:
        book = xlrd.open_workbook(filepath+'/'+dir)
        sheet = book.sheet_by_index(0)
        row=sheet.row_values(0)
        idx = find(row,"学号")
        try:
            idx_seq = [find(row,j)[0] for j in need_seq]
        except:
            print(dir)
            break
        xh = sheet.col_values(idx[0])
        idx1 = find(xh,ipt)
        if len(idx1)==0:
            continue
        else:
            for k in idx1:
                print([sheet.cell(k,ix).value for ix in idx_seq])


    
