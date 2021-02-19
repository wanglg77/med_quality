from define import *

def recstastic():
    year = input('请输入年份: ')
    month = input('请输入月份: ')
    title = year+'年'+month+'月'+'病历质量奖罚汇总表'
    file1 = input('请输入每月病历质控原始数据表路径和文件名: ')
    file2 = input('请输入每月病历归档原始数据表路径和文件名: ')
    wb = load_workbook(file1,read_only=True)
    ws = wb.active
    #病历评级数据
    dept_op = {} #科室门诊病历字典
    dept_opa = {}
    dept_opb = {}
    dept_hj = {} #科室环节病历字典
    dept_hja = {}
    dept_hjb = {}
    dept_zm = {} #科室终末病历字典
    dept_zma = {}
    dept_zmb = {}
    dept_yx = {} #科室优秀病历字典
    staffs = {}

    for r in range(3,ws.max_row+1):
        print('查找第 %d 行' % r)
        staff = ws.cell(r, 6).value
        staffs.setdefault(staff, 200)
        if ws.cell(r, 3).value == '门诊病历':
            dept = ws.cell(r, 4).value
            dept_op.setdefault(dept,0)
            dept_op[dept] += 1
            print('门诊病历%s检查%s份' % (dept,dept_op[dept]))
            if ws.cell(r,12).value == '甲级':
                dept_opa.setdefault(dept,0)
                dept_opa[dept] += 1
                print('甲级%s份' % dept_opa[dept])
            elif ws.cell(r,12).value == '乙级':
                dept_opb.setdefault(dept,0)
                dept_opb[dept] += 1
                print('乙级%s份' % dept_opb[dept])
        elif ws.cell(r, 3).value == '环节病历':
            dept2 = ws.cell(r, 4).value
            dept_hj.setdefault(dept2,0)
            dept_hj[dept2] += 1
            print('环节病历%s检查%s份' % (dept2, dept_hj[dept2]))
            if ws.cell(r,12).value == '甲级':
                dept_hja.setdefault(dept2,0)
                dept_hja[dept2] += 1
                print('甲级%s份' % dept_hja[dept2])
            elif ws.cell(r,12).value == '乙级':
                dept_hjb.setdefault(dept2,0)
                dept_hjb[dept2] += 1
                print('乙级%s份' % dept_hjb[dept2])
        elif ws.cell(r, 3).value == '终末病历':
            dept3 = ws.cell(r, 4).value
            dept_zm.setdefault(dept3,0)
            dept_zm[dept3] += 1
            print('终末病历%s检查%s份' % (dept3, dept_zm[dept3]))
            if ws.cell(r,12).value == '甲级':
                dept_zma.setdefault(dept3,0)
                dept_zma[dept3] += 1
                print('甲级%s份' % dept_zma[dept3])
                if ws.cell(r, 13).value == '优秀':
                    dept4 = ws.cell(r, 4).value
                    dept_yx.setdefault(dept4, 0)
                    dept_yx[dept4] += 1
                    print('优秀%s份' % dept_yx[dept4])
            elif ws.cell(r,12).value == '乙级':
                dept_zmb.setdefault(dept3,0)
                dept_zmb[dept3] += 1
                print('乙级%s份' % dept_zmb[dept3])

    #创建汇总表
    wb_new = Workbook()
    ws_new = wb_new.active
    font1 = Font(size=26,bold=True)
    font2 = Font(size=14,bold=True)
    ws_new.cell(1,1).value=title
    ws_new.cell(1,1).font=font1
    print('写入门诊数据')
    ws_new.cell(2,1).value='第一部分 门诊病历结果及奖惩汇总'
    ws_new.cell(2,1).font=font2
    ws_new.column_dimensions['A'].width = 14.0
    ws_new.column_dimensions['B'].width = 14.0
    ws_new.column_dimensions['C'].width = 14.0
    ws_new.column_dimensions['D'].width = 14.0
    ws_new.column_dimensions['E'].width = 14.0
    ws_new.cell(3,1).value='科室'
    ws_new.cell(3,1).font=font2
    ws_new.cell(3,2).value='门诊病历检查份数'
    ws_new.cell(3,2).font=font2
    ws_new.cell(3,3).value='门诊甲级份数'
    ws_new.cell(3,3).font=font2
    ws_new.cell(3,4).value='门诊乙级份数'
    ws_new.cell(3,4).font=font2
    ws_new.cell(3,5).value='门诊病历扣罚(元)'
    ws_new.cell(3,5).font=font2
    #遍历门诊结果字典
    #科室名称列表
    keys=list(dept_op.keys())
    opall = 0
    opaall = 0
    opball = 0
    opvalue = 0
    for r in range(4,len(dept_op)+4):
        ws_new.cell(r,1).value=keys[r-4]
        ws_new.cell(r,2).value=dept_op[keys[r-4]]
        if keys[r-4] in dept_opa:
            ws_new.cell(r,3).value=dept_opa[keys[r-4]]
        else:
            ws_new.cell(r,3).value=0
        if keys[r-4] in dept_opb:
            ws_new.cell(r,4).value=dept_opb[keys[r-4]]
        else:
            ws_new.cell(r,4).value=0
        ws_new.cell(r,5).value=int(ws_new.cell(r,4).value)*100
        opvalue += ws_new.cell(r,5).value
    for key in keys:
        opall +=  dept_op[key]
        opaall += dept_opa[key]
        if key in list(dept_opb.keys()):
            opball += dept_opb[key]
    ws_new.cell(r+1,1).value='合计'
    ws_new.cell(r+1,1).font=font2
    ws_new.cell(r+1,2).value=opall
    ws_new.cell(r+1,2).font=font2
    ws_new.cell(r+1,3).value=opaall
    ws_new.cell(r+1,3).font=font2
    ws_new.cell(r+1,4).value=opball
    ws_new.cell(r+1,4).font=font2
    ws_new.cell(r+1,5).value=opvalue
    ws_new.cell(r+1,5).font=font2

    ws_new.cell(r+2,1).value='第二部分 环节病历结果及奖惩汇总'
    ws_new.cell(r+2,1).font=font2
    print('写入环节数据')
    ws_new.cell(r+3,1).value='科室'
    ws_new.cell(r+3,1).font=font2
    ws_new.cell(r+3,2).value='环节病历检查份数'
    ws_new.cell(r+3,2).font=font2
    ws_new.cell(r+3,3).value='环节甲级份数'
    ws_new.cell(r+3,3).font=font2
    ws_new.cell(r+3,4).value='环节乙级份数'
    ws_new.cell(r+3,4).font=font2
    ws_new.cell(r+3,5).value='环节病历扣罚(元)'
    ws_new.cell(r+3,5).font=font2
    #遍历门环节结果字典
    #科室名称列表
    keys=list(dept_hj.keys())
    hjall = 0
    hjaall = 0
    hjball = 0
    hjvalue = 0
    for r1 in range(r+4,len(dept_hj)+r+4):
        ws_new.cell(r1,1).value=keys[r1-r-4]
        ws_new.cell(r1,2).value=dept_hj[keys[r1-r-4]]
        if keys[r1-r-4] in dept_hja:
            ws_new.cell(r1,3).value=dept_hja[keys[r1-r-4]]
        else:
            ws_new.cell(r1,3).value=0
        if keys[r1-r-4] in dept_hjb:
            ws_new.cell(r1,4).value=dept_hjb[keys[r1-r-4]]
        else:
            ws_new.cell(r1,4).value=0
        ws_new.cell(r1,5).value=int(ws_new.cell(r1,4).value)*100
        hjvalue += ws_new.cell(r1, 5).value
    for key in keys:
        hjall += dept_hj[key]
        hjaall += dept_hja[key]
        if key in list(dept_hjb.keys()):
            hjball += dept_hjb[key]
    ws_new.cell(r1 + 1, 1).value = '合计'
    ws_new.cell(r1+ 1, 1).font = font2
    ws_new.cell(r1 + 1, 2).value = hjall
    ws_new.cell(r1 + 1, 2).font = font2
    ws_new.cell(r1 + 1, 3).value = hjaall
    ws_new.cell(r1 + 1, 3).font = font2
    ws_new.cell(r1 + 1, 4).value = hjball
    ws_new.cell(r1 + 1, 4).font = font2
    ws_new.cell(r1 + 1, 5).value = hjvalue
    ws_new.cell(r1 + 1, 5).font = font2

    ws_new.cell(r1+2,1).value='第三部分 终末病历结果及奖惩汇总'
    ws_new.cell(r1+2,1).font=font2
    print('写入终末数据')
    ws_new.cell(r1+3,1).value='科室'
    ws_new.cell(r1+3,1).font=font2
    ws_new.cell(r1+3,2).value='终末病历检查份数'
    ws_new.cell(r1+3,2).font=font2
    ws_new.cell(r1+3,3).value='终末甲级份数'
    ws_new.cell(r1+3,3).font=font2
    ws_new.cell(r1+3,4).value='终末乙级份数'
    ws_new.cell(r1+3,4).font=font2
    ws_new.cell(r1+3,5).value='终末病历扣罚(元)'
    ws_new.cell(r1+3,5).font=font2
    #遍历终末结果字典
    #科室名称列表
    keys=list(dept_zm.keys())
    zmall = 0
    zmaall = 0
    zmball = 0
    zmvalue = 0
    for r2 in range(r1+4,len(dept_zm)+r1+4):
        ws_new.cell(r2,1).value=keys[r2-r1-4]
        ws_new.cell(r2,2).value=dept_zm[keys[r2-r1-4]]
        if keys[r2-r1-4] in dept_zma:
            ws_new.cell(r2,3).value=dept_zma[keys[r2-r1-4]]
        else:
            ws_new.cell(r2,3).value=0
        if keys[r2-r1-4] in dept_zmb:
            ws_new.cell(r2,4).value=dept_zmb[keys[r2-r1-4]]
        else:
            ws_new.cell(r2,4).value=0
        ws_new.cell(r2,5).value=int(ws_new.cell(r2,4).value)*200
        zmvalue += ws_new.cell(r2, 5).value
    for key in keys:
        zmall += dept_zm[key]
        zmaall += dept_zma[key]
        if key in list(dept_zmb.keys()):
            zmball += dept_zmb[key]
    ws_new.cell(r2 + 1, 1).value = '合计'
    ws_new.cell(r2 + 1, 1).font = font2
    ws_new.cell(r2 + 1, 2).value = zmall
    ws_new.cell(r2 + 1, 2).font = font2
    ws_new.cell(r2 + 1, 3).value = zmaall
    ws_new.cell(r2 + 1, 3).font = font2
    ws_new.cell(r2 + 1, 4).value = zmball
    ws_new.cell(r2 + 1, 4).font = font2
    ws_new.cell(r2 + 1, 5).value = zmvalue
    ws_new.cell(r2 + 1, 5).font = font2

    ws_new.cell(r2+2,1).value='第四部分 优秀病历奖励汇总'
    ws_new.cell(r2+2,1).font=font2
    ws_new.cell(r2+3,1).value='科室'
    ws_new.cell(r2+3,1).font=font2
    ws_new.cell(r2+3,2).value='优秀病历份数'
    ws_new.cell(r2+3,2).font=font2
    ws_new.cell(r2+3,3).value='优秀病历奖励(元)'
    ws_new.cell(r2+3,3).font=font2
    print('写入优秀数据')
    #遍历优秀病历结果字典
    #科室名称列表
    keys=list(dept_yx.keys())
    yxall = 0
    yxvalue = 0
    for r3 in range(r2+4,len(dept_yx)+r2+4):
        ws_new.cell(r3,1).value=keys[r3-r2-4]
        ws_new.cell(r3,2).value=dept_yx[keys[r3-r2-4]]
        ws_new.cell(r3, 3).value = int(ws_new.cell(r3,2).value)*300
        yxvalue += ws_new.cell(r3, 3).value
        yxall += ws_new.cell(r3,2).value
    ws_new.cell(r3 + 1, 1).value = '合计'
    ws_new.cell(r3 + 1, 1).font = font2
    ws_new.cell(r3 + 1, 2).value = yxall
    ws_new.cell(r3 + 1, 2).font = font2
    ws_new.cell(r3 + 1, 3).value = yxvalue
    ws_new.cell(r3 + 1, 3).font = font2


    ws_new.cell(r3+2,1).value='第五部分 质控人员奖励汇总'
    ws_new.cell(r3+2,1).font=font2
    ws_new.cell(r3+3,1).value='姓名'
    ws_new.cell(r3+3,1).font=font2
    ws_new.cell(r3+3,2).value='科室'
    ws_new.cell(r3+3,2).font=font2
    ws_new.cell(r3+3,3).value='奖励金额(元)'
    ws_new.cell(r3+3,3).font=font2
    print('写入人员奖励数据')
    #遍历优秀病历结果字典
    #科室名称列表
    keys=list(staffs.keys())
    staffvalue = 0
    for r4 in range(r3+4,len(staffs)+r3+4):
        if keys[r4-r3-4] in ('质管办','牟园芬'):
            r4 += 1
        else:
            ws_new.cell(r4,1).value=keys[r4-r3-4]
            deptname = Staff.query.filter(Staff.name == keys[r4-r3-4]).first()
            ws_new.cell(r4, 2).value= str(deptname)
            ws_new.cell(r4,3).value=staffs[keys[r4-r3-4]]
            staffvalue += ws_new.cell(r4,3).value
    ws_new.cell(r4 + 1, 1).value = '合计'
    ws_new.cell(r4 + 1, 1).font = font2
    ws_new.cell(r4 + 1, 3).value = staffvalue
    ws_new.cell(r4 + 1, 3).font = font2

    print('打开病历迟交原始数据表')
    wbcj = load_workbook(file2,read_only=True,data_only=True)
    wscj = wbcj.active
    deptcj = {}
    deptfk = {}
    for r in range (4,wscj.max_row+1):
        print('Read line %s' % r)
        if int(wscj.cell(row=r,column=6).value) > 0:
            dept = wscj.cell(row=r,column=1).value
            deptcj.setdefault(dept,0)
            deptfk.setdefault(dept,0)
            deptcj[dept] += 1
            deptfk[dept] += int(wscj.cell(row=r,column=7).value)

    print('Writing to new workbook...')
    ws_new.cell(r4+2,1).value='第六部分 病历迟交情况及罚款汇总表'
    ws_new.cell(r4+2,1).font=font2
    ws_new.cell(r4+3,1).value='科室'
    ws_new.cell(r4+3,1).font=font2
    ws_new.cell(r4+3,2).value='迟交份数'
    ws_new.cell(r4+3,2).font=font2
    ws_new.cell(r4+3,3).value='罚款金额(元)'
    ws_new.cell(r4+3,3).font=font2
    keys = list(deptcj.keys())
    cjall = 0
    fkall = 0
    for r5 in range(r4+4,len(deptcj)+r4+4):
        ws_new.cell(r5,1).value=keys[r5-r4-4]
        ws_new.cell(r5,2).value=deptcj[keys[r5-r4-4]]
        ws_new.cell(r5,3).value=deptfk[keys[r5-r4-4]]
        cjall += ws_new.cell(r5,2).value
        fkall += ws_new.cell(r5,3).value
    ws_new.cell(r5 + 1, 1).value = '合计'
    ws_new.cell(r5 + 1, 1).font = font2
    ws_new.cell(r5 + 1, 2).value = cjall
    ws_new.cell(r5 + 1, 2).font = font2
    ws_new.cell(r5 + 1, 3).value = fkall
    ws_new.cell(r5 + 1, 3).font = font2
    ws_new.cell(r5 + 3, 1).value = '本月奖励合计'
    ws_new.cell(r5 + 3, 1).font = font2
    ws_new.cell(r5 + 3, 2).value = yxvalue+staffvalue
    ws_new.cell(r5 + 3, 2).font = font2
    ws_new.cell(r5 + 3, 3).value = '本月处罚合计'
    ws_new.cell(r5 + 3, 3).font = font2
    ws_new.cell(r5 + 3, 4).value = opvalue+hjvalue+zmvalue+fkall
    ws_new.cell(r5 + 3, 4).font = font2
    print('写入领导签名部分')
    ws_new.cell(r5+6,1).value='分管领导意见:'
    ws_new.cell(r5+8,1).value='主要领导意见:'
    ws_new.cell(r5+6,1).font=font1
    ws_new.cell(r5+8,1).font=font1



    wb_new.save(year+'年'+month+'月'+'病历奖罚通知.xlsx')
    wb_new.close()
    wb.close()



