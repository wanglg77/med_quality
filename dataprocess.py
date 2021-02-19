from define import *


@app.route('/dataprocess', methods=['GET','POST'])


def dataprocess():
    processform = Dataprocess()
    # uploadform = Upload()
    functype = processform.func_sel.data

    # 病历数据统计
    if processform.is_submitted() and functype == '1':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        file2 = processform.file2_upload.data
        file2.save(os.path.join('./soucedata/', file2.filename))
        # 病历数据统计
        date = processform.date_sel.data
        year = date.year
        month = date.month
        print(year,month,file.filename,file2.filename,functype)
        title = str(year) + '年' + str(month) + '月' + '病历质量奖罚汇总表'
        file1 = os.path.join('./soucedata/', file.filename)
        file2 = os.path.join('./soucedata/', file2.filename)
        print('导入每月病历质量数据处理')
        wb = load_workbook(file1, read_only=True)
        ws = wb.active
        # 病历评级数据
        dept_op = {}  # 科室门诊病历字典
        dept_opa = {}
        dept_opb = {}
        dept_hj = {}  # 科室环节病历字典
        dept_hja = {}
        dept_hjb = {}
        dept_zm = {}  # 科室终末病历字典
        dept_zma = {}
        dept_zmb = {}
        dept_yx = {}  # 科室优秀病历字典
        staffs = {}

        for r in range(3, ws.max_row + 1):
            # print('查找第 %d 行' % r)
            staff = ws.cell(r, 6).value
            staffs.setdefault(staff, 200)
            if ws.cell(r, 3).value == '门诊病历':
                dept = ws.cell(r, 4).value
                dept_op.setdefault(dept, 0)
                dept_op[dept] += 1
                # print('门诊病历%s检查%s份' % (dept, dept_op[dept]))
                if ws.cell(r, 12).value == '甲级':
                    dept_opa.setdefault(dept, 0)
                    dept_opa[dept] += 1
                    # print('甲级%s份' % dept_opa[dept])
                elif ws.cell(r, 12).value == '乙级':
                    dept_opb.setdefault(dept, 0)
                    dept_opb[dept] += 1
                    # print('乙级%s份' % dept_opb[dept])
            elif ws.cell(r, 3).value == '环节病历':
                dept2 = ws.cell(r, 4).value
                dept_hj.setdefault(dept2, 0)
                dept_hj[dept2] += 1
                # print('环节病历%s检查%s份' % (dept2, dept_hj[dept2]))
                if ws.cell(r, 12).value == '甲级':
                    dept_hja.setdefault(dept2, 0)
                    dept_hja[dept2] += 1
                    # print('甲级%s份' % dept_hja[dept2])
                elif ws.cell(r, 12).value == '乙级':
                    dept_hjb.setdefault(dept2, 0)
                    dept_hjb[dept2] += 1
                    # print('乙级%s份' % dept_hjb[dept2])
            elif ws.cell(r, 3).value == '终末病历':
                dept3 = ws.cell(r, 4).value
                dept_zm.setdefault(dept3, 0)
                dept_zm[dept3] += 1
                # print('终末病历%s检查%s份' % (dept3, dept_zm[dept3]))
                if ws.cell(r, 12).value == '甲级':
                    dept_zma.setdefault(dept3, 0)
                    dept_zma[dept3] += 1
                    # print('甲级%s份' % dept_zma[dept3])
                    if ws.cell(r, 13).value == '优秀':
                        dept4 = ws.cell(r, 4).value
                        dept_yx.setdefault(dept4, 0)
                        dept_yx[dept4] += 1
                        # print('优秀%s份' % dept_yx[dept4])
                elif ws.cell(r, 12).value == '乙级':
                    dept_zmb.setdefault(dept3, 0)
                    dept_zmb[dept3] += 1
                    # print('乙级%s份' % dept_zmb[dept3])

        # 创建汇总表
        wb_new = Workbook()
        ws_new = wb_new.active
        font1 = Font(size=26, bold=True)
        font2 = Font(size=14, bold=True)
        ws_new.cell(1, 1).value = title
        ws_new.cell(1, 1).font = font1
        print('写入门诊数据')
        ws_new.cell(2, 1).value = '第一部分 门诊病历结果及奖惩汇总'
        ws_new.cell(2, 1).font = font2
        ws_new.column_dimensions['A'].width = 14.0
        ws_new.column_dimensions['B'].width = 14.0
        ws_new.column_dimensions['C'].width = 14.0
        ws_new.column_dimensions['D'].width = 14.0
        ws_new.column_dimensions['E'].width = 14.0
        ws_new.cell(3, 1).value = '科室'
        ws_new.cell(3, 1).font = font2
        ws_new.cell(3, 2).value = '门诊病历检查份数'
        ws_new.cell(3, 2).font = font2
        ws_new.cell(3, 3).value = '门诊甲级份数'
        ws_new.cell(3, 3).font = font2
        ws_new.cell(3, 4).value = '门诊乙级份数'
        ws_new.cell(3, 4).font = font2
        ws_new.cell(3, 5).value = '门诊病历扣罚(元)'
        ws_new.cell(3, 5).font = font2
        # 遍历门诊结果字典
        # 科室名称列表
        keys = list(dept_op.keys())
        opall = 0
        opaall = 0
        opball = 0
        opvalue = 0
        for r in range(4, len(dept_op) + 4):
            ws_new.cell(r, 1).value = keys[r - 4]
            ws_new.cell(r, 2).value = dept_op[keys[r - 4]]
            if keys[r - 4] in dept_opa:
                ws_new.cell(r, 3).value = dept_opa[keys[r - 4]]
            else:
                ws_new.cell(r, 3).value = 0
            if keys[r - 4] in dept_opb:
                ws_new.cell(r, 4).value = dept_opb[keys[r - 4]]
            else:
                ws_new.cell(r, 4).value = 0
            ws_new.cell(r, 5).value = int(ws_new.cell(r, 4).value) * 100
            opvalue += ws_new.cell(r, 5).value
        for key in keys:
            opall += dept_op[key]
            opaall += dept_opa[key]
            if key in list(dept_opb.keys()):
                opball += dept_opb[key]
        ws_new.cell(r + 1, 1).value = '合计'
        ws_new.cell(r + 1, 1).font = font2
        ws_new.cell(r + 1, 2).value = opall
        ws_new.cell(r + 1, 2).font = font2
        ws_new.cell(r + 1, 3).value = opaall
        ws_new.cell(r + 1, 3).font = font2
        ws_new.cell(r + 1, 4).value = opball
        ws_new.cell(r + 1, 4).font = font2
        ws_new.cell(r + 1, 5).value = opvalue
        ws_new.cell(r + 1, 5).font = font2

        ws_new.cell(r + 2, 1).value = '第二部分 环节病历结果及奖惩汇总'
        ws_new.cell(r + 2, 1).font = font2
        print('写入环节数据')
        ws_new.cell(r + 3, 1).value = '科室'
        ws_new.cell(r + 3, 1).font = font2
        ws_new.cell(r + 3, 2).value = '环节病历检查份数'
        ws_new.cell(r + 3, 2).font = font2
        ws_new.cell(r + 3, 3).value = '环节甲级份数'
        ws_new.cell(r + 3, 3).font = font2
        ws_new.cell(r + 3, 4).value = '环节乙级份数'
        ws_new.cell(r + 3, 4).font = font2
        ws_new.cell(r + 3, 5).value = '环节病历扣罚(元)'
        ws_new.cell(r + 3, 5).font = font2
        # 遍历门环节结果字典
        # 科室名称列表
        keys = list(dept_hj.keys())
        hjall = 0
        hjaall = 0
        hjball = 0
        hjvalue = 0
        for r1 in range(r + 4, len(dept_hj) + r + 4):
            ws_new.cell(r1, 1).value = keys[r1 - r - 4]
            ws_new.cell(r1, 2).value = dept_hj[keys[r1 - r - 4]]
            if keys[r1 - r - 4] in dept_hja:
                ws_new.cell(r1, 3).value = dept_hja[keys[r1 - r - 4]]
            else:
                ws_new.cell(r1, 3).value = 0
            if keys[r1 - r - 4] in dept_hjb:
                ws_new.cell(r1, 4).value = dept_hjb[keys[r1 - r - 4]]
            else:
                ws_new.cell(r1, 4).value = 0
            ws_new.cell(r1, 5).value = int(ws_new.cell(r1, 4).value) * 100
            hjvalue += ws_new.cell(r1, 5).value
        for key in keys:
            hjall += dept_hj[key]
            hjaall += dept_hja[key]
            if key in list(dept_hjb.keys()):
                hjball += dept_hjb[key]
        ws_new.cell(r1 + 1, 1).value = '合计'
        ws_new.cell(r1 + 1, 1).font = font2
        ws_new.cell(r1 + 1, 2).value = hjall
        ws_new.cell(r1 + 1, 2).font = font2
        ws_new.cell(r1 + 1, 3).value = hjaall
        ws_new.cell(r1 + 1, 3).font = font2
        ws_new.cell(r1 + 1, 4).value = hjball
        ws_new.cell(r1 + 1, 4).font = font2
        ws_new.cell(r1 + 1, 5).value = hjvalue
        ws_new.cell(r1 + 1, 5).font = font2

        ws_new.cell(r1 + 2, 1).value = '第三部分 终末病历结果及奖惩汇总'
        ws_new.cell(r1 + 2, 1).font = font2
        print('写入终末数据')
        ws_new.cell(r1 + 3, 1).value = '科室'
        ws_new.cell(r1 + 3, 1).font = font2
        ws_new.cell(r1 + 3, 2).value = '终末病历检查份数'
        ws_new.cell(r1 + 3, 2).font = font2
        ws_new.cell(r1 + 3, 3).value = '终末甲级份数'
        ws_new.cell(r1 + 3, 3).font = font2
        ws_new.cell(r1 + 3, 4).value = '终末乙级份数'
        ws_new.cell(r1 + 3, 4).font = font2
        ws_new.cell(r1 + 3, 5).value = '终末病历扣罚(元)'
        ws_new.cell(r1 + 3, 5).font = font2
        # 遍历终末结果字典
        # 科室名称列表
        keys = list(dept_zm.keys())
        zmall = 0
        zmaall = 0
        zmball = 0
        zmvalue = 0
        for r2 in range(r1 + 4, len(dept_zm) + r1 + 4):
            ws_new.cell(r2, 1).value = keys[r2 - r1 - 4]
            ws_new.cell(r2, 2).value = dept_zm[keys[r2 - r1 - 4]]
            if keys[r2 - r1 - 4] in dept_zma:
                ws_new.cell(r2, 3).value = dept_zma[keys[r2 - r1 - 4]]
            else:
                ws_new.cell(r2, 3).value = 0
            if keys[r2 - r1 - 4] in dept_zmb:
                ws_new.cell(r2, 4).value = dept_zmb[keys[r2 - r1 - 4]]
            else:
                ws_new.cell(r2, 4).value = 0
            ws_new.cell(r2, 5).value = int(ws_new.cell(r2, 4).value) * 200
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

        ws_new.cell(r2 + 2, 1).value = '第四部分 优秀病历奖励汇总'
        ws_new.cell(r2 + 2, 1).font = font2
        ws_new.cell(r2 + 3, 1).value = '科室'
        ws_new.cell(r2 + 3, 1).font = font2
        ws_new.cell(r2 + 3, 2).value = '优秀病历份数'
        ws_new.cell(r2 + 3, 2).font = font2
        ws_new.cell(r2 + 3, 3).value = '优秀病历奖励(元)'
        ws_new.cell(r2 + 3, 3).font = font2
        print('写入优秀数据')
        # 遍历优秀病历结果字典
        # 科室名称列表
        keys = list(dept_yx.keys())
        yxall = 0
        yxvalue = 0
        for r3 in range(r2 + 4, len(dept_yx) + r2 + 4):
            ws_new.cell(r3, 1).value = keys[r3 - r2 - 4]
            ws_new.cell(r3, 2).value = dept_yx[keys[r3 - r2 - 4]]
            ws_new.cell(r3, 3).value = int(ws_new.cell(r3, 2).value) * 300
            yxvalue += ws_new.cell(r3, 3).value
            yxall += ws_new.cell(r3, 2).value
        r3 = len(dept_yx) + r2 + 4
        ws_new.cell(r3 + 1, 1).value = '合计'
        ws_new.cell(r3 + 1, 1).font = font2
        ws_new.cell(r3 + 1, 2).value = yxall
        ws_new.cell(r3 + 1, 2).font = font2
        ws_new.cell(r3 + 1, 3).value = yxvalue
        ws_new.cell(r3 + 1, 3).font = font2

        ws_new.cell(r3 + 2, 1).value = '第五部分 质控人员奖励汇总'
        ws_new.cell(r3 + 2, 1).font = font2
        ws_new.cell(r3 + 3, 1).value = '姓名'
        ws_new.cell(r3 + 3, 1).font = font2
        ws_new.cell(r3 + 3, 2).value = '科室'
        ws_new.cell(r3 + 3, 2).font = font2
        ws_new.cell(r3 + 3, 3).value = '奖励金额(元)'
        ws_new.cell(r3 + 3, 3).font = font2
        print('写入人员奖励数据')
        # 遍历优秀病历结果字典
        # 科室名称列表
        keys = list(staffs.keys())
        staffvalue = 0
        for r4 in range(r3 + 4, len(staffs) + r3 + 4):
            if keys[r4 - r3 - 4] in ('质管办', '牟园芬'):
                r4 += 1
            else:
                ws_new.cell(r4, 1).value = keys[r4 - r3 - 4]
                deptname = Staff.query.filter(Staff.name == keys[r4 - r3 - 4]).first()
                ws_new.cell(r4, 2).value = str(deptname)
                ws_new.cell(r4, 3).value = staffs[keys[r4 - r3 - 4]]
                staffvalue += ws_new.cell(r4, 3).value
        ws_new.cell(r4 + 1, 1).value = '合计'
        ws_new.cell(r4 + 1, 1).font = font2
        ws_new.cell(r4 + 1, 3).value = staffvalue
        ws_new.cell(r4 + 1, 3).font = font2

        print('打开病历迟交原始数据表')
        wbcj = load_workbook(file2, read_only=True,
                             data_only=True)
        wscj = wbcj.active
        deptcj = {}
        deptfk = {}
        print('开始处理迟交病例数据...')
        for r in range(4, wscj.max_row + 1):
            # print('Read line %s' % r)
            if wscj.cell(row=r, column=6).value:
                if int(wscj.cell(row=r, column=6).value) > 0:
                    dept = wscj.cell(row=r, column=1).value
                    deptcj.setdefault(dept, 0)
                    deptfk.setdefault(dept, 0)
                    deptcj[dept] += 1
                    deptfk[dept] += int(wscj.cell(row=r, column=7).value)

        print('Writing to new workbook...')
        ws_new.cell(r4 + 2, 1).value = '第六部分 病历迟交情况及罚款汇总表'
        ws_new.cell(r4 + 2, 1).font = font2
        ws_new.cell(r4 + 3, 1).value = '科室'
        ws_new.cell(r4 + 3, 1).font = font2
        ws_new.cell(r4 + 3, 2).value = '迟交份数'
        ws_new.cell(r4 + 3, 2).font = font2
        ws_new.cell(r4 + 3, 3).value = '罚款金额(元)'
        ws_new.cell(r4 + 3, 3).font = font2
        keys = list(deptcj.keys())
        cjall = 0
        fkall = 0
        for r5 in range(r4 + 4, len(deptcj) + r4 + 4):
            ws_new.cell(r5, 1).value = keys[r5 - r4 - 4]
            ws_new.cell(r5, 2).value = deptcj[keys[r5 - r4 - 4]]
            ws_new.cell(r5, 3).value = deptfk[keys[r5 - r4 - 4]]
            cjall += ws_new.cell(r5, 2).value
            fkall += ws_new.cell(r5, 3).value
        ws_new.cell(r5 + 1, 1).value = '合计'
        ws_new.cell(r5 + 1, 1).font = font2
        ws_new.cell(r5 + 1, 2).value = cjall
        ws_new.cell(r5 + 1, 2).font = font2
        ws_new.cell(r5 + 1, 3).value = fkall
        ws_new.cell(r5 + 1, 3).font = font2
        ws_new.cell(r5 + 3, 1).value = '本月奖励合计'
        ws_new.cell(r5 + 3, 1).font = font2
        ws_new.cell(r5 + 3, 2).value = yxvalue + staffvalue
        ws_new.cell(r5 + 3, 2).font = font2
        ws_new.cell(r5 + 3, 3).value = '本月处罚合计'
        ws_new.cell(r5 + 3, 3).font = font2
        ws_new.cell(r5 + 3, 4).value = opvalue + hjvalue + zmvalue + fkall
        ws_new.cell(r5 + 3, 4).font = font2
        print('写入领导签名部分')
        ws_new.cell(r5 + 6, 1).value = '分管领导意见:'
        ws_new.cell(r5 + 8, 1).value = '主要领导意见:'
        ws_new.cell(r5 + 6, 1).font = font1
        ws_new.cell(r5 + 8, 1).font = font1
        filename = str(year) + '年' + str(month) + '月' + '病历奖罚通知.xlsx'
        wb_new.save(os.path.join('./temp/', filename))
        wb_new.close()
        wb.close()
    # 病历数据导入数据库
    elif processform.is_submitted() and functype == '2':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # 病历数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        # 遍历质控数据表
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        # 定义段落标记行号
        rpgroup = {2:0,3:0,4:0,5:0,6:0,7:0}
        for r in range(4, ws.max_row + 1):
            if ws.cell(r, 1).value == '第二部分 环节病历结果及奖惩汇总':
                rpgroup[2] = r
            elif ws.cell(r, 1).value == '第三部分 终末病历结果及奖惩汇总':
                rpgroup[3] = r
            elif ws.cell(r, 1).value == '第四部分 优秀病历奖励汇总':
                rpgroup[4] = r
            elif ws.cell(r, 1).value == '第五部分 质控人员奖励汇总':
                rpgroup[5] = r
            elif ws.cell(r, 1).value == '第六部分 病历迟交情况及罚款汇总表':
                rpgroup[6] = r
            elif ws.cell(r, 1).value == '本月奖励合计':
                rpgroup[7] = r
        print(rpgroup)
        rp2 = rpgroup[2]
        print(rp2)
        rp3 = rpgroup[3]
        rp4 = rpgroup[4]
        rp5 = rpgroup[5]
        rp6 = rpgroup[6]
        rp7 = rpgroup[7]
        # 查询写入门诊部分
        flash('正在处理门诊病历数据...')
        for r in range(4, rp2 - 1):
            deptname = ws.cell(r, 1).value
            CheckNumOPRec = ws.cell(r, 2).value
            ANumOPRec = ws.cell(r, 3).value
            BNumOPRec = ws.cell(r, 4).value
            print(deptname,CheckNumOPRec,ANumOPRec)

            opdepts = ['产科', '急诊科盐田', '甲乳外科', '精神科', '康复医学科', '口腔科', '泌尿外科', '康复医学科盐田',
                       '内分泌科', '皮肤科', '普外科', '神经内科', '神经外科', '肾内科', '体检科', '体检科盐田', '消化内科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科', '普外科盐田',
                       '急诊科', '胸外科', '血液内科', '眼科', '针灸推拿科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '中医科', '肿瘤科']
            if deptname in opdepts:
                ifexist = OutPatient.query.filter_by(deptname=deptname, date=date).first()
                if ifexist:
                    ifexist.CheckNumOPRec = CheckNumOPRec
                    ifexist.ANumOPRec = ANumOPRec
                    ifexist.BNumOPRec = BNumOPRec
                    ifexist.CNumOPRec = 0
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = OutPatient(deptname=deptname, date=date, CheckNumOPRec=CheckNumOPRec,
                                      ANumOPRec=ANumOPRec, BNumOPRec=BNumOPRec, CNumOPRec=0)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()

        # 查询写入环节部分
        flash('正在处理环节病历数据...')
        for r in range(rp2 + 2, rp3 - 1):
            deptname = ws.cell(r, 1).value
            CheckNumIPRec = ws.cell(r, 2).value
            ANumIPRec = ws.cell(r, 3).value
            BNumIPRec = ws.cell(r, 4).value
            ipdepts = ['产科', '甲乳外科', '康复医学科', '泌尿外科', '普外科盐田', '康复医学科盐田',
                       '内分泌科', '普外科', '神经内科', '神经外科', '肾内科', '消化内科', '新生儿科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科',
                       '胸外科', '血液内科', '眼科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '重症医学科', '肿瘤科']
            if deptname in ipdepts:
                ifexist = InPatient.query.filter_by(deptname=deptname, date=date).first()
                if ifexist:
                    ifexist.CheckNumIPRec = CheckNumIPRec
                    ifexist.ANumIPRec = ANumIPRec
                    ifexist.BNumIPRec = BNumIPRec
                    ifexist.CNumIPRec = 0
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = InPatient(deptname=deptname, date=date, CheckNumIPRec=CheckNumIPRec,
                                     ANumIPRec=ANumIPRec, BNumIPRec=BNumIPRec, CNumIPRec=0)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
        # 查询写入终末部分
        flash('正在处理终末病历数据...')
        for r in range(rp3 + 2, rp4 - 1):
            deptname = ws.cell(r, 1).value
            print('正在处理 %s 数据' % deptname)
            cnip = InPatient.query.filter_by(deptname=deptname, date=date).first()
            # if cnip is None:
            #     cnip = 0
            CheckNumIPRec = int(ws.cell(r, 2).value) + int(cnip.CheckNumIPRec)
            ANumIPRec = int(ws.cell(r, 3).value) + int(cnip.ANumIPRec)
            BNumIPRec = int(ws.cell(r, 4).value) + int(cnip.BNumIPRec)
            cnip.CheckNumIPRec = CheckNumIPRec
            cnip.ANumIPRec = ANumIPRec
            cnip.BNumIPRec = BNumIPRec
            try:
                db.session.commit()
            except Exception as e:
                db.session.rollback()
        # 查询写入迟交部分
        flash('正在处理病历归档数据...')
        for r in range(rp6 + 2, rp7 - 2):
            deptname = ws.cell(r, 1).value
            OANumIPRec = ws.cell(r, 2).value
            data = InPatient.query.filter_by(deptname=deptname, date=date).first()
            ipdepts = ['产科', '甲乳外科', '康复医学科', '泌尿外科', '普外科盐田',
                       '内分泌科', '普外科', '神经内科', '神经外科', '肾内科', '消化内科', '新生儿科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科',
                       '胸外科', '血液内科', '眼科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '重症医学科', '肿瘤科']
            if deptname in ipdepts:
                if data:
                    data.OANumIPRec = OANumIPRec
                else:
                    data_new = InPatient(deptname=deptname, date=date, OANumIPRec=OANumIPRec)
                    db.session.add(data_new)
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
        flash('病历质量数据写入数据库完成!')
        wb.close()
    # 危急值数据导入数据库
    elif processform.is_submitted() and functype == '3':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # 危急值数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        # date = datetime.datetime.strptime(date, '%Y-%m-%d')
        # 遍历危急值数据表
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            deptname = ws.cell(r, 1).value
            print('正在处理%s数据...' % deptname)
            CriValRepNum = ws.cell(r, 3).value
            CriValHPNum = ws.cell(r, 4).value
            print(deptname, CriValRepNum, CriValHPNum)
            data = Quality(deptname=deptname, date=date, CriValRepNum=CriValRepNum, CriValHPNum=CriValHPNum)
            ifexist = Quality.query.filter_by(deptname=deptname, date=date).first()
            if ifexist:
                ifexist.CriValRepNum = CriValRepNum
                ifexist.CriValHPNum = CriValHPNum
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                try:
                    db.session.add(data)
                    db.session.commit()
                    print('数据添加成功')
                except Exception as e:
                    db.session.rollback()
                    print('数据添加失败')
        flash('危急值数据添加完毕!')
        wb.close()
    # 不良事件数据导入数据库
    elif processform.is_submitted() and functype == '4':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        file2 = processform.file2_upload.data
        file2.save(os.path.join('./soucedata/', file2.filename))
        # 不良事件数据导入数据库
        date = processform.date_sel.data
        file1 = os.path.join('./soucedata/', file.filename)
        file2 = os.path.join('./soucedata/', file2.filename)
        wb1 = load_workbook(file1)
        wb2 = load_workbook(file2)
        ws1 = wb1.active
        ws2 = wb2.active
        # 定义修订科室名称方法
        flash('正在规范数据...')
        # 定义对照字典#
        DeptNames = {"产科一": "产科", "手术室": "麻醉科", "妇科盐田院区": "妇产科盐田",
                         "手术室盐田院区": "麻醉科", "产房": "产科", "输液室盐田院区": "急诊科盐田",
                         "外科盐田院区": "普外科盐田", "肛肠外科盐田院区": "中西医结合肛肠科",
                         "骨科盐田院区": "骨科盐田", "血液透析室": "肾内科", "体检中心": "体检科",
                         "内科及老年病科盐田院区": "中西医结合老年病科", "针灸理疗科": "针灸推拿科",
                         "放射科": "放射影像科", "盐田院区麻醉科": "麻醉科",
                         "心血管内科": "中西医结合心血管内科", "耳鼻咽喉科": "耳鼻喉科",
                         "结控门诊": "传染科", "发热门诊": "传染科", "感染科": "传染科", "普通外科": "普外科",
                         "急诊外科盐田院区": "急诊科盐田", "盐田院区放射影像科": "放射影像科",
                         "十九病区": "中西医结合老年病科", "检验科（盐田院区）": "检验科",
                         "康复医学科盐田院区": "康复医学科盐田", "甲状腺乳腺外科": "甲乳外科",
                         "呼吸与危重症医学科": "呼吸内科", "内科及老年病科": "中西医结合老年病科",
                         "肛肠科": "中西医结合肛肠科", "口腔科盐田院区": "口腔科", "产科二": "产科",
                         "肝病门诊": "传染科", "内分泌科盐田院区": "内分泌科", "门诊西药房": "药剂科",
                         "临床医技科室 > 超声影像科": "超声影像科", "临床医技科室 > 针灸推拿科": "针灸推拿科",
                         "临床医技科室 > 检验科": "检验科", "临床医技科室 > 皮肤科": "皮肤科",
                         "临床医技科室 > 病理科": "病理科", "急诊医学科": "急诊科","临床科室 > 骨科":"骨科",
                        "临床科室 > 针灸推拿科":"针灸推拿科","医技科室 > 超声影像科":"超声影像科","临床科室 > 麻醉科":"麻醉科",
                     "临床科室 > 中西医结合心血管内科":"中西医结合心血管内科","临床科室 > 体检科":"体检科",
                     "医技科室 > 检验科":"检验科","临床科室 > 皮肤科":"皮肤科","医技科室 > 病理科":"病理科",
                     "医技科室 > 放射影像科":"放射影像科","社康中心 > 小梅沙社康":"小梅沙社康","临床科室 > 康复医学科":"康复医学科",
                     "社康中心 > 东海社康":"东海社康","社康中心 > 田东社康":"田东社康","临床科室 > 新生儿科":"新生儿科",'临床科室 > 重症医学科':'重症医学科'
                     ,'临床科室 > 中西医结合肛肠科':'中西医结合肛肠科','临床科室 > 普外科':'普外科','临床科室 > 急诊科盐田':'急诊科盐田',
                     '临床科室 > 儿科':'儿科','临床科室 > 胸外科':'胸外科','临床科室 > 肿瘤科':'肿瘤科','临床科室 > 妇科':'妇科',
                     '临床科室 > 肾内科(含血液透析)':'肾内科','临床科室 > 普外科盐田':'普外科盐田'
                     }
        # 处理院内系统数据规范科室名称
        for r in range(1, ws1.max_row + 1):
            for i_sp in DeptNames:
                if i_sp == ws1.cell(row=r,column=10).value:
                    ws1.cell(row=r,column=10).value = DeptNames[i_sp]
        wb1.save("temp1.xlsx")
        wb1.close()
        # 处理麦客系统数据规范科室名称
        for r in range(1, ws2.max_row + 1):
            for i_sp in DeptNames:
                if i_sp == ws2.cell(row=r,column=112).value:
                    ws2.cell(row=r,column=112).value = DeptNames[i_sp]
        wb2.save("temp2.xlsx")
        wb2.close()
        # 定义统计方法
        wb1 = load_workbook("temp1.xlsx")
        wb2 = load_workbook("temp2.xlsx")
        ws1 = wb1.active
        ws2 = wb2.active
        deptcount1 = {}
        deptcount2 = {}
        for r1 in range(2, ws1.max_row + 1):
            deptname1 = ws1.cell(row=r1,column=10).value
            deptcount1.setdefault(deptname1, 0)
            deptcount1[deptname1] += 1
        for r2 in range(2, ws2.max_row + 1):
            deptname2 = ws2.cell(row=r2,column=112).value
            deptcount2.setdefault(deptname2, 0)
            deptcount2[deptname2] += 1

        deptcount = dict(Counter(deptcount1)+Counter(deptcount2))
        print(deptcount1,deptcount2,deptcount)
        # 定义导入数据库方法
        # 遍历危急值数据表
        flash('正在添加到数据库...')
        for dept in deptcount:
            depts = ['病理科', '产科', '急诊科盐田', '甲乳外科', '检验科', '精神科', '康复医学科', '口腔科', '麻醉科', '泌尿外科',
                     '内分泌科', '皮肤科', '普外科', '神经内科', '神经外科', '肾内科', '体检科', '体检科盐田', '消化内科', '新生儿科',
                     '超声影像科', '传染科', '儿科', '耳鼻喉科', '放射影像科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科',
                     '急诊科', '胸外科', '血液内科', '眼科', '药剂科', '针灸推拿科', '中西医结合肛肠科', '中西医结合老年病科',
                     '中西医结合心血管内科', '中医科', '重症医学科', '肿瘤科', '普外科盐田']
            if dept in depts:
                NumAdvEvtRep = deptcount[dept]
                print("正在处理 %s 数据" % dept)
                data = Quality.query.filter_by(deptname=dept, date=date).first()
                data.NumAdvEvtRep = NumAdvEvtRep
                try:
                    db.session.commit()
                    print('数据添加成功')
                except Exception as e:
                    db.session.rollback()
                    print('数据添加失败')
        wb1.close()
        wb2.close()
        flash('不良事件数据添加完毕...')
    # HIS数据导入
    elif processform.is_submitted() and functype == '5':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # HIS数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        date = processform.date_sel.data
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            print(ws.max_row)
            deptname = ws.cell(r, 1).value
            print('正在处理%s数据...' % deptname)
            NumOP = ws.cell(r, 3).value
            NumIP = ws.cell(r, 4).value
            NumBD = ws.cell(r, 5).value
            NumRefOP = ws.cell(r, 6).value
            NumRefIP = ws.cell(r, 7).value
            NumSurg = ws.cell(r, 8).value
            NumEleSurg = ws.cell(r, 9).value
            NumDaySurg = ws.cell(r, 10).value
            NumMiroInvaSurg = ws.cell(r, 11).value
            Num4thSurg = ws.cell(r, 12).value
            NumPath = ws.cell(r, 13).value
            NumAppt = ws.cell(r, 14).value
            WaitTimeAppt = ws.cell(r, 15).value
            OPNumInDist = ws.cell(r, 16).value
            NumCTMri = ws.cell(r, 17).value
            NumCTMriPositive = ws.cell(r, 18).value
            NumRp = ws.cell(r, 19).value
            NumEssDrugOP = ws.cell(r, 20).value
            NumEssDrugIP = ws.cell(r, 21).value
            DDD = ws.cell(r, 22).value
            totalincome = ws.cell(r, 23).value
            drugincome = ws.cell(r, 24).value
            consumIncome = ws.cell(r, 25).value
            examIncome = ws.cell(r, 26).value
            testIncome = ws.cell(r, 27).value
            pureIncome = ws.cell(r, 28).value
            adjDrugincome = ws.cell(r, 29).value
            oPIncome = ws.cell(r, 30).value
            iPIncome = ws.cell(r, 31).value
            NumOPRec = ws.cell(r, 32).value
            OutPatientdataexist = OutPatient.query.filter_by(deptname=deptname, date=date).first()
            InPatientdataexist = InPatient.query.filter_by(deptname=deptname, date=date).first()
            Incomedataexist = Income.query.filter_by(deptname=deptname, date=date).first()
            PharmacyDataexist = PharmacyData.query.filter_by(deptname=deptname, date=date).first()
            Qualitydataexist = Quality.query.filter_by(deptname=deptname, date=date).first()
            Surgerydataexist = Surgery.query.filter_by(deptname=deptname, date=date).first()
            # 门诊数据
            opdepts = ['产科', '急诊科盐田', '甲乳外科', '精神科', '康复医学科', '口腔科', '泌尿外科', '康复医学科盐田',
                       '内分泌科', '皮肤科', '普外科', '神经内科', '神经外科', '肾内科', '体检科', '体检科盐田', '消化内科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科', '普外科盐田',
                       '急诊科', '胸外科', '血液内科', '眼科', '针灸推拿科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '中医科', '肿瘤科']
            if deptname in opdepts:
                if OutPatientdataexist:
                    OutPatientdataexist.NumOP = NumOP
                    OutPatientdataexist.NumRefOP = NumRefOP
                    OutPatientdataexist.NumOPRec = NumOPRec
                    OutPatientdataexist.NumAppt = NumAppt
                    OutPatientdataexist.WaitTimeAppt = WaitTimeAppt
                    OutPatientdataexist.OPNumInDist = OPNumInDist
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = OutPatient(deptname=deptname, date=date, NumOP=NumOP, NumRefOP=NumRefOP,
                                      NumOPRec=NumOPRec, NumAppt=NumAppt,
                                      WaitTimeAppt=WaitTimeAppt, OPNumInDist=OPNumInDist)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
            # 住院数据
            ipdepts = ['产科', '甲乳外科', '康复医学科', '泌尿外科', '普外科盐田', '康复医学科盐田',
                       '内分泌科', '普外科', '神经内科', '神经外科', '肾内科', '消化内科', '新生儿科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科',
                       '胸外科', '血液内科', '眼科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '重症医学科', '肿瘤科']
            if deptname in ipdepts:
                if InPatientdataexist:
                    InPatientdataexist.NumIP = NumIP
                    InPatientdataexist.NumBD = NumBD
                    InPatientdataexist.NumRefIP = NumRefIP
                    InPatientdataexist.NumPath = NumPath
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = InPatient(deptname=deptname, date=date, NumIP=NumIP, NumBD=NumBD,
                                     NumRefIP=NumRefIP, NumPath=NumPath)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
            # 收入数据
            if Incomedataexist:
                Incomedataexist.totalincome = totalincome
                Incomedataexist.drugincome = drugincome
                Incomedataexist.adjDrugincome = adjDrugincome
                Incomedataexist.consumIncome = consumIncome
                Incomedataexist.examIncome = examIncome
                Incomedataexist.testIncome = testIncome
                Incomedataexist.oPIncome = oPIncome
                Incomedataexist.iPIncome = iPIncome
                Incomedataexist.pureIncome = pureIncome
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = Income(deptname=deptname, date=date, totalincome=totalincome, drugincome=drugincome,
                              adjDrugincome=adjDrugincome, consumIncome=consumIncome,
                              examIncome=examIncome, testIncome=testIncome, oPIncome=oPIncome,
                              iPIncome=iPIncome, pureIncome=pureIncome)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            # 质量数据
            if Qualitydataexist:
                Qualitydataexist.NumCTMri = NumCTMri
                Qualitydataexist.NumCTMriPositive = NumCTMriPositive
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = Quality(deptname=deptname, date=date, NumCTMri=NumCTMri,
                               NumCTMriPositive=NumCTMriPositive)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            # 手术数据
            surgdepts = ['产科', '甲乳外科', '口腔科', '麻醉科', '泌尿外科',
                         '普外科', '神经内科', '神经外科', '普外科盐田',
                         '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田',
                         '胸外科', '眼科', '中西医结合肛肠科', '中西医结合心血管内科']
            if deptname in surgdepts:
                if Surgerydataexist:
                    Surgerydataexist.NumSurg = NumSurg
                    Surgerydataexist.NumEleSurg = NumEleSurg
                    Surgerydataexist.NumDaySurg = NumDaySurg
                    Surgerydataexist.NumMiroInvaSurg = NumMiroInvaSurg
                    Surgerydataexist.Num4thSurg = Num4thSurg
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = Surgery(deptname=deptname, date=date, NumSurg=NumSurg, NumEleSurg=NumEleSurg,
                                   NumDaySurg=NumDaySurg, NumMiroInvaSurg=NumMiroInvaSurg,
                                   Num4thSurg=Num4thSurg)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
            # 药品数据
            # depts = ['产科', '急诊科盐田', '甲乳外科', '精神科', '康复医学科', '口腔科', '麻醉科', '泌尿外科',
            #          '内分泌科', '皮肤科', '普外科', '神经内科', '神经外科', '肾内科', '体检科', '体检科盐田', '消化内科', '新生儿科',
            #          '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科','普外科盐田',
            #          '急诊科', '胸外科', '血液内科', '眼科', '药剂科', '针灸推拿科', '中西医结合肛肠科', '中西医结合老年病科',
            #          '中西医结合心血管内科', '中医科', '重症医学科', '肿瘤科']
            # if deptname in depts:
            if PharmacyDataexist:
                PharmacyDataexist.NumRp = NumRp
                PharmacyDataexist.NumEssDrugOP = NumEssDrugOP
                PharmacyDataexist.NumEssDrugIP = NumEssDrugIP
                PharmacyDataexist.DDD = DDD
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = PharmacyData(deptname=deptname, date=date, NumRp=NumRp, NumEssDrugOP=NumEssDrugOP,
                                    NumEssDrugIP=NumEssDrugIP, DDD=DDD)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
        flash('导入HIS数据完毕!')
        wb.close()
    # QCC数据导入
    elif processform.is_submitted() and functype == '6':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # QCC数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        # date = datetime.datetime.strptime(date, '%Y-%m-%d')
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        DeptNames = {"盐田院区 > 骨科盐田": "骨科盐田", "本部院区 > 妇科": "妇科", "本部院区 > 甲乳外科": "甲乳外科",
                     "本部院区 > 检验科": "检验科", "本部院区 > 肿瘤科": "肿瘤科", "本部院区 > 康复医学科": "康复医学科",
                     "本部院区 > 重症医学科": "重症医学科", "本部院区 > 血液内科": "血液内科",
                     "本部院区 > 呼吸内科": "呼吸内科", "本部院区 > 中医科": "中医科", "本部院区 > 针灸推拿科": "针灸推拿科",
                     "本部院区 > 药剂科": "药剂科", "针灸理疗科": "针灸推拿科", "盐田院区 > 急诊科盐田": "急诊科盐田",
                     "盐田院区 > 检验科盐田": "检验科盐田", "盐田院区 > 妇产科盐田": "妇产科盐田",
                     "本部院区 > 消化内科": "消化内科", "本部院区 > 肾内科": "肾内科", "本部院区 > 病理科": "病理科",
                     "本部院区 > 神经内科": "神经内科", "本部院区 > 普外科": "普外科", "盐田院区 > 内分泌科": "内分泌科",
                     "本部院区 > 麻醉科": "麻醉科", "本部院区 > 口腔科": "口腔科", "盐田院区 > 中西医结合肛肠科": "中西医结合肛肠科",
                     "本部院区 > 传染科": "传染科", "本部院区 > 皮肤科": "皮肤科", "本部院区 > 中西医结合心血管内科": "中西医结合心血管内科",
                     "盐田院区 > 中西医结合老年病科": "中西医结合老年病科", "甲状腺乳腺外科": "甲乳外科",
                     "本部院区 > 放射影像科": "放射影像科", "本部院区 > 儿科": "儿科", "本部院区 > 耳鼻喉科": "耳鼻喉科",
                     "本部院区 > 超声影像科": "超声影像科", "本部院区 > 产科": "产科", "盐田院区 > 康复医学科盐田": "康复医学科盐田",
                     "本部院区 > 体检科": "体检科","本部院区 > 骨科":"骨科","本部院区 > 急诊科":"急诊科"
                     }
        for r in range(3, ws.max_row + 1):
            if ws.cell(r, 3).value:
                deptname = DeptNames[ws.cell(r, 3).value]
                QCCScore = ws.cell(r, 6).value
                print("Processing %s data" % deptname)
                Qualitydataexist = Quality.query.filter_by(deptname=deptname, date=date).first()
                if Qualitydataexist:
                    Qualitydataexist.QCCScore = QCCScore
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = Quality(deptname=deptname, date=date, QCCScore=QCCScore)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
        wb.close()
        flash('质量管理小组评分数据导入完毕!')
    # 院感药剂数据导入
    elif processform.is_submitted() and functype == '7':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # 院感药剂数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        # date = datetime.datetime.strptime(date, '%Y-%m-%d')
        wb = load_workbook(file, read_only=True)
        ws = wb['院感']
        for r in range(2, ws.max_row + 1):
            deptname = ws.cell(r, 1).value
            ScoreEDM = ws.cell(r, 3).value
            ScoreHandHygi = ws.cell(r, 4).value
            Qualitydataexist = Quality.query.filter_by(deptname=deptname, date=date).first()
            if Qualitydataexist:
                Qualitydataexist.ScoreEDM = ScoreEDM
                Qualitydataexist.ScoreHandHygi = ScoreHandHygi
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = Quality(deptname=deptname, date=date, ScoreEDM=ScoreEDM, ScoreHandHygi=ScoreHandHygi)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
        ws = wb['药剂']
        for r in range(2, ws.max_row + 1):
            deptname = ws.cell(r, 1).value
            NumRpRev = ws.cell(r, 3).value
            NumQuaRp = ws.cell(r, 4).value
            NumDrugConsultation = ws.cell(r, 7).value
            NumPharmCliRd = ws.cell(r, 6).value
            PharmacyDataexist = PharmacyData.query.filter_by(deptname=deptname, date=date).first()
            if PharmacyDataexist:
                PharmacyDataexist.NumRpRev = NumRpRev
                PharmacyDataexist.NumQuaRp = NumQuaRp
                PharmacyDataexist.NumDrugConsultation = NumDrugConsultation
                PharmacyDataexist.NumPharmCliRd = NumPharmCliRd
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = PharmacyData(deptname=deptname, date=date, NumRpRev=NumRpRev, NumQuaRp=NumQuaRp,
                                    NumDrugConsultation=NumDrugConsultation, NumPharmCliRd=NumPharmCliRd)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
        wb.close()
        flash('院感药剂数据导入完毕!')
    # 病案系统数据导入
    elif processform.is_submitted() and functype == '8':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # 病案数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        wb = load_workbook(file, read_only=True)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            deptname = ws.cell(r, 2).value
            IPNumInDist = ws.cell(r, 7).value
            NumDie = ws.cell(r, 36).value
            NumSurgComp = ws.cell(r, 4).value
            NumType1Surg = ws.cell(r, 5).value
            NumType1Infect = ws.cell(r, 6).value
            Qualitydataexist = Quality.query.filter_by(deptname=deptname, date=date).first()
            Surgerydataexist = Surgery.query.filter_by(deptname=deptname, date=date).first()
            InPatientdataexist = InPatient.query.filter_by(deptname=deptname, date=date).first()
            if Qualitydataexist:
                Qualitydataexist.NumDie = NumDie
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = Quality(deptname=deptname, date=date, NumDie=NumDie)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            surgdepts = ['产科', '甲乳外科', '口腔科', '麻醉科', '泌尿外科',
                         '普外科', '神经内科', '神经外科', '普外科盐田',
                         '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田',
                         '胸外科', '眼科', '中西医结合肛肠科', '中西医结合心血管内科']
            if deptname in surgdepts:
                if Surgerydataexist:
                    Surgerydataexist.NumSurgComp = NumSurgComp
                    Surgerydataexist.NumType1Surg = NumType1Surg
                    Surgerydataexist.NumType1Infect = NumType1Infect
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = Surgery(deptname=deptname, date=date, NumSurgComp=NumSurgComp,
                                   NumType1Surg=NumType1Surg, NumType1Infect=NumType1Infect)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
            ipdepts = ['产科', '甲乳外科', '康复医学科', '泌尿外科', '普外科盐田', '康复医学科盐田',
                       '内分泌科', '普外科', '神经内科', '神经外科', '肾内科', '消化内科', '新生儿科',
                       '传染科', '儿科', '耳鼻喉科', '妇产科盐田', '妇科', '骨科', '骨科盐田', '呼吸内科',
                       '胸外科', '血液内科', '眼科', '中西医结合肛肠科', '中西医结合老年病科',
                       '中西医结合心血管内科', '重症医学科', '肿瘤科']
            if deptname in ipdepts:
                if InPatientdataexist:
                    InPatientdataexist.IPNumInDist = IPNumInDist
                    try:
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
                else:
                    data = InPatient(deptname=deptname, date=date, IPNumInDist=IPNumInDist)
                    try:
                        db.session.add(data)
                        db.session.commit()
                    except Exception as e:
                        db.session.rollback()
        flash('导入病案系统数据完毕!')
        wb.close()
    # DRG数据导入
    elif processform.is_submitted() and functype == '9':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # DRG数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        wb = load_workbook(file)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            DRGGrp = int(ws.cell(r, 2).value)
            CMI = float(ws.cell(r, 5).value)
            deptname = ws.cell(r, 1).value
            print('%s 的 DRG组数是 %d CMI是 %.2f' % (deptname, DRGGrp, float(CMI)))
            InPatientexist = InPatient.query.filter_by(deptname=deptname, date=date).first()
            if InPatientexist:
                InPatientexist.DRGGrp = DRGGrp
                InPatientexist.CMI = CMI
                try:
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
            else:
                data = InPatient(deptname=deptname, date=date, DRGGrp=DRGGrp, CMI=CMI)
                try:
                    db.session.add(data)
                    db.session.commit()
                except Exception as e:
                    db.session.rollback()
        wb.close()
        flash('DRG数据导入完毕!')
    # 处理医技数据导入
    elif processform.is_submitted() and functype == '10':
        # 文件上传函数
        file = processform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        # 处理医技数据导入数据库
        date = processform.date_sel.data
        file = os.path.join('./soucedata/', file.filename)
        wb = load_workbook(file)
        ws_fsk1 = wb['放射科（本部）']
        ws_fsk2 = wb['放射科（盐田院区）']
        ws_jyk1 = wb['检验科（本部）']
        ws_jyk2 = wb['检验科（盐田院区）']
        ws_blk = wb['病理科']
        ws_csk1 = wb['超声影像科(本部）']
        ws_csk2 = wb['超声影像科(盐田）']
        # 处理放射科数据
        fsk_time = 0
        for r in range(3, ws_fsk1.max_row + 1):
            ws_fsk1.cell(r, 9).value = 0 if ws_fsk1.cell(r, 9).value is None else ws_fsk1.cell(r, 9).value
            fsk_time += int(ws_fsk1.cell(r, 9).value)
        for r in range(3, ws_fsk2.max_row + 1):
            if ws_fsk2.cell(r, 9).value:
                fsk_time += int(ws_fsk2.cell(r, 9).value)
            else:
                fsk_time += 0
        print(fsk_time)

        # 处理检验科数据
        jyk_time = 0
        for c in range(1, 12):
            if ws_jyk1.cell(2, c).value == '使用时长（小时）':
                c_right = c
        for r in range(3, ws_jyk1.max_row + 1):
            if ws_jyk1.cell(r, c_right).value:
                jyk_time += int(ws_jyk1.cell(r, c_right).value)
            else:
                jyk_time += 0
        for r in range(3, ws_jyk2.max_row + 1):
            if ws_jyk2.cell(r, 9).value:
                jyk_time += int(ws_jyk2.cell(r, 9).value)
            else:
                jyk_time += 0
        print(jyk_time)

        # 处理病理科数据
        blk_time = 0
        for c in range(1, 12):
            if ws_blk.cell(1, c).value == '使用时长（小时）':
                c_right = c
        for r in range(2, ws_blk.max_row + 1):
            if ws_blk.cell(r, c_right).value:
                blk_time += int(ws_blk.cell(r, c_right).value)
            else:
                blk_time += 0
        print(blk_time)

        # 处理超声科数据
        csk_time = 0
        for c in range(1, 12):
            if ws_csk1.cell(2, c).value == '使用时长（小时）':
                c_right = c
        for r in range(3, ws_csk1.max_row + 1):
            if ws_csk1.cell(r, c_right).value:
                csk_time += int(ws_csk1.cell(r, c_right).value)
            else:
                csk_time += 0
        for r in range(3, ws_csk2.max_row + 1):
            if ws_csk2.cell(r, 9).value:
                csk_time += int(ws_csk2.cell(r, 9).value)
            else:
                csk_time += 0
        print(csk_time)

        # 写入数据库
        # for deptname in ['放射影像科','检验科','病理科','超声影像科']:
        data1 = ExamTestData(deptname='放射影像科', date=date, MedEquipTime=fsk_time, MedEquipMalFunc=0,
                             RateIQC=random.uniform(0.96, 0.99), NumPassEQA=0,
                             AccuRateInspRept=random.uniform(0.97, 0.99),
                             TimelyRateInspRept=random.uniform(0.97, 0.99), NumForumwithClin=0)
        data2 = ExamTestData(deptname='检验科', date=date, MedEquipTime=jyk_time, MedEquipMalFunc=0,
                             RateIQC=random.uniform(0.96, 0.99), NumPassEQA=0,
                             AccuRateInspRept=random.uniform(0.97, 0.99),
                             TimelyRateInspRept=random.uniform(0.97, 0.99), NumForumwithClin=0)
        data3 = ExamTestData(deptname='病理科', date=date, MedEquipTime=blk_time, MedEquipMalFunc=0,
                             RateIQC=random.uniform(0.96, 0.99), NumPassEQA=0,
                             AccuRateInspRept=random.uniform(0.97, 0.99),
                             TimelyRateInspRept=random.uniform(0.97, 0.99), NumForumwithClin=0)
        data4 = ExamTestData(deptname='超声影像科', date=date, MedEquipTime=csk_time, MedEquipMalFunc=0,
                             RateIQC=random.uniform(0.96, 0.99), NumPassEQA=0,
                             AccuRateInspRept=random.uniform(0.97, 0.99),
                             TimelyRateInspRept=random.uniform(0.97, 0.99), NumForumwithClin=0)
        # 生成麻醉科数据
        data5 = Anestdata(deptname='麻醉科', date=date, ratepreaneseval=random.uniform(0.96, 0.99),
                          ratesurgsafeveri=random.uniform(0.97, 0.99),
                          rateanesresu=random.uniform(0.97, 0.99),
                          analysisanescomp=random.uniform(0.97, 0.99),
                          ratepostsurgvisit=random.uniform(0.97, 0.99), )
        try:
            db.session.add_all([data1, data2, data3, data4, data5])
            db.session.commit()
            print('医技数据处理成功！')
        except Exception as e:
            db.session.rollback()
            print('wrong')
        wb.close()
    # 所有空数据处理
    elif processform.is_submitted() and functype == '11':
        Income.query.filter(Income.drugincome.is_(None)).update({'drugincome': 0})
        Income.query.filter(Income.adjDrugincome.is_(None)).update({'adjDrugincome': 0})
        Income.query.filter(Income.consumIncome.is_(None)).update({'consumIncome': 0})
        Income.query.filter(Income.examIncome.is_(None)).update({'examIncome': 0})
        Income.query.filter(Income.testIncome.is_(None)).update({'testIncome': 0})
        Income.query.filter(Income.oPIncome.is_(None)).update({'oPIncome': 0})
        Income.query.filter(Income.iPIncome.is_(None)).update({'iPIncome': 0})
        Income.query.filter(Income.pureIncome.is_(None)).update({'pureIncome': 0})

        InPatient.query.filter(InPatient.NumIP.is_(None)).update({'NumIP': 0})
        InPatient.query.filter(InPatient.NumBD.is_(None)).update({'NumBD': 0})
        InPatient.query.filter(InPatient.NumRefIP.is_(None)).update({'NumRefIP': 0})
        InPatient.query.filter(InPatient.CheckNumIPRec.is_(None)).update({'CheckNumIPRec': 0})
        InPatient.query.filter(InPatient.ANumIPRec.is_(None)).update({'ANumIPRec': 0})
        InPatient.query.filter(InPatient.BNumIPRec.is_(None)).update({'BNumIPRec': 0})
        InPatient.query.filter(InPatient.CNumIPRec.is_(None)).update({'CNumIPRec': 0})
        InPatient.query.filter(InPatient.OANumIPRec.is_(None)).update({'OANumIPRec': 0})
        InPatient.query.filter(InPatient.DRGGrp.is_(None)).update({'DRGGrp': 0})
        InPatient.query.filter(InPatient.CMI.is_(None)).update({'CMI': 0})
        InPatient.query.filter(InPatient.NumPath.is_(None)).update({'NumPath': 0})
        InPatient.query.filter(InPatient.IPNumInDist.is_(None)).update({'IPNumInDist': 0})

        OutPatient.query.filter(OutPatient.NumOP.is_(None)).update({'NumOP': 0})
        OutPatient.query.filter(OutPatient.NumRefOP.is_(None)).update({'NumRefOP': 0})
        OutPatient.query.filter(OutPatient.NumOPRec.is_(None)).update({'NumOPRec': 0})
        OutPatient.query.filter(OutPatient.CheckNumOPRec.is_(None)).update({'CheckNumOPRec': 0})
        OutPatient.query.filter(OutPatient.ANumOPRec.is_(None)).update({'ANumOPRec': 0})
        OutPatient.query.filter(OutPatient.BNumOPRec.is_(None)).update({'BNumOPRec': 0})
        OutPatient.query.filter(OutPatient.CNumOPRec.is_(None)).update({'CNumOPRec': 0})
        OutPatient.query.filter(OutPatient.NumAppt.is_(None)).update({'NumAppt': 0})
        OutPatient.query.filter(OutPatient.WaitTimeAppt.is_(None)).update({'WaitTimeAppt': 0})
        OutPatient.query.filter(OutPatient.OPNumInDist.is_(None)).update({'OPNumInDist': 0})

        PharmacyData.query.filter(PharmacyData.NumRp.is_(None)).update({'NumRp': 0})
        PharmacyData.query.filter(PharmacyData.NumRpRev.is_(None)).update({'NumRpRev': 0})
        PharmacyData.query.filter(PharmacyData.NumQuaRp.is_(None)).update({'NumQuaRp': 0})
        PharmacyData.query.filter(PharmacyData.NumEssDrugOP.is_(None)).update({'NumEssDrugOP': 0})
        PharmacyData.query.filter(PharmacyData.NumEssDrugIP.is_(None)).update({'NumEssDrugIP': 0})
        PharmacyData.query.filter(PharmacyData.DDD.is_(None)).update({'DDD': 0})
        PharmacyData.query.filter(PharmacyData.NumDrugConsultation.is_(None)).update({'NumDrugConsultation': 0})
        PharmacyData.query.filter(PharmacyData.NumPharmCliRd.is_(None)).update({'NumPharmCliRd': 0})

        Quality.query.filter(Quality.NumDie.is_(None)).update({'NumDie': 0})
        Quality.query.filter(Quality.QCCScore.is_(None)).update({'QCCScore': 0})
        Quality.query.filter(Quality.CriValRepNum.is_(None)).update({'CriValRepNum': 0})
        Quality.query.filter(Quality.CriValHPNum.is_(None)).update({'CriValHPNum': 0})
        Quality.query.filter(Quality.NumAdvEvtRep.is_(None)).update({'NumAdvEvtRep': 0})
        Quality.query.filter(Quality.NumCTMri.is_(None)).update({'NumCTMri': 0})
        Quality.query.filter(Quality.NumCTMriPositive.is_(None)).update({'NumCTMriPositive': 0})
        Quality.query.filter(Quality.ScoreEDM.is_(None)).update({'ScoreEDM': 0})
        Quality.query.filter(Quality.ScoreHandHygi.is_(None)).update({'ScoreHandHygi': 0})
        Quality.query.filter(Quality.NumVioCoreSys.is_(None)).update({'NumVioCoreSys': 0})

        Surgery.query.filter(Surgery.NumSurg.is_(None)).update({'NumSurg': 0})
        Surgery.query.filter(Surgery.NumEleSurg.is_(None)).update({'NumEleSurg': 0})
        Surgery.query.filter(Surgery.NumDaySurg.is_(None)).update({'NumDaySurg': 0})
        Surgery.query.filter(Surgery.NumMiroInvaSurg.is_(None)).update({'NumMiroInvaSurg': 0})
        Surgery.query.filter(Surgery.Num4thSurg.is_(None)).update({'Num4thSurg': 0})
        Surgery.query.filter(Surgery.NumSurgComp.is_(None)).update({'NumSurgComp': 0})
        Surgery.query.filter(Surgery.NumType1Surg.is_(None)).update({'NumType1Surg': 0})
        Surgery.query.filter(Surgery.NumType1Infect.is_(None)).update({'NumType1Infect': 0})
        try:
            db.session.commit()
        except:
            db.session.rollback()
        flash('处理完毕')


    return render_template('dataprocess.html', processform=processform, functype=functype)

if __name__ == '__main__':
    app.run()