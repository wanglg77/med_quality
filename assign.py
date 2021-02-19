from define import *

@app.route('/assign', methods=['GET','POST'])

def assign():
    assignform = Assign()
    pd.set_option('display.unicode.ambiguous_as_wide', True)
    pd.set_option('display.unicode.east_asian_width', True)
    pd.set_option('display.width', 80)
    pd.options.display.max_columns = 22

    if assignform.is_submitted():
        # 文件上传函数
        file = assignform.file_upload.data
        file.save(os.path.join('./soucedata/', file.filename))
        date = assignform.date_sel.data
        year = date.year
        month = date.month
        file = os.path.join('./soucedata/', file.filename)
        # 从病例列表中提取符合规则的病例放入rec_checkist列表,每个科室10份
        rec_list = pd.read_excel(file,engine='openpyxl')
        depts = rec_list['出院科别'].unique()
        rec_checklist = []
        for dept in depts:
            dept_rec_death = rec_list[(rec_list['出院科别'] == dept) & (rec_list['离院方式'] == '死亡')]
            dept_rec_dtype = rec_list[
                (rec_list['出院科别'] == dept) & (rec_list['病例分型'] == 'D') & (rec_list['离院方式'] != '死亡')]
            dept_rec_ctype = rec_list[
                (rec_list['出院科别'] == dept) & (rec_list['病例分型'] == 'C') & (rec_list['离院方式'] != '死亡')]
            dept_rec_rest = rec_list[(rec_list['出院科别'] == dept) & (rec_list['离院方式'] != '死亡') & (rec_list['病例分型'] != 'D') & (rec_list['病例分型'] != 'C')]

            if len(dept_rec_death.values) >= 10:
                dept_checklist = dept_rec_death[['出院科别', '病案号', '姓名']].sample(n=10).values.tolist()
                print(dept,'死亡够10', dept_checklist)
            else:
                num_add = 10 - len(dept_rec_death.values)
                if num_add < len(dept_rec_dtype[['出院科别', '病案号', '姓名']].values):
                    dept_checklist = dept_rec_death[['出院科别', '病案号', '姓名']].values.tolist() + \
                                     dept_rec_dtype[['出院科别', '病案号', '姓名']].sample(n=num_add).values.tolist()
                    print(dept,'死亡加D型够10', len(dept_checklist))
                else:
                    dept_checklist = dept_rec_death[['出院科别', '病案号', '姓名']].values.tolist() + \
                                     dept_rec_dtype[['出院科别', '病案号', '姓名']].values.tolist()
                    num_add = 10 - len(dept_checklist)
                    if num_add < len(dept_rec_ctype[['出院科别', '病案号', '姓名']].values.tolist()):
                        dept_checklist += dept_rec_ctype[['出院科别', '病案号', '姓名']].sample(n=num_add).values.tolist()
                        print(dept,'死亡加D型加C型够10', len(dept_checklist))
                    else:
                        dept_checklist = dept_rec_death[['出院科别', '病案号', '姓名']].values.tolist() + \
                                         dept_rec_dtype[['出院科别', '病案号', '姓名']].values.tolist() + \
                                         dept_rec_ctype[['出院科别', '病案号', '姓名']].values.tolist()
                        num_add = 10-len(dept_checklist)
                        if num_add < len(dept_rec_rest[['出院科别', '病案号', '姓名']].values.tolist()):
                            dept_checklist += dept_rec_rest[['出院科别', '病案号', '姓名']].sample(n=num_add).values.tolist()
                        else:
                            dept_checklist += dept_rec_rest[['出院科别', '病案号', '姓名']].values.tolist()
                        print(dept, '死亡加D型加C型加其他够10', len(dept_checklist))

            rec_checklist += dept_checklist

        # 将列表中的病例分配给科主任和质控员
        wb = openpyxl.Workbook()
        ws = wb.create_sheet('科主任')
        ws2 = wb.create_sheet('质控员')
        chiefs = {'付小义': '传染科', '陈彧': '眼科', '宋瑾': '耳鼻喉科', '唐红': '儿科',
                  '刘庆仪': '甲乳外科', '吴正国': '胸外科', '温贺龙': '骨科', '潘志铣': '肾内科',
                  '李显文': '泌尿外科', '刘其强': '康复医学科', '肖柏成': '神经内科',
                  '晏斌林': '呼吸内科', '冯清洲': '重症医学科', '肖虎': '内分泌科',
                  '柳金': '血液内科', '袁克华': '肿瘤科', '巫晓蓉': '营养科', '郑丽芳': '神经内科',
                  '李林娜': '产科', '涂文斌': '康复医学科', '许文顺': '普外科盐田', '钟昌戎': '骨科盐田',
                  '马华': '中西医结合心血管内科', '党登峰': '普外科', '郑志刚': '中西医结合老年病科', '鞠大闯': '中西医结合肛肠科'}
        # 科主任分配
        chiefs_list = list(chiefs.keys())
        ws.cell(1, 1, '科主任病历质控分配表')
        titles = ['检查者', '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名',
                  '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名']
        for c in range(1, len(titles) + 1):
            ws.cell(2, c, titles[c - 1])
        for r in range(3, 27):
            ws.cell(r, 1, chiefs_list[r - 2])
            rec_chief_checklist = random.sample(rec_checklist, 5)
            for c in range(2, 17):
                if (c - 1) % 3 == 1:
                    ws.cell(r, c, rec_chief_checklist[int((c + 1) / 3 - 1)][0])
                elif (c - 1) % 3 == 2:
                    ws.cell(r, c, rec_chief_checklist[int((c + 1) / 3 - 1)][1])
                elif (c - 1) % 3 == 0:
                    ws.cell(r, c, rec_chief_checklist[int((c + 1) / 3 - 1)][2])
            for i in range(0, 5):
                rec_checklist.remove(rec_chief_checklist[i])

        # 质控员分配
        qca_list = ['徐晓婧', '郑晓丽', '钟木生', '陈宝洁', '杨逢生', '黄金河', '杜娟', '王敏', '温展翀',
                    '钟瑜', '董航', '张绍敏', '王珍妮', '刘昭红', '左旋', '古晓珊', '彭守仙', '赵飞',
                    '邹云东', '梁比记', '汪谢嫒', '张雷', '沈晓燕', '陈娟']
        qca_out_list = ['曹丽军', '王燕', '张伟艺', '李龙振', '韦伟', '单菲', '廖伟东', '冯晓剑']
        dept_list = ['产科', '耳鼻喉科', '眼科', '妇科', '骨科', '甲乳外科', '神经外科', '胸外科', '普外科',
                     '神经内科', '传染科', '康复医学科', '呼吸内科', '肾内科', '中西医结合心血管内科', '重症医学科', '消化内科', '肿瘤科',
                     '血液内科', '儿科', '精神科', '中医科', '皮肤科', '口腔科']
        dept_out_list = ['针灸推拿科', '急诊科', '普外科盐田', '骨科盐田', '中西医结合肛肠科', '中西医结合老年病科',
                         '泌尿外科盐田', '妇产科盐田', '内分泌科', '儿科盐田', '耳鼻喉科盐田', '急诊科盐田', '口腔科盐田',
                         '皮肤科盐田', '眼科盐田', '康复医学科盐田', '传染科盐田', '中医科盐田']
        ws2.cell(1, 1, '质控员病历质控分配表')
        ws2.cell(1, 1).font = Font(size=16, bold=True)
        titles = ['检查者', '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名',
                  '出院科室', '住院号', '患者姓名', '出院科室', '住院号', '患者姓名', '门诊每科10份']
        for c in range(1, len(titles) + 1):
            ws2.cell(2, c, titles[c - 1])
            ws2.cell(2, c).font = Font(bold=True)
        for r in range(3, 27):
            ws2.cell(r, 1, qca_list[r - 3])
            ws2.cell(r, 17, dept_list[r - 3])
            rec_qca_checklist = random.sample(rec_checklist, 5)
            for c in range(2, 17):
                if (c - 1) % 3 == 1:
                    ws2.cell(r, c, rec_qca_checklist[int((c + 1) / 3 - 1)][0])
                elif (c - 1) % 3 == 2:
                    ws2.cell(r, c, rec_qca_checklist[int((c + 1) / 3 - 1)][1])
                elif (c - 1) % 3 == 0:
                    ws2.cell(r, c, rec_qca_checklist[int((c + 1) / 3 - 1)][2])
            for i in range(0, 5):
                rec_checklist.remove(rec_qca_checklist[i])
        # 分配剩余门诊病历
        ws2.cell(27, 1, '曹丽军')
        ws2.cell(27, 17, '针灸推拿科,急诊科')
        ws2.cell(28, 1, '王燕')
        ws2.cell(28, 17, '普外科盐田,骨科盐田')
        ws2.cell(29, 1, '张伟艺')
        ws2.cell(29, 17, '中西医结合肛肠科,中西医结合老年病科')
        ws2.cell(30, 1, '李龙振')
        ws2.cell(30, 17, '泌尿外科盐田,妇产科盐田')
        ws2.cell(31, 1, '韦伟')
        ws2.cell(31, 17, '内分泌科,儿科盐田,耳鼻喉科盐田')
        ws2.cell(32, 1, '单菲')
        ws2.cell(32, 17, '急诊科盐田,口腔科盐田')
        ws2.cell(33, 1, '廖伟东')
        ws2.cell(33, 17, '皮肤科盐田,眼科盐田')
        ws2.cell(34, 1, '冯晓剑')
        ws2.cell(34, 17, '康复医学科盐田,传染科盐田,中医科盐田')

        wb.active = wb['科主任']

        print(len(rec_checklist), rec_checklist)

        wb.save('病历分配表.xlsx')

    return render_template('assign.html', assignform = assignform)

if __name__ == '__main__':
    app.run()