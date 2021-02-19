from define import *


@app.route('/query', methods=['GET','POST'])

def query():
    queryform = Queryform()
    deptclass = ''
    dataqs = []
    dataos = []
    dataps = []
    dataics = []
    datais = []
    datass = []
    dataes = []
    dataas = []
    scores = []
    if queryform.submit.data and queryform.is_submitted():
        deptclass = queryform.deptname.data
        start = queryform.startdate.data
        over = queryform.overdate.data
        if start.month != 1:
            lastmonth = start.replace(month=start.month-1)
        else:
            lastmonth = start.replace(year=start.year-1,month=12)
        depts = db.session.query(Department.name).filter(Department.classname == deptclass).all()

        if deptclass == '手术':
            for d in depts:
                dataq = db.session.query(Quality.deptname,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.sum(Quality.NumCTMriPositive)/func.sum(Quality.NumCTMri),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Quality.deptname==d[0], Quality.date.between(start,over)).\
                                        all()
                dataqs.append(dataq[0])
                datao = db.session.query(
                    func.sum(OutPatient.NumOPRec) / func.sum(OutPatient.NumOP),
                    func.sum(OutPatient.ANumOPRec)/func.sum(OutPatient.CheckNumOPRec),
                    func.sum(OutPatient.NumAppt) / func.sum(OutPatient.NumOP),
                    func.avg(OutPatient.WaitTimeAppt),
                    1-(func.sum(OutPatient.OPNumInDist) / func.sum(OutPatient.NumOP))). \
                    filter(OutPatient.deptname == d[0], OutPatient.date.between(start, over)). \
                    all()
                dataos.append(datao[0])
                datap = db.session.query(
                    func.sum(PharmacyData.NumQuaRp) / func.sum(PharmacyData.NumRpRev),
                    func.sum(PharmacyData.NumEssDrugOP) / func.sum(PharmacyData.NumRp)). \
                    filter(PharmacyData.deptname == d[0], PharmacyData.date.between(start, over)). \
                    all()
                dataps.append(datap[0])
                dataic = db.session.query(
                    func.sum(Income.adjDrugincome) / func.sum(Income.drugincome),
                    func.sum(Income.pureIncome) / func.sum(Income.totalincome)). \
                    filter(Income.deptname == d[0], Income.date.between(start, over)). \
                    all()
                dataics.append(dataic[0])
                datai = db.session.query(
                    func.sum(InPatient.NumIP),
                    func.sum(InPatient.NumBD),
                    func.sum(InPatient.ANumIPRec) / func.sum(InPatient.CheckNumIPRec),
                    1-(func.sum(InPatient.OANumIPRec) / func.sum(InPatient.NumIP)),
                    1-(func.sum(InPatient.IPNumInDist) / func.sum(InPatient.NumIP)),
                    func.sum(InPatient.NumPath) / func.sum(InPatient.NumIP),
                    func.avg(InPatient.DRGGrp),
                    func.avg(InPatient.CMI)). \
                    filter(InPatient.deptname == d[0], InPatient.date.between(start, over)). \
                    all()
                datais.append(datai[0])
                datas = db.session.query(
                    func.sum(Surgery.NumSurg),
                    func.sum(Surgery.NumEleSurg),
                    func.sum(Surgery.NumDaySurg)/func.sum(Surgery.NumEleSurg),
                    func.sum(Surgery.NumMiroInvaSurg)/func.sum(Surgery.NumSurg),
                    func.sum(Surgery.Num4thSurg)/func.sum(Surgery.NumSurg),
                    func.sum(Surgery.NumSurgComp)/func.sum(Surgery.NumEleSurg),
                    func.sum(Surgery.NumType1Infect)/func.sum(Surgery.NumType1Surg)
                    ).\
                    filter(Surgery.deptname == d[0], Surgery.date.between(start, over)). \
                    all()
                datass.append(datas[0])
                if d[0] == '产科':
                    score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                            float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                            float(0 if dataq[0][3] is None else dataq[0][3]) * 5 + \
                            float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                            float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                            float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                            float(0 if dataq[0][7] is None else dataq[0][7]) * 10 + \
                            float(0 if datao[0][0] is None else datao[0][0]) * 0 + \
                            float(0 if datao[0][1] is None else datao[0][1]) * 20 + \
                            float(0 if datao[0][2] is None else datao[0][2]) * 10 + \
                            float(0 if datao[0][4] is None else datao[0][4]) * 10 + \
                            float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                            float(0 if datap[0][1] is None else datap[0][1]) * 10 + \
                            (1 - float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                            float(0 if dataic[0][1] is None else dataic[0][1]) * 10 + \
                            float(0 if datai[0][2] is None else datai[0][2]) * 10 + \
                            float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                            float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                            float(0 if datai[0][4] is None else datai[0][4]) * 10 + \
                            float(0 if datai[0][5] is None else datai[0][5]) * 10 + \
                            min((float(0 if datai[0][6] is None else datai[0][6]) * 1), 10) + \
                            min((float(0 if datai[0][7] is None else datai[0][7]) * 50), 10) + \
                            float(0 if datas[0][0] is None else datas[0][0]) / float(
                        0 if datai[0][0] is None else datai[0][0]) * 10 + \
                            float(0 if datas[0][2] is None else datas[0][2]) * 10 + \
                            float(0 if datas[0][3] is None else datas[0][3]) * 10 + \
                            float(0 if datas[0][4] is None else datas[0][4]) * 10 + \
                            (1 - float(0 if datas[0][5] is None else datas[0][5])) * 10 + \
                            (1 - float(0 if datas[0][6] is None else datas[0][6])) * 10
                else:
                    score = float(0 if dataq[0][1] is None else dataq[0][1])*0.1+\
                        float(1 if dataq[0][2] is None else dataq[0][2])*10+ \
                        float(0 if dataq[0][3] is None else dataq[0][3])*5+ \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                        float(0 if dataq[0][7] is None else dataq[0][7]) * 10+ \
                        float(0 if datao[0][0] is None else datao[0][0]) * 10 + \
                        float(0 if datao[0][1] is None else datao[0][1]) * 10 + \
                        float(0 if datao[0][2] is None else datao[0][2]) * 10 + \
                        float(0 if datao[0][4] is None else datao[0][4]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                        float(0 if datap[0][1] is None else datap[0][1]) * 10 + \
                        (1-float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                        float(0 if dataic[0][1] is None else dataic[0][1]) * 10 + \
                        float(0 if datai[0][2] is None else datai[0][2]) * 10 + \
                        float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                        float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                        float(0 if datai[0][4] is None else datai[0][4]) * 10 + \
                        float(0 if datai[0][5] is None else datai[0][5]) * 10 + \
                        min((float(0 if datai[0][6] is None else datai[0][6]) * 1),10) + \
                        min((float(0 if datai[0][7] is None else datai[0][7]) * 50),10) + \
                        float(0 if datas[0][0] is None else datas[0][0])/float(0 if datai[0][0] is None else datai[0][0]) * 10 + \
                        float(0 if datas[0][2] is None else datas[0][2]) * 10 + \
                        float(0 if datas[0][3] is None else datas[0][3]) * 10 + \
                        float(0 if datas[0][4] is None else datas[0][4]) * 10 + \
                        (1-float(0 if datas[0][5] is None else datas[0][5])) * 10 + \
                        (1 - float(0 if datas[0][6] is None else datas[0][6])) * 10
                scores.append(score/2.1)
                print(d,score)

        elif deptclass == '住院':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.sum(Quality.NumCTMriPositive)/func.sum(Quality.NumCTMri),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                datao = db.session.query(
                    func.sum(OutPatient.NumOPRec) / func.sum(OutPatient.NumOP),
                    func.sum(OutPatient.ANumOPRec)/func.sum(OutPatient.CheckNumOPRec),
                    func.sum(OutPatient.NumAppt) / func.sum(OutPatient.NumOP),
                    func.avg(OutPatient.WaitTimeAppt),
                    1-(func.sum(OutPatient.OPNumInDist) / func.sum(OutPatient.NumOP))). \
                    filter(OutPatient.deptname == d[0], OutPatient.date.between(start, over)). \
                    all()
                dataos.append(datao[0])
                datap = db.session.query(
                    func.sum(PharmacyData.NumQuaRp) / func.sum(PharmacyData.NumRpRev),
                    func.sum(PharmacyData.NumEssDrugOP) / func.sum(PharmacyData.NumRp)). \
                    filter(PharmacyData.deptname == d[0], PharmacyData.date.between(start, over)). \
                    all()
                dataps.append(datap[0])
                dataic = db.session.query(
                    func.sum(Income.adjDrugincome) / func.sum(Income.drugincome),
                    func.sum(Income.pureIncome) / func.sum(Income.totalincome)). \
                    filter(Income.deptname == d[0], Income.date.between(start, over)). \
                    all()
                dataics.append(dataic[0])
                datai = db.session.query(
                    func.sum(InPatient.NumIP),
                    func.sum(InPatient.NumBD),
                    func.sum(InPatient.ANumIPRec) / func.sum(InPatient.CheckNumIPRec),
                    1-(func.sum(InPatient.OANumIPRec) / func.sum(InPatient.NumIP)),
                    1-(func.sum(InPatient.IPNumInDist) / func.sum(InPatient.NumIP)),
                    func.sum(InPatient.NumPath) / func.sum(InPatient.NumIP),
                    func.avg(InPatient.DRGGrp),
                    func.avg(InPatient.CMI)). \
                    filter(InPatient.deptname == d[0], InPatient.date.between(start, over)). \
                    all()
                datais.append(datai[0])
                score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        min((float(0 if dataq[0][3] is None else dataq[0][3]) * 5),10) + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                        float(0 if dataq[0][7] is None else dataq[0][7]) * 10 + \
                        float(1 if datao[0][0] is None else datao[0][0]) * 10 + \
                        float(1 if datao[0][1] is None else datao[0][1]) * 10 + \
                        float(1 if datao[0][2] is None else datao[0][2]) * 10 + \
                        float(0 if datao[0][4] is None else datao[0][4]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                        float(0 if datap[0][1] is None else datap[0][1]) * 10 + \
                        (1 - float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                        float(0 if dataic[0][1] is None else dataic[0][1]) * 10 + \
                        float(0 if datai[0][2] is None else datai[0][2]) * 10 + \
                        float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                        float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                        float(0 if datai[0][4] is None else datai[0][4]) * 10 + \
                        float(0 if datai[0][5] is None else datai[0][5]) * 10 + \
                        min((float(0 if datai[0][6] is None else datai[0][6]) * 1), 10) + \
                        min((float(0 if datai[0][7] is None else datai[0][7]) * 50), 10)

                scores.append(score / 1.8)

        elif deptclass == '门诊':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.sum(Quality.NumCTMriPositive)/func.sum(Quality.NumCTMri),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                datao = db.session.query(
                    func.sum(OutPatient.NumOPRec) / func.sum(OutPatient.NumOP),
                    func.sum(OutPatient.ANumOPRec)/func.sum(OutPatient.CheckNumOPRec),
                    func.sum(OutPatient.NumAppt) / func.sum(OutPatient.NumOP),
                    func.avg(OutPatient.WaitTimeAppt),
                    1-(func.sum(OutPatient.OPNumInDist) / func.sum(OutPatient.NumOP))). \
                    filter(OutPatient.deptname == d[0], OutPatient.date.between(start, over)). \
                    all()
                dataos.append(datao[0])
                datap = db.session.query(
                    func.sum(PharmacyData.NumQuaRp) / func.sum(PharmacyData.NumRpRev),
                    func.sum(PharmacyData.NumEssDrugOP) / func.sum(PharmacyData.NumRp)). \
                    filter(PharmacyData.deptname == d[0], PharmacyData.date.between(start, over)). \
                    all()
                dataps.append(datap[0])
                dataic = db.session.query(
                    func.sum(Income.adjDrugincome) / func.sum(Income.drugincome),
                    func.sum(Income.pureIncome) / func.sum(Income.totalincome)). \
                    filter(Income.deptname == d[0], Income.date.between(start, over)). \
                    all()
                dataics.append(dataic[0])
                if d[0] == '精神科':
                    score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        float(0 if dataq[0][3] is None else dataq[0][3]) * 5 + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                        float(0 if dataq[0][7] is None else dataq[0][7]) * 10 + \
                        float(1 if datao[0][0] is None else datao[0][0]) * 10 + \
                        float(1 if datao[0][1] is None else datao[0][1]) * 10 + \
                        float(1 if datao[0][2] is None else datao[0][2]) * 20 + \
                        float(0 if datao[0][4] is None else datao[0][4]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                        float(0 if datap[0][1] is None else datap[0][1]) * 10 + \
                        (1 - float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                        float(0 if dataic[0][1] is None else dataic[0][1]) * 5 + 15
                else:
                    score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        float(0 if dataq[0][3] is None else dataq[0][3]) * 5 + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                        float(0 if dataq[0][7] is None else dataq[0][7]) * 10 + \
                        float(1 if datao[0][0] is None else datao[0][0]) * 10 + \
                        float(1 if datao[0][1] is None else datao[0][1]) * 10 + \
                        float(1 if datao[0][2] is None else datao[0][2]) * 10 + \
                        float(0 if datao[0][4] is None else datao[0][4]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                        float(0 if datap[0][1] is None else datap[0][1]) * 10 + \
                        (1 - float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                        float(0 if dataic[0][1] is None else dataic[0][1]) * 5
                scores.append(score / 1.1)

        elif deptclass == '医技':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                datae = db.session.query(
                                    func.sum(ExamTestData.MedEquipMalFunc) / func.sum(ExamTestData.MedEquipTime),
                                    func.avg(ExamTestData.RateIQC),
                                    func.sum(ExamTestData.NumPassEQA),
                                    func.avg(ExamTestData.AccuRateInspRept),
                                    func.avg(ExamTestData.TimelyRateInspRept),
                                    func.sum(ExamTestData.NumForumwithClin)). \
                                    filter(ExamTestData.deptname == d[0], ExamTestData.date.between(start, over)).all()
                dataes.append(datae[0])
                print(datae[0])
                score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        min((float(0 if dataq[0][3] is None else dataq[0][3]) * 5),10) + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 0.1 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 10 - \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 + \
                        float(0 if datae[0][0] is None else datae[0][0]) * 10 + \
                        float(0 if datae[0][1] is None else datae[0][1]) * 10 + \
                        float(0 if datae[0][2] is None else datae[0][2]) * 10 + \
                        float(0 if datae[0][3] is None else datae[0][3]) * 10 + \
                        float(0 if datae[0][4] is None else datae[0][4]) * 10 + \
                        float(0 if datae[0][5] is None else datae[0][5]) * 10
                scores.append(score/8*10)

        elif deptclass == '麻醉':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                dataa = db.session.query(
                    func.avg(Anestdata.ratepreaneseval),
                    func.avg(Anestdata.ratesurgsafeveri),
                    func.avg(Anestdata.rateanesresu),
                    func.sum(Anestdata.analysisanescomp),
                    func.avg(Anestdata.ratepostsurgvisit)).\
                    filter(Anestdata.deptname == d[0], Anestdata.date.between(start, over)).\
                    all()
                dataas.append(dataa[0])
                score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        min((float(0 if dataq[0][3] is None else dataq[0][3]) * 5), 10) + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 0.1 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 10 - \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 + \
                        float(0 if dataa[0][0] is None else dataa[0][0]) * 10 + \
                        float(0 if dataa[0][1] is None else dataa[0][1]) * 10 + \
                        float(0 if dataa[0][2] is None else dataa[0][2]) * 10 + \
                        float(0 if dataa[0][3] is None else dataa[0][3]) * 10 + \
                        float(0 if dataa[0][4] is None else dataa[0][4]) * 10
                scores.append( score )

        elif deptclass == '药剂':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                datap = db.session.query(
                    func.sum(PharmacyData.NumRpRev)/func.sum(PharmacyData.NumRp),
                    func.sum(PharmacyData.NumDrugConsultation),
                    func.sum(PharmacyData.NumPharmCliRd)).\
                    filter(PharmacyData.deptname == d[0], PharmacyData.date.between(start, over)).\
                    all()
                dataps.append(datap[0])
                score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        min((float(0 if dataq[0][3] is None else dataq[0][3]) * 5), 10) + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 0.1 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 10 - \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 100 + \
                        float(0 if datap[0][1] is None else datap[0][1]) * 0.1 + \
                        float(0 if datap[0][2] is None else datap[0][2]) * 1
                scores.append(score)

        elif deptclass == '重症':
            for d in depts:
                dataq = db.session.query(Department.name,
                                        func.avg(Quality.QCCScore),
                                        func.avg(Quality.CriValHPNum)/func.avg(Quality.CriValRepNum),
                                        func.sum(Quality.NumAdvEvtRep),
                                        func.sum(Quality.NumCTMriPositive)/func.sum(Quality.NumCTMri),
                                        func.avg(Quality.ScoreEDM),
                                        func.avg(Quality.ScoreHandHygi),
                                        func.sum(Quality.NumVioCoreSys)).\
                                        filter(Department.name==d[0], Quality.date.between(start,over)).\
                                        join(Department,Department.name==Quality.deptname).\
                                        all()
                dataqs.append(dataq[0])
                datap = db.session.query(
                    func.sum(PharmacyData.NumQuaRp) / func.sum(PharmacyData.NumRpRev)). \
                    filter(PharmacyData.deptname == d[0], PharmacyData.date.between(start, over)). \
                    all()
                dataps.append(datap[0])
                dataic = db.session.query(
                    func.sum(Income.adjDrugincome) / func.sum(Income.drugincome),
                    func.sum(Income.pureIncome) / func.sum(Income.totalincome)). \
                    filter(Income.deptname == d[0], Income.date.between(start, over)). \
                    all()
                dataics.append(dataic[0])
                datai = db.session.query(
                    func.sum(InPatient.NumIP),
                    func.sum(InPatient.NumBD),
                    func.sum(InPatient.ANumIPRec) / func.sum(InPatient.CheckNumIPRec),
                    1-(func.sum(InPatient.OANumIPRec) / func.sum(InPatient.NumIP)),
                    1-(func.sum(InPatient.IPNumInDist) / func.sum(InPatient.NumIP)),
                    func.avg(InPatient.DRGGrp),
                    func.avg(InPatient.CMI)). \
                    filter(InPatient.deptname == d[0], InPatient.date.between(start, over)). \
                    all()
                datais.append(datai[0])
                score = float(0 if dataq[0][1] is None else dataq[0][1]) * 0.1 + \
                        float(1 if dataq[0][2] is None else dataq[0][2]) * 10 + \
                        float(0 if dataq[0][3] is None else dataq[0][3]) * 5 + \
                        float(0 if dataq[0][4] is None else dataq[0][4]) * 10 + \
                        float(0 if dataq[0][5] is None else dataq[0][5]) * 0.1 + \
                        float(0 if dataq[0][6] is None else dataq[0][6]) * 10 - \
                        float(0 if dataq[0][7] is None else dataq[0][7]) * 10 + \
                        float(0 if datap[0][0] is None else datap[0][0]) * 10 + \
                        (1 - float(0 if dataic[0][0] is None else dataic[0][0])) * 10 + \
                        float(0 if dataic[0][1] is None else dataic[0][1]) * 10 + \
                        float(0 if datai[0][2] is None else datai[0][2]) * 10 + \
                        float(0 if datai[0][3] is None else datai[0][3]) * 10 + \
                        float(0 if datai[0][4] is None else datai[0][4]) * 10 + \
                        min((float(0 if datai[0][5] is None else datai[0][5]) * 1), 10) + \
                        min((float(0 if datai[0][6] is None else datai[0][6]) * 50), 10)

                scores.append(score / 1.2)

    if queryform.toexcel.data and queryform.is_submitted():
        start = queryform.startdate.data
        over = queryform.overdate.data
        lastyear = start.replace(start.year-1,12,31)
        depts = db.session.query(Department.name).all()
        deptssurg = db.session.query(Department.name).filter(Department.classname == '手术').all()
        deptsIP = db.session.query(Department.name).filter(Department.classname == '住院').all()
        deptsOP = db.session.query(Department.name).filter(Department.classname == '门诊').all()
        deptsExam = db.session.query(Department.name).filter(Department.classname == '医技').all()
        depts_score_list = {}  #定义所有科室分数数据字典

        # 定义查询计算门诊6个指标的得分的函数
        def queryOP(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 0
            # 5--门诊病历书写率得分
            NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept, OutPatient.date == start).first()
            NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
                                                                    OutPatient.date == start).first()
            OPRecRate_score = NumOPRec[0] / NumOP[0] * 10
            # 6--门诊病历甲级率得分
            CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
                                                                              OutPatient.date == start).first()
            ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
                                                                      OutPatient.date == start).first()
            if CheckNumOPRec[0] != 0:
                ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            else:
                ARateOPRec_score = 10
            # 计算科室总分
            total_score = qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score + OPRecRate_score + ARateOPRec_score
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score, '门诊病历书写率得分': OPRecRate_score,
                               '门诊病历甲级率得分': ARateOPRec_score, '总分': total_score}
            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算住院11个指标的得分的函数
        def queryIP(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 10
            # 5--门诊病历书写率得分
            NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept, OutPatient.date == start).first()
            NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
                                                                    OutPatient.date == start).first()
            if dept == '肾内科':
                OPRecRate_score = 9.8
            else:
                OPRecRate_score = NumOPRec[0] / NumOP[0] * 10
            # 6--门诊病历甲级率得分
            CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
                                                                              OutPatient.date == start).first()
            ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
                                                                      OutPatient.date == start).first()
            if CheckNumOPRec[0] != 0:
                ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            else:
                ARateOPRec_score = 10
            # 7--住院病历甲级率得分
            CheckNumIPRec = db.session.query(InPatient.CheckNumIPRec).filter(InPatient.deptname == dept,
                                                                              InPatient.date == start).first()
            ANumIPRec = db.session.query(InPatient.ANumIPRec).filter(InPatient.deptname == dept,
                                                                      InPatient.date == start).first()
            if CheckNumIPRec[0] != 0:
                ARateIPRec_score = ANumIPRec[0] / CheckNumIPRec[0] * 10
            else:
                ARateIPRec_score = 10
            # 8--住院病历按时归档率
            OANumIPRec = db.session.query(InPatient.OANumIPRec).filter(InPatient.deptname == dept,
                                                                             InPatient.date == start).first()
            NumIP = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                     InPatient.date == start).first()
            OANumIPRec_score = (1-OANumIPRec[0] / NumIP[0]) * 10
            # 9--临床路径入径率得分
            NumPath = db.session.query(InPatient.NumPath).filter(InPatient.deptname == dept,
                                                                       InPatient.date == start).first()
            NumPath_score = NumPath[0] / NumIP[0] * 10
            # 10--CMI得分
            CMI_Now = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            CMI_Lastyear = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                             InPatient.date == lastyear).first()
            if CMI_Lastyear is None:
                CMI_score = 10
            else:
                CMI_score = max(min((CMI_Now[0] - CMI_Lastyear[0])/CMI_Lastyear[0]*150,10),-10)

            # 11--平均住院日得分
            NumBD_Now = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            NumBD_Lastyear = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                  InPatient.date == lastyear).first()
            NumIP_Lastyear = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            if NumBD_Lastyear is None:
                AverageBedTime_score = 10
            else:
                AverageBedTime_score = max(min((NumBD_Lastyear[0]/NumIP_Lastyear[0] - NumBD_Now[0]/NumIP[0])/(NumBD_Lastyear[0] / NumIP_Lastyear[0]) * 50 , 10),-10)

            # 计算科室总分
            total_score = qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score + \
                          OPRecRate_score + ARateOPRec_score + ARateIPRec_score +\
                          OANumIPRec_score + NumPath_score + CMI_score + AverageBedTime_score
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score, '门诊病历书写率得分': OPRecRate_score,
                               '门诊病历甲级率得分': ARateOPRec_score, '住院病历甲级率':ARateIPRec_score , '住院按时归档率':OANumIPRec_score,
                               '临床路径入径率得分': NumPath_score ,'CMI得分':CMI_score ,'平均住院日得分':AverageBedTime_score , '总分': total_score}
            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算手术科室14个指标的得分的函数
        def querySurgery(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 10
            # 5--门诊病历书写率得分
            NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept, OutPatient.date == start).first()
            NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
                                                                    OutPatient.date == start).first()
            OPRecRate_score = NumOPRec[0] / NumOP[0] * 10
            # 6--门诊病历甲级率得分
            CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
                                                                              OutPatient.date == start).first()
            ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
                                                                      OutPatient.date == start).first()
            if CheckNumOPRec[0] != 0:
                ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            else:
                ARateOPRec_score = 10
            # 7--住院病历甲级率得分
            CheckNumIPRec = db.session.query(InPatient.CheckNumIPRec).filter(InPatient.deptname == dept,
                                                                              InPatient.date == start).first()
            ANumIPRec = db.session.query(InPatient.ANumIPRec).filter(InPatient.deptname == dept,
                                                                      InPatient.date == start).first()
            if CheckNumIPRec[0] != 0:
                ARateIPRec_score = ANumIPRec[0] / CheckNumIPRec[0] * 10
            else:
                ARateIPRec_score = 10
            # 8--住院病历按时归档率
            OANumIPRec = db.session.query(InPatient.OANumIPRec).filter(InPatient.deptname == dept,
                                                                             InPatient.date == start).first()
            NumIP = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                     InPatient.date == start).first()
            OANumIPRec_score = (1-OANumIPRec[0] / NumIP[0]) * 10
            # 9--临床路径入径率得分
            NumPath = db.session.query(InPatient.NumPath).filter(InPatient.deptname == dept,
                                                                       InPatient.date == start).first()
            NumPath_score = NumPath[0] / NumIP[0] * 10
            # 10--CMI得分
            CMI_Now = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            CMI_Lastyear = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                             InPatient.date == lastyear).first()
            print(dept)
            if CMI_Lastyear is None or CMI_Lastyear==0:
                CMI_score = 10
            else:
                CMI_score = max(min((CMI_Now[0] - CMI_Lastyear[0])/CMI_Lastyear[0]*150,10),-10)

            # 11--平均住院日得分
            NumBD_Now = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            NumBD_Lastyear = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                  InPatient.date == lastyear).first()
            NumIP_Lastyear = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            if NumBD_Lastyear is None or NumBD_Lastyear==0:
                AverageBedTime_score = 10
            else:
                AverageBedTime_score = max(min((NumBD_Lastyear[0]/NumIP_Lastyear[0] - NumBD_Now[0]/NumIP[0])/(NumBD_Lastyear[0] / NumIP_Lastyear[0]) * 50 , 10),-10)

            # 12--手术患者占比得分
            NumSurg = db.session.query(Surgery.NumSurg).filter(Surgery.deptname == dept,
                                                             Surgery.date == start).first()
            NumSurg_score = NumSurg[0]/NumIP[0] * 10

            # 12--微创手术占比得分
            NumMiroInvaSurg = db.session.query(Surgery.NumMiroInvaSurg).filter(Surgery.deptname == dept,
                                                             Surgery.date == start).first()
            NumEleSurg = db.session.query(Surgery.NumEleSurg).filter(Surgery.deptname == dept,
                                                                               Surgery.date == start).first()
            NumMiroInvaSurg_score = NumMiroInvaSurg[0]/NumEleSurg[0] * 10

            # 13--四级手术占比得分
            Num4thSurg = db.session.query(Surgery.Num4thSurg).filter(Surgery.deptname == dept,
                                                                               Surgery.date == start).first()
            Num4thSurg_score = Num4thSurg[0] / NumSurg[0] * 10

            # 14--手术并发症占比得分
            NumSurgComp = db.session.query(Surgery.NumSurgComp).filter(Surgery.deptname == dept,
                                                                     Surgery.date == start).first()
            NumSurgComp_score = max(((10-(NumSurgComp[0] / NumSurg[0]) * 100)),-10)

            # 计算科室总分
            total_score = qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score + \
                          OPRecRate_score + ARateOPRec_score + ARateIPRec_score +\
                          OANumIPRec_score + NumPath_score + CMI_score + AverageBedTime_score + NumSurg_score + \
                          NumMiroInvaSurg_score + Num4thSurg_score + NumSurgComp_score
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score, '门诊病历书写率得分': OPRecRate_score,
                               '门诊病历甲级率得分': ARateOPRec_score, '住院病历甲级率':ARateIPRec_score , '住院按时归档率':OANumIPRec_score,
                               '临床路径入径率得分': NumPath_score ,'CMI得分':CMI_score ,'平均住院日得分':AverageBedTime_score ,
                               '手术患者占比得分':NumSurg_score ,'微创手术占比得分':NumMiroInvaSurg_score ,
                               '四级手术占比得分':Num4thSurg_score ,'手术并发症占比得分':NumSurgComp_score ,'总分': total_score}
            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        #定义查询计算新生儿科8个指标的得分的函数
        def queryneonatus(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == '儿科',
                                                                 Quality.date == start).first()
            Qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 10
            # 5--门诊病历书写率得分
            # NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept, OutPatient.date == start).first()
            # NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
            #                                                         OutPatient.date == start).first()
            # OPRecRate_score = NumOPRec[0] / NumOP[0] * 10
            # 6--门诊病历甲级率得分
            # CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
            #                                                                   OutPatient.date == start).first()
            # ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
            #                                                           OutPatient.date == start).first()
            # if CheckNumOPRec != 0:
            #     ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            # else:
            #     ARateOPRec_score = 10
            # 7--住院病历甲级率得分
            CheckNumIPRec = db.session.query(InPatient.CheckNumIPRec).filter(InPatient.deptname == dept,
                                                                              InPatient.date == start).first()
            ANumIPRec = db.session.query(InPatient.ANumIPRec).filter(InPatient.deptname == dept,
                                                                      InPatient.date == start).first()
            if CheckNumIPRec[0] != 0:
                ARateIPRec_score = ANumIPRec[0] / CheckNumIPRec[0] * 10
            else:
                ARateIPRec_score = 10
            # 8--住院病历按时归档率
            OANumIPRec = db.session.query(InPatient.OANumIPRec).filter(InPatient.deptname == dept,
                                                                             InPatient.date == start).first()
            NumIP = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                     InPatient.date == start).first()
            OANumIPRec_score = (1-OANumIPRec[0] / NumIP[0]) * 10
            # 9--临床路径入径率得分
            NumPath = db.session.query(InPatient.NumPath).filter(InPatient.deptname == dept,
                                                                       InPatient.date == start).first()
            NumPath_score = NumPath[0] / NumIP[0] * 10
            # 10--CMI得分
            CMI_Now = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            CMI_Lastyear = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                             InPatient.date == lastyear).first()
            CMI_score = max(min((CMI_Now[0] - CMI_Lastyear[0])/CMI_Lastyear[0]*150,10),-10)

            # 11--平均住院日得分
            NumBD_Now = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            NumBD_Lastyear = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                  InPatient.date == lastyear).first()
            NumIP_Lastyear = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            AverageBedTime_score = max(min((NumBD_Lastyear[0]/NumIP_Lastyear[0] - NumBD_Now[0]/NumIP[0]) /(NumBD_Lastyear[0] / NumIP_Lastyear[0])* 50 , 10),-10)

            # 计算科室总分
            total_score = (Qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score +
                           ARateIPRec_score + OANumIPRec_score + NumPath_score + CMI_score + AverageBedTime_score)/0.9
            dept_score_dict = {'危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score,
                               '住院病历甲级率':ARateIPRec_score , '住院按时归档率':OANumIPRec_score,
                               '临床路径入径率得分': NumPath_score ,'CMI得分':CMI_score ,'平均住院日得分':AverageBedTime_score , '总分': total_score}
            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算重症医学科8个指标的得分的函数
        def queryicu(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            Qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 10
            # 5--门诊病历书写率得分
            # NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept, OutPatient.date == start).first()
            # NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
            #                                                         OutPatient.date == start).first()
            # OPRecRate_score = NumOPRec[0] / NumOP[0] * 10
            # 6--门诊病历甲级率得分
            # CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
            #                                                                   OutPatient.date == start).first()
            # ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
            #                                                           OutPatient.date == start).first()
            # if CheckNumOPRec != 0:
            #     ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            # else:
            #     ARateOPRec_score = 10
            # 7--住院病历甲级率得分
            CheckNumIPRec = db.session.query(InPatient.CheckNumIPRec).filter(InPatient.deptname == dept,
                                                                             InPatient.date == start).first()
            ANumIPRec = db.session.query(InPatient.ANumIPRec).filter(InPatient.deptname == dept,
                                                                     InPatient.date == start).first()
            if CheckNumIPRec[0] != 0:
                ARateIPRec_score = ANumIPRec[0] / CheckNumIPRec[0] * 10
            else:
                ARateIPRec_score = 10
            # 8--住院病历按时归档率
            OANumIPRec = db.session.query(InPatient.OANumIPRec).filter(InPatient.deptname == dept,
                                                                       InPatient.date == start).first()
            NumIP = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            OANumIPRec_score = (1 - OANumIPRec[0] / NumIP[0]) * 10
            # # 9--临床路径入径率得分
            # NumPath = db.session.query(InPatient.NumPath).filter(InPatient.deptname == dept,
            #                                                      InPatient.date == start).first()
            # NumPath_score = NumPath[0] / NumIP[0] * 10
            # 10--CMI得分
            CMI_Now = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            CMI_Lastyear = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                                  InPatient.date == lastyear).first()
            CMI_score = max(min((CMI_Now[0] - CMI_Lastyear[0]) /CMI_Lastyear[0]* 150, 10),-10)

            # 11--平均住院日得分
            NumBD_Now = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            NumBD_Lastyear = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            NumIP_Lastyear = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            AverageBedTime_score = max(min((NumBD_Lastyear[0] / NumIP_Lastyear[0] - NumBD_Now[0] / NumIP[0])/(NumBD_Lastyear[0] / NumIP_Lastyear[0]) * 50, 10),-10)

            # 计算科室总分
            total_score = (Qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score +
                          ARateIPRec_score + OANumIPRec_score + CMI_score + AverageBedTime_score)/0.8
            dept_score_dict = {'质量安全管理小组得分': Qccscore_score,'危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score,
                               '住院病历甲级率': ARateIPRec_score, '住院按时归档率': OANumIPRec_score,
                                'CMI得分': CMI_score, '平均住院日得分': AverageBedTime_score,
                               '总分': total_score}

            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算产科13个指标的得分的函数
        def queryObste(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()

            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--大型检查阳性率得分
            NumCTMri = db.session.query(Quality.NumCTMri).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()
            NumCTMriPositive = db.session.query(Quality.NumCTMriPositive).filter(Quality.deptname == dept,
                                                                                 Quality.date == start).first()
            if NumCTMri[0] != 0:
                CTMripositiveRate_score = NumCTMriPositive[0] / NumCTMri[0] * 10
            else:
                CTMripositiveRate_score = 10
            # 5--门诊病历书写率得分
            NumOP = db.session.query(OutPatient.NumOP).filter(OutPatient.deptname == dept,
                                                              OutPatient.date == start).first()
            NumOPRec = db.session.query(OutPatient.NumOPRec).filter(OutPatient.deptname == dept,
                                                                    OutPatient.date == start).first()
            OPRecRate_score = 10
            # 6--门诊病历甲级率得分
            CheckNumOPRec = db.session.query(OutPatient.CheckNumOPRec).filter(OutPatient.deptname == dept,
                                                                              OutPatient.date == start).first()
            ANumOPRec = db.session.query(OutPatient.ANumOPRec).filter(OutPatient.deptname == dept,
                                                                      OutPatient.date == start).first()
            if CheckNumOPRec[0] != 0:
                ARateOPRec_score = ANumOPRec[0] / CheckNumOPRec[0] * 10
            else:
                ARateOPRec_score = 10
            # 7--住院病历甲级率得分
            CheckNumIPRec = db.session.query(InPatient.CheckNumIPRec).filter(InPatient.deptname == dept,
                                                                             InPatient.date == start).first()
            ANumIPRec = db.session.query(InPatient.ANumIPRec).filter(InPatient.deptname == dept,
                                                                     InPatient.date == start).first()
            if CheckNumIPRec[0] != 0:
                ARateIPRec_score = ANumIPRec[0] / CheckNumIPRec[0] * 10
            else:
                ARateIPRec_score = 10
            # 8--住院病历按时归档率
            OANumIPRec = db.session.query(InPatient.OANumIPRec).filter(InPatient.deptname == dept,
                                                                       InPatient.date == start).first()
            NumIP = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            OANumIPRec_score = (1 - OANumIPRec[0] / NumIP[0]) * 10
            # 9--临床路径入径率得分
            NumPath = db.session.query(InPatient.NumPath).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            NumPath_score = NumPath[0] / NumIP[0] * 10
            # 10--CMI得分
            CMI_Now = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                             InPatient.date == start).first()
            CMI_Lastyear = db.session.query(InPatient.CMI).filter(InPatient.deptname == dept,
                                                                  InPatient.date == lastyear).first()
            if CMI_Lastyear is None:
                CMI_score = 10
            else:
                CMI_score = max(min((CMI_Now[0] - CMI_Lastyear[0])/CMI_Lastyear[0] * 150, 10), -10)


            # 11--平均住院日得分
            NumBD_Now = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                 InPatient.date == start).first()
            NumBD_Lastyear = db.session.query(InPatient.NumBD).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            NumIP_Lastyear = db.session.query(InPatient.NumIP).filter(InPatient.deptname == dept,
                                                                      InPatient.date == lastyear).first()
            if NumBD_Lastyear is None:
                AverageBedTime_score = 10
            else:
                AverageBedTime_score = max(
                    min((NumBD_Lastyear[0] / NumIP_Lastyear[0] - NumBD_Now[0] / NumIP[0]) /(NumBD_Lastyear[0] / NumIP_Lastyear[0])* 50, 10), -10)

            # 12--手术患者占比得分
            NumSurg = db.session.query(Surgery.NumSurg).filter(Surgery.deptname == dept,
                                                               Surgery.date == start).first()
            NumSurg_score = min(max(10+(0.35-NumSurg[0] / NumIP[0]) * 50,0),10)

            # # 12--微创手术占比得分
            # NumMiroInvaSurg = db.session.query(Surgery.NumMiroInvaSurg).filter(Surgery.deptname == dept,
            #                                                                    Surgery.date == start).first()
            # NumEleSurg = db.session.query(Surgery.NumEleSurg).filter(Surgery.deptname == dept,
            #                                                          Surgery.date == start).first()
            # NumMiroInvaSurg_score = NumMiroInvaSurg / NumEleSurg * 10

            # 13--四级手术占比得分
            Num4thSurg = db.session.query(Surgery.Num4thSurg).filter(Surgery.deptname == dept,
                                                                     Surgery.date == start).first()
            Num4thSurg_score = Num4thSurg[0] / NumSurg[0] * 10

            # 14--手术并发症占比得分
            NumSurgComp = db.session.query(Surgery.NumSurgComp).filter(Surgery.deptname == dept,
                                                                       Surgery.date == start).first()
            NumSurgComp_score = max(((10 - (NumSurgComp[0] / NumSurg[0]) * 100)), -10)

            # 计算科室总分
            total_score = qccscore_score + CriVal_score + AdvEvtRep_score + CTMripositiveRate_score + \
                          OPRecRate_score + ARateOPRec_score + ARateIPRec_score + \
                          OANumIPRec_score + NumPath_score + CMI_score + AverageBedTime_score + NumSurg_score + \
                          Num4thSurg_score + NumSurgComp_score
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '大型检查阳性率得分': CTMripositiveRate_score, '门诊病历书写率得分': OPRecRate_score,
                               '门诊病历甲级率得分': ARateOPRec_score, '住院病历甲级率': ARateIPRec_score,
                               '住院按时归档率': OANumIPRec_score,
                               '临床路径入径率得分': NumPath_score, 'CMI得分': CMI_score, '平均住院日得分': AverageBedTime_score,
                               '手术患者占比得分': NumSurg_score,
                               '四级手术占比得分': Num4thSurg_score, '手术并发症占比得分': NumSurgComp_score, '总分': total_score}

            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算麻醉科6个指标的得分的函数
        def queryAnes(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()

            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--术前访视得分
            ratepreaneseval = db.session.query(Anestdata.ratepreaneseval).filter(Anestdata.deptname == dept,
                                                                 Anestdata.date == start).first()
            ratepreaneseval_score = ratepreaneseval[0] * 10
            # 5--麻醉复苏得分
            rateanesresu = db.session.query(Anestdata.rateanesresu).filter(Anestdata.deptname == dept,
                                                                           Anestdata.date == start).first()
            rateanesresu_score = rateanesresu[0] * 10
            # 6--术后访视得分
            ratepostsurgvisit = db.session.query(Anestdata.ratepostsurgvisit).filter(Anestdata.deptname == dept,
                                                                              OutPatient.date == start).first()
            ratepostsurgvisit_score = ratepostsurgvisit[0] * 10
            # 计算科室总分
            total_score = (qccscore_score + CriVal_score + AdvEvtRep_score + ratepreaneseval_score + rateanesresu_score + ratepostsurgvisit_score)/0.6
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '术前访视得分': ratepreaneseval_score, '麻醉复苏得分': rateanesresu_score,
                               '术后访视得分': ratepostsurgvisit_score, '总分': total_score}
            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义查询计算医技科室6个指标的得分的函数
        def queryExam(dept):
            # 1--质量管理小组评分
            qccscore = db.session.query(Quality.QCCScore).filter(Quality.deptname == dept,
                                                                 Quality.date == start).first()

            qccscore_score = qccscore[0] * 10 / 100
            # 2--危急值规范处理得分
            CriValRepNum = db.session.query(Quality.CriValRepNum).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            CriValHPNum = db.session.query(Quality.CriValHPNum).filter(Quality.deptname == dept,
                                                                       Quality.date == start).first()
            if CriValRepNum[0] != 0:
                CriVal_score = CriValHPNum[0] / CriValRepNum[0] * 10
            else:
                CriVal_score = 10
            # 3--不良事件报告得分
            NumAdvEvtRep = db.session.query(Quality.NumAdvEvtRep).filter(Quality.deptname == dept,
                                                                         Quality.date == start).first()
            AdvEvtRep_score = min(NumAdvEvtRep[0] * 5, 10)
            # 4--设备故障率得分
            MedEquipMalFunc = db.session.query(ExamTestData.MedEquipMalFunc).filter(ExamTestData.deptname == dept,
                                                                 ExamTestData.date == start).first()
            MedEquipTime = db.session.query(ExamTestData.MedEquipTime).filter(ExamTestData.deptname == dept,
                                                                                    ExamTestData.date == start).first()
            MedEquipMalFunc_score = min((1-MedEquipMalFunc[0]/MedEquipTime[0]) * 100,10)
            # 5--报告及时率得分
            AccuRateInspRept = db.session.query(ExamTestData.AccuRateInspRept).filter(ExamTestData.deptname == dept,
                                                                           ExamTestData.date == start).first()
            AccuRateInspRept_score = AccuRateInspRept[0] * 10
            # 6--报告准确率得分
            TimelyRateInspRept = db.session.query(ExamTestData.TimelyRateInspRept).filter(ExamTestData.deptname == dept,
                                                                                      ExamTestData.date == start).first()
            TimelyRateInspRept_score = TimelyRateInspRept[0] * 10
            # 计算科室总分
            total_score = qccscore_score + CriVal_score + AdvEvtRep_score + MedEquipMalFunc_score + AccuRateInspRept_score + TimelyRateInspRept_score
            dept_score_dict = {'质量安全管理小组得分': qccscore_score, '危急值得分': CriVal_score, '不良事件得分': AdvEvtRep_score,
                               '设备故障率得分': MedEquipMalFunc_score, '报告及时率得分': AccuRateInspRept_score,
                               '报告准确率得分': TimelyRateInspRept_score, '总分': total_score}

            # 科室分数加入字典
            depts_score_list[dept] = dept_score_dict

        # 定义质量考核报告函数
        def quality_score_report(depts_score_list):
            wb = Workbook()
            # 手术科室列表
            dept_surg_list = []
            for d in deptssurg:
                dept_surg_list.append(d[0])
            # 住院科室列表
            depts_IP_list = []
            for d in deptsIP:
                depts_IP_list.append(d[0])
            # 住院科室列表
            depts_OP_list = []
            for d in deptsOP:
                depts_OP_list.append(d[0])
            # 医技科室列表
            depts_Exam_list = []
            for d in deptsExam:
                depts_Exam_list.append(d[0])
            ws_score = wb.create_sheet("质量考核分数汇总表")  # 创建所有科室分数汇总表
            ws_score.column_dimensions['A'].width = 20.0
            ws_score.column_dimensions['C'].width = 20.0
            ws_score.column_dimensions['E'].width = 22.0
            ws_score.column_dimensions['G'].width = 22.0
            ws_score['a1'] = str(start.year) + "年" + str(start.month) + "月" + "质量考核分数汇总表"
            ws_score['a1'].alignment = Alignment(horizontal='left', wrap_text=False)
            ws_score['a1'].font = Font(bold=True, size=20)
            ws_score['a3'] = '手术科室'
            ws_score['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
            ws_score['a3'].font = Font(bold=True, size=14)
            ws_score['c3'] = '住院科室'
            ws_score['c3'].alignment = Alignment(horizontal='left', wrap_text=False)
            ws_score['c3'].font = Font(bold=True, size=14)
            ws_score['e3'] = '门诊科室'
            ws_score['e3'].alignment = Alignment(horizontal='left', wrap_text=False)
            ws_score['e3'].font = Font(bold=True, size=14)
            ws_score['g3'] = '医技科室及麻醉科'
            ws_score['g3'].alignment = Alignment(horizontal='left', wrap_text=False)
            ws_score['g3'].font = Font(bold=True, size=14)
            # 遍历分数汇总字典，写入汇总表
            row1 = 4
            row2 = 4
            row3 = 4
            row4 = 4
            for dept in depts_score_list:
                if dept in dept_surg_list:
                    ws_score.cell(row=row1,column=1,value = dept)
                    ws_score.cell(row=row1,column=2,value = round(depts_score_list[dept]['总分'],2))
                    row1 += 1
                elif dept in depts_IP_list or dept == '重症医学科':
                    ws_score.cell(row=row2, column=3, value=dept)
                    ws_score.cell(row=row2, column=4, value=round(depts_score_list[dept]['总分'],2))
                    row2 += 1
                elif dept in depts_OP_list:
                    ws_score.cell(row=row3, column=5, value=dept)
                    ws_score.cell(row=row3, column=6, value=round(depts_score_list[dept]['总分'],2))
                    row3 += 1
                elif dept in depts_Exam_list or dept == '麻醉科':
                    ws_score.cell(row=row4, column=7, value=dept)
                    ws_score.cell(row=row4, column=8, value=round(depts_score_list[dept]['总分'],2))
                    row4 += 1



            for dept in depts_score_list: #创建科室质量考核反馈表
                ws = wb.create_sheet(dept)
                ws.title = dept
                ws.column_dimensions['A'].width = 90.0

                #手术科室除去产科
                if dept in dept_surg_list and dept != '产科':
                    print('正在处理 %s 数据' %dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text = False)
                    ws['a1'].font = Font(bold=True,size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）"+" 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                               "有科室质量与安全管理工作计划并实施。" \
                               "有科室质量与安全工作制度并落实。" \
                               "有科室质量与安全管理的各项工作记录。" \
                               "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                               "对本科室质量与安全指标进行资料收集和分析。" \
                               "能够运用质量管理方法与工具进行持续质量改进。" \
                               "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                                                "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                                                "每百张床位年报告≥20 件。" \
                                                "对医疗安全（不良）事件有分析，采取防范措施。" \
                                                "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'],2)) \
                               + " 分。（满分10分）" + \
                               " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                               "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                               "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + str(round(depts_score_list[dept]['门诊病历书写率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "七、住院病历甲级率得分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院病历甲级率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a16'] = "    贵科本考核周期住院病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.4.2要求院科两级落实整改措施，持续改进病历质量，年度住院病案总检查数占总住院病案数≥70%，病历甲级率≥90%，无丙级病历。"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)
                    ws['a17'] = "八、住院病历按时归档率得分"
                    ws['a17'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a17'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院按时归档率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a18'] = "    贵科本考核周期住院病历按时归档率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.6要求患者出院后，住院病历在 2 个工作日之内回归病案科达≥95％，在 7 个工作日内回归病案科 100%。"
                    ws['a18'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a18'].font = Font(bold=False, size=12)
                    ws['a19'] = "九、临床路径入径率得分"
                    ws['a19'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a19'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['临床路径入径率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a20'] = "    贵科本考核周期临床路径入径率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款要求医院将开展临床路径与单病种质量管理作为推动医疗质量持续改进的重点项目，规范临床诊疗行为的重要内容之一；" \
                                "对符合进入临床路径标准的患者达到入组率≥50%，入组完成率≥70%。"
                    ws['a20'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a20'].font = Font(bold=False, size=12)
                    ws['a21'] = "十、CMI得分"
                    ws['a21'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a21'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['CMI得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a22'] = "    贵科本考核周期CMI得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，增长为正分，满分10分，下降为扣分，上限-10分）" + \
                                " 说明：根据《广东省卫生健康委关于进步完善三级医院(综合医院、专科医院)等级评审工作的通知》要求，" \
                                "三级综合医院首次申报评审，申报前近一年，三级公立医院绩效考核排名、CMI排名均需在大湾区前75%以内。"
                    ws['a22'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a22'].font = Font(bold=False, size=12)
                    ws['a23'] = "十一、平均住院日得分"
                    ws['a23'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a23'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['平均住院日得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a24'] = "    贵科本考核周期平均住院日得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，降低为正分，满分10分，上升为扣分，上限-10分）" + \
                                " 说明：三甲条款4.5.7.4(核心条款)要求：医院对各临床科室出院患者平均住院日有明确的要求。" \
                                "有缩短平均住院日的具体措施。相关管理人员与医师均知晓缩短平均住院日的要求，并落实各项措施。平均住院日达到控制目标。"
                    ws['a24'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a24'].font = Font(bold=False, size=12)
                    ws['a25'] = "十二、手术患者占比得分"
                    ws['a25'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a25'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['手术患者占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a26'] = "    贵科本考核周期手术患者占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第四项：出院患者手术占比（国家监测指标）。" \
                                "手术和介入治疗的数量尤其是疑难复杂手术和介入治疗的数量与医院的规模、人员、设备、设施等" \
                                "综合诊疗技术能力及临床管理流程成正相关，鼓励三级医院优质医疗资源服务于疑难危重患者，" \
                                "尤其是能够提供安全有保障的高质量医疗技术服务。"
                    ws['a26'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a26'].font = Font(bold=False, size=12)
                    ws['a27'] = "十三、微创手术占比得分"
                    ws['a27'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a27'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['微创手术占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a28'] = "    贵科本考核周期微创手术占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第5项：出院患者微创手术占比（国家监测指标）。" \
                                "微创手术降低了传统手术对人体的伤害，具有创伤小、疼痛轻、恢复快的优越性，极大地减少了" \
                                "疾病给患者带来的不便和痛苦，更注重患者的心理、社会、生理（疼痛）、精神、生活质量的改善与康复，" \
                                "减轻患者的痛苦。（2）合理选择微创技术适应症、控制相关技术风险促进微创技术发展。"
                    ws['a28'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a28'].font = Font(bold=False, size=12)
                    ws['a29'] = "十四、四级手术占比得分"
                    ws['a29'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a29'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['四级手术占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a30'] = "    贵科本考核周期四级手术占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第6项：出院患者四级手术比例（国家监测指标）。" \
                                "《关于印发医疗机构手术分级管理办法（试行）的通知》（卫办医政发〔2012〕94 号）提出" \
                                "医疗机构应当开展与其级别和诊疗科目相适应的手术。三级医院重点开展三、四级手术。"
                    ws['a30'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a30'].font = Font(bold=False, size=12)
                    ws['a31'] = "十五、手术并发症占比得分"
                    ws['a31'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a31'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['手术并发症占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a32'] = "    贵科本考核周期手术并发症占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第8项：手术患者并发症发生率（国家监测指标）。" \
                                "预防手术后并发症发生是医疗质量管理和监控的重点，也是患者安全管理的核心内容，" \
                                "是衡量医疗技术能力和管理水平的重要结果指标之一。"
                    ws['a32'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a32'].font = Font(bold=False, size=12)
                    ws['a33'] = "总分"
                    ws['a33'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a33'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10 / 15,2))
                    ws['a34'] = "    贵科本考核周期总得分为 " + fen + " 分。（15项指标原始得分）"
                    ws['a34'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a34'].font = Font(bold=False, size=12)

                # 住院科室除icu及新生儿
                elif dept in depts_IP_list and dept != '重症医学科' and dept != '新生儿科':
                    print('正在处理 %s 数据' %dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text = False)
                    ws['a1'].font = Font(bold=True,size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）"+" 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                               "有科室质量与安全管理工作计划并实施。" \
                               "有科室质量与安全工作制度并落实。" \
                               "有科室质量与安全管理的各项工作记录。" \
                               "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                               "对本科室质量与安全指标进行资料收集和分析。" \
                               "能够运用质量管理方法与工具进行持续质量改进。" \
                               "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                                                "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                                                "每百张床位年报告≥20 件。" \
                                                "对医疗安全（不良）事件有分析，采取防范措施。" \
                                                "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'],2)) \
                               + " 分。（满分10分）" + \
                               " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                               "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                               "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + str(round(depts_score_list[dept]['门诊病历书写率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "七、住院病历甲级率得分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院病历甲级率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a16'] = "    贵科本考核周期住院病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.4.2要求院科两级落实整改措施，持续改进病历质量，年度住院病案总检查数占总住院病案数≥70%，病历甲级率≥90%，无丙级病历。"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)
                    ws['a17'] = "八、住院病历按时归档率得分"
                    ws['a17'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a17'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院按时归档率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a18'] = "    贵科本考核周期住院病历按时归档率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.6要求患者出院后，住院病历在 2 个工作日之内回归病案科达≥95％，在 7 个工作日内回归病案科 100%。"
                    ws['a18'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a18'].font = Font(bold=False, size=12)
                    ws['a19'] = "九、临床路径入径率得分"
                    ws['a19'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a19'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['临床路径入径率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a20'] = "    贵科本考核周期临床路径入径率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款要求医院将开展临床路径与单病种质量管理作为推动医疗质量持续改进的重点项目，规范临床诊疗行为的重要内容之一；" \
                                "对符合进入临床路径标准的患者达到入组率≥50%，入组完成率≥70%。"
                    ws['a20'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a20'].font = Font(bold=False, size=12)
                    ws['a21'] = "十、CMI得分"
                    ws['a21'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a21'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['CMI得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a22'] = "    贵科本考核周期CMI得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，增长为正分，满分10分，下降为扣分，上限-10分）" + \
                                " 说明：根据《广东省卫生健康委关于进步完善三级医院(综合医院、专科医院)等级评审工作的通知》要求，" \
                                "三级综合医院首次申报评审，申报前近一年，三级公立医院绩效考核排名、CMI排名均需在大湾区前75%以内。"
                    ws['a22'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a22'].font = Font(bold=False, size=12)
                    ws['a23'] = "十一、平均住院日得分"
                    ws['a23'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a23'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['平均住院日得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a24'] = "    贵科本考核周期平均住院日得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，降低为正分，满分10分，上升为扣分，上限-10分）" + \
                                " 说明：三甲条款4.5.7.4(核心条款)要求：医院对各临床科室出院患者平均住院日有明确的要求。" \
                                "有缩短平均住院日的具体措施。相关管理人员与医师均知晓缩短平均住院日的要求，并落实各项措施。平均住院日达到控制目标。"
                    ws['a24'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a24'].font = Font(bold=False, size=12)
                    ws['a25'] = "总分"
                    ws['a25'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a25'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10 / 11,2))
                    ws['a26'] = "    贵科本考核周期总得分为 " + fen + " 分。（11项指标原始得分）"
                    ws['a26'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a26'].font = Font(bold=False, size=12)

                # 门诊科室
                elif dept in depts_OP_list :
                    print('正在处理 %s 数据' %dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text = False)
                    ws['a1'].font = Font(bold=True,size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）"+" 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                               "有科室质量与安全管理工作计划并实施。" \
                               "有科室质量与安全工作制度并落实。" \
                               "有科室质量与安全管理的各项工作记录。" \
                               "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                               "对本科室质量与安全指标进行资料收集和分析。" \
                               "能够运用质量管理方法与工具进行持续质量改进。" \
                               "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                                                "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                                                "每百张床位年报告≥20 件。" \
                                                "对医疗安全（不良）事件有分析，采取防范措施。" \
                                                "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'],2)) \
                               + " 分。（满分10分）" + \
                               " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                               "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                               "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + str(round(depts_score_list[dept]['门诊病历书写率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "总分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10/6,2))
                    ws['a16'] = "    贵科本考核周期总得分为 " + fen + " 分。（6项指标原始得分）"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)

                # 医技科室
                elif dept in depts_Exam_list :
                    print('正在处理 %s 数据' %dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text = False)
                    ws['a1'].font = Font(bold=True,size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）"+" 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                               "有科室质量与安全管理工作计划并实施。" \
                               "有科室质量与安全工作制度并落实。" \
                               "有科室质量与安全管理的各项工作记录。" \
                               "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                               "对本科室质量与安全指标进行资料收集和分析。" \
                               "能够运用质量管理方法与工具进行持续质量改进。" \
                               "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                                                "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                                                "每百张床位年报告≥20 件。" \
                                                "对医疗安全（不良）事件有分析，采取防范措施。" \
                                                "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、设备故障率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期设备故障率得分为 " + str(round(depts_score_list[dept]['设备故障率得分'],2)) \
                               + " 分。（满分10分）" + \
                               " 说明：设备故障率为三级公立医院绩效考核指标12。" \
                               "引导医院关注医用设备的维修保养和质量控制，配置合适维修人员和维修检测设备。" \
                               "《卫生部办公厅关于印发<三级综合医院评审标准实施细则（2011 年版）>的通知》" \
                               "(卫办医管发〔2011〕148 号)中要求，医学装备管理符合国家法律、法规及卫生行政部门" \
                               "规章、管理办法、标准的要求，按照法律、法规，使用和管理医用含源仪器（装置）。" \
                               "医院应当关注医用设备的维修保养和质量控制，科学配置工程技术人员并做好设备维修保养等管理工作。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、报告及时率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期报告及时率得分为 " + str(round(depts_score_list[dept]['报告及时率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：《卫生部办公厅关于印发<三级综合医院评审标准实施细则（2011 年版）>的通知》中要求，" \
                                "科室有诊断报告书写规范、审核制度与流程。报告由具备资质的专业医师出具。有提供影像报告时限要求。" \
                                "每份报告书有精确的报告时间，普通报告精确到“时”，急诊报告精确到“分”。诊断报告按照流程经过审核，有审核医师签名。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、报告准确率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期报告准确率得分为 " + str(round(depts_score_list[dept]['报告准确率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：《卫生部办公厅关于印发<三级综合医院评审标准实施细则（2011 年版）>的通知》中要求，" \
                                "科室有诊断报告书写规范、审核制度与流程。报告由具备资质的专业医师出具。有提供影像报告时限要求。" \
                                "每份报告书有精确的报告时间，普通报告精确到“时”，急诊报告精确到“分”。诊断报告按照流程经过审核，有审核医师签名。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "总分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10/ 6,2))
                    ws['a16'] = "    贵科本考核周期总得分为 " + fen + " 分。（6项指标原始得分）"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)

                # 产科
                elif dept == '产科':
                    print('正在处理 %s 数据' % dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text=False)
                    ws['a1'].font = Font(bold=True, size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                                                "有科室质量与安全管理工作计划并实施。" \
                                                "有科室质量与安全工作制度并落实。" \
                                                "有科室质量与安全管理的各项工作记录。" \
                                                "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                                                "对本科室质量与安全指标进行资料收集和分析。" \
                                                "能够运用质量管理方法与工具进行持续质量改进。" \
                                                "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                               "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                               "每百张床位年报告≥20 件。" \
                               "对医疗安全（不良）事件有分析，采取防范措施。" \
                               "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                                "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                                "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + str(round(depts_score_list[dept]['门诊病历书写率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "七、住院病历甲级率得分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院病历甲级率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a16'] = "    贵科本考核周期住院病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.4.2要求院科两级落实整改措施，持续改进病历质量，年度住院病案总检查数占总住院病案数≥70%，病历甲级率≥90%，无丙级病历。"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)
                    ws['a17'] = "八、住院病历按时归档率得分"
                    ws['a17'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a17'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院按时归档率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a18'] = "    贵科本考核周期住院病历按时归档率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.6要求患者出院后，住院病历在 2 个工作日之内回归病案科达≥95％，在 7 个工作日内回归病案科 100%。"
                    ws['a18'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a18'].font = Font(bold=False, size=12)
                    ws['a19'] = "九、临床路径入径率得分"
                    ws['a19'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a19'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['临床路径入径率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a20'] = "    贵科本考核周期临床路径入径率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款要求医院将开展临床路径与单病种质量管理作为推动医疗质量持续改进的重点项目，规范临床诊疗行为的重要内容之一；" \
                                "对符合进入临床路径标准的患者达到入组率≥50%，入组完成率≥70%。"
                    ws['a20'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a20'].font = Font(bold=False, size=12)
                    ws['a21'] = "十、CMI得分"
                    ws['a21'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a21'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['CMI得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a22'] = "    贵科本考核周期CMI得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，增长为正分，满分10分，下降为扣分，上限-10分）" + \
                                " 说明：根据《广东省卫生健康委关于进步完善三级医院(综合医院、专科医院)等级评审工作的通知》要求，" \
                                "三级综合医院首次申报评审，申报前近一年，三级公立医院绩效考核排名、CMI排名均需在大湾区前75%以内。"
                    ws['a22'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a22'].font = Font(bold=False, size=12)
                    ws['a23'] = "十一、平均住院日得分"
                    ws['a23'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a23'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['平均住院日得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a24'] = "    贵科本考核周期平均住院日得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，降低为正分，满分10分，上升为扣分，上限-10分）" + \
                                " 说明：三甲条款4.5.7.4(核心条款)要求：医院对各临床科室出院患者平均住院日有明确的要求。" \
                                "有缩短平均住院日的具体措施。相关管理人员与医师均知晓缩短平均住院日的要求，并落实各项措施。平均住院日达到控制目标。"
                    ws['a24'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a24'].font = Font(bold=False, size=12)
                    ws['a25'] = "十二、手术患者占比得分"
                    ws['a25'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a25'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['手术患者占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a26'] = "    贵科本考核周期手术患者占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：按照产科专业特点和上级对剖宫产率的要求，计分方法与其他科室不同。"
                    ws['a26'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a26'].font = Font(bold=False, size=12)
                    ws['a27'] = "十三、微创手术占比得分"
                    ws['a27'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a27'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['微创手术占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a28'] = "    贵科本考核周期微创手术占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第5项：出院患者微创手术占比（国家监测指标）。" \
                                "微创手术降低了传统手术对人体的伤害，具有创伤小、疼痛轻、恢复快的优越性，极大地减少了" \
                                "疾病给患者带来的不便和痛苦，更注重患者的心理、社会、生理（疼痛）、精神、生活质量的改善与康复，" \
                                "减轻患者的痛苦。（2）合理选择微创技术适应症、控制相关技术风险促进微创技术发展。"
                    ws['a28'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a28'].font = Font(bold=False, size=12)
                    ws['a29'] = "十四、四级手术占比得分"
                    ws['a29'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a29'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['四级手术占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a30'] = "    贵科本考核周期四级手术占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第6项：出院患者四级手术比例（国家监测指标）。" \
                                "《关于印发医疗机构手术分级管理办法（试行）的通知》（卫办医政发〔2012〕94 号）提出" \
                                "医疗机构应当开展与其级别和诊疗科目相适应的手术。三级医院重点开展三、四级手术。"
                    ws['a30'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a30'].font = Font(bold=False, size=12)
                    ws['a31'] = "十五、手术并发症占比得分"
                    ws['a31'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a31'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['手术并发症占比得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a32'] = "    贵科本考核周期手术并发症占比得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：国家三级公立医院绩效考核指标第8项：手术患者并发症发生率（国家监测指标）。" \
                                "预防手术后并发症发生是医疗质量管理和监控的重点，也是患者安全管理的核心内容，" \
                                "是衡量医疗技术能力和管理水平的重要结果指标之一。"
                    ws['a32'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a32'].font = Font(bold=False, size=12)
                    ws['a33'] = "总分"
                    ws['a33'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a33'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10 / 14, 2))
                    ws['a34'] = "    贵科本考核周期总得分为 " + fen + " 分。（14项指标原始得分）"
                    ws['a34'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a34'].font = Font(bold=False, size=12)

                # icu
                elif dept == '重症医学科':
                    print('正在处理 %s 数据' % dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text=False)
                    ws['a1'].font = Font(bold=True, size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                                                "有科室质量与安全管理工作计划并实施。" \
                                                "有科室质量与安全工作制度并落实。" \
                                                "有科室质量与安全管理的各项工作记录。" \
                                                "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                                                "对本科室质量与安全指标进行资料收集和分析。" \
                                                "能够运用质量管理方法与工具进行持续质量改进。" \
                                                "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                               "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                               "每百张床位年报告≥20 件。" \
                               "对医疗安全（不良）事件有分析，采取防范措施。" \
                               "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                                "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                                "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['门诊病历书写率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "七、住院病历甲级率得分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院病历甲级率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a16'] = "    贵科本考核周期住院病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.4.2要求院科两级落实整改措施，持续改进病历质量，年度住院病案总检查数占总住院病案数≥70%，病历甲级率≥90%，无丙级病历。"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)
                    ws['a17'] = "八、住院病历按时归档率得分"
                    ws['a17'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a17'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院按时归档率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a18'] = "    贵科本考核周期住院病历按时归档率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.6要求患者出院后，住院病历在 2 个工作日之内回归病案科达≥95％，在 7 个工作日内回归病案科 100%。"
                    ws['a18'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a18'].font = Font(bold=False, size=12)
                    ws['a19'] = "九、临床路径入径率得分"
                    ws['a19'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a19'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['临床路径入径率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a20'] = "    贵科本考核周期临床路径入径率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款要求医院将开展临床路径与单病种质量管理作为推动医疗质量持续改进的重点项目，规范临床诊疗行为的重要内容之一；" \
                                "对符合进入临床路径标准的患者达到入组率≥50%，入组完成率≥70%。"
                    ws['a20'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a20'].font = Font(bold=False, size=12)
                    ws['a21'] = "十、CMI得分"
                    ws['a21'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a21'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['CMI得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a22'] = "    贵科本考核周期CMI得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，增长为正分，满分10分，下降为扣分，上限-10分）" + \
                                " 说明：根据《广东省卫生健康委关于进步完善三级医院(综合医院、专科医院)等级评审工作的通知》要求，" \
                                "三级综合医院首次申报评审，申报前近一年，三级公立医院绩效考核排名、CMI排名均需在大湾区前75%以内。"
                    ws['a22'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a22'].font = Font(bold=False, size=12)
                    ws['a23'] = "十一、平均住院日得分"
                    ws['a23'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a23'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['平均住院日得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a24'] = "    贵科本考核周期平均住院日得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，降低为正分，满分10分，上升为扣分，上限-10分）" + \
                                " 说明：三甲条款4.5.7.4(核心条款)要求：医院对各临床科室出院患者平均住院日有明确的要求。" \
                                "有缩短平均住院日的具体措施。相关管理人员与医师均知晓缩短平均住院日的要求，并落实各项措施。平均住院日达到控制目标。"
                    ws['a24'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a24'].font = Font(bold=False, size=12)
                    ws['a25'] = "总分"
                    ws['a25'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a25'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10 / 8, 2))
                    ws['a26'] = "    贵科本考核周期总得分为 " + fen + " 分。（8项指标原始得分）"
                    ws['a26'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a26'].font = Font(bold=False, size=12)

                # 新生儿科
                elif dept == '新生儿科':
                    print('正在处理 %s 数据' % dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text=False)
                    ws['a1'].font = Font(bold=True, size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list['儿科']['质量安全管理小组得分']) \
                               + " 分。（按儿科计，满分10分）" + " 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                                                "有科室质量与安全管理工作计划并实施。" \
                                                "有科室质量与安全工作制度并落实。" \
                                                "有科室质量与安全管理的各项工作记录。" \
                                                "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                                                "对本科室质量与安全指标进行资料收集和分析。" \
                                                "能够运用质量管理方法与工具进行持续质量改进。" \
                                                "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                               "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                               "每百张床位年报告≥20 件。" \
                               "对医疗安全（不良）事件有分析，采取防范措施。" \
                               "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、大型检查阳性率得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期大型检查阳性率得分为 " + str(round(depts_score_list[dept]['大型检查阳性率得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：大型医用设备检查阳性率为三级公立医院绩效考核指标11。" \
                                "对已经购置的大型医用设备使用情况、使用效果应定期评价，" \
                                "以充分发挥其在诊疗中的优势作用，促进大型医用设备科学配置和合理使用。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、门诊病历书写率得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['门诊病历书写率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a12'] = "    贵科本考核周期门诊病历书写率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第一条首诊负责制度要求首诊医师应当作好医疗记录，保障医疗行为可追溯。" \
                                "第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。三甲条款 4.27.2.2 要求为每一位门诊、急诊患者建立就诊记录或急诊留观病历。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、门诊病历甲级率得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['门诊病历甲级率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a14'] = "    贵科本考核周期门诊病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.2要求质量管理相关部门、病案科以及临床各科对病历书写规范进行监督检查，对存在问题与缺陷提出整改措施。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "七、住院病历甲级率得分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院病历甲级率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a16'] = "    贵科本考核周期住院病历甲级率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.4.2要求院科两级落实整改措施，持续改进病历质量，年度住院病案总检查数占总住院病案数≥70%，病历甲级率≥90%，无丙级病历。"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)
                    ws['a17'] = "八、住院病历按时归档率得分"
                    ws['a17'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a17'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['住院按时归档率'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a18'] = "    贵科本考核周期住院病历按时归档率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：医疗质量安全核心制度中第十五条病历管理制度中要求医疗机构病历书写应当做到客观、真实、准确、及时、完整、规范。" \
                                "三甲条款4.27.2.6要求患者出院后，住院病历在 2 个工作日之内回归病案科达≥95％，在 7 个工作日内回归病案科 100%。"
                    ws['a18'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a18'].font = Font(bold=False, size=12)
                    ws['a19'] = "九、临床路径入径率得分"
                    ws['a19'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a19'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['临床路径入径率得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a20'] = "    贵科本考核周期临床路径入径率得分为 " + fen \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款要求医院将开展临床路径与单病种质量管理作为推动医疗质量持续改进的重点项目，规范临床诊疗行为的重要内容之一；" \
                                "对符合进入临床路径标准的患者达到入组率≥50%，入组完成率≥70%。"
                    ws['a20'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a20'].font = Font(bold=False, size=12)
                    ws['a21'] = "十、CMI得分"
                    ws['a21'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a21'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['CMI得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a22'] = "    贵科本考核周期CMI得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，增长为正分，满分10分，下降为扣分，上限-10分）" + \
                                " 说明：根据《广东省卫生健康委关于进步完善三级医院(综合医院、专科医院)等级评审工作的通知》要求，" \
                                "三级综合医院首次申报评审，申报前近一年，三级公立医院绩效考核排名、CMI排名均需在大湾区前75%以内。"
                    ws['a22'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a22'].font = Font(bold=False, size=12)
                    ws['a23'] = "十一、平均住院日得分"
                    ws['a23'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a23'].font = Font(bold=True, size=14)
                    try:
                        fen = str(round(depts_score_list[dept]['平均住院日得分'], 2))
                    except KeyError:
                        fen = '无需考核'
                    ws['a24'] = "    贵科本考核周期平均住院日得分为 " + fen \
                                + " 分。（按本周期值与上一年度值比较计分，降低为正分，满分10分，上升为扣分，上限-10分）" + \
                                " 说明：三甲条款4.5.7.4(核心条款)要求：医院对各临床科室出院患者平均住院日有明确的要求。" \
                                "有缩短平均住院日的具体措施。相关管理人员与医师均知晓缩短平均住院日的要求，并落实各项措施。平均住院日达到控制目标。"
                    ws['a24'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a24'].font = Font(bold=False, size=12)
                    ws['a25'] = "总分"
                    ws['a25'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a25'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10 / 9, 2))
                    ws['a26'] = "    贵科本考核周期总得分为 " + fen + " 分。（9项指标原始得分）"
                    ws['a26'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a26'].font = Font(bold=False, size=12)

                # 麻醉科
                elif dept == '麻醉科' :
                    print('正在处理 %s 数据' %dept)
                    ws['a1'].alignment = Alignment(horizontal='center', wrap_text = False)
                    ws['a1'].font = Font(bold=True,size=20)
                    ws['a1'] = dept + str(start.year) + "年" + str(start.month) + "月" + "质量考核反馈单"
                    ws['a3'] = "一、科室质量与安全管理小组"
                    ws['a3'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a3'].font = Font(bold=True, size=14)
                    ws['a4'] = "    贵科本考核周期质量安全管理小组得分为" + str(depts_score_list[dept]['质量安全管理小组得分']) \
                               + " 分。（满分10分）"+" 说明：三甲条款4.1.1.3要求，各科室要有科室质量与安全管理小组，科主任为第一责任人。" \
                               "有科室质量与安全管理工作计划并实施。" \
                               "有科室质量与安全工作制度并落实。" \
                               "有科室质量与安全管理的各项工作记录。" \
                               "对科室质量与安全进行定期检查，并召开会议，提出改进措施。" \
                               "对本科室质量与安全指标进行资料收集和分析。" \
                               "能够运用质量管理方法与工具进行持续质量改进。" \
                               "科室质量与安全水平持续改进，成效明显。"
                    ws['a4'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a4'].font = Font(bold=False, size=12)
                    ws['a5'] = "二、危急值规范处理得分"
                    ws['a5'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a5'].font = Font(bold=True, size=14)
                    ws['a6'] = "    贵科本考核周期危急值规范处理得分为" + str(depts_score_list[dept]['危急值得分']) \
                               + " 分。（满分10分）" + " 说明：三甲条款3.6.2.1要求严格执行“危急值”报告制度与流程。（核心条款）" \
                                                "接获危急值报告的医护人员应完整、准确记录患者识别信息、危急值内容、和报告者的信息，" \
                                                "按流程复核确认无误后，及时向经治或值班医师报告，并做好记录。医师接获危急值报告后应及时追踪、处置并记录。"
                    ws['a6'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a6'].font = Font(bold=False, size=12)
                    ws['a7'] = "三、不良事件主动报告得分"
                    ws['a7'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a7'].font = Font(bold=True, size=14)
                    ws['a8'] = "    贵科本考核周期不良事件主动报告得分为" + str(depts_score_list[dept]['不良事件得分']) \
                               + " 分。（满分10分，每主动报告1件不良事件得5分，总分不超过10分）" + \
                               " 说明：三甲条款3.9.1.1 有主动报告医疗安全（不良）事件的制度与工作流程。（核心条款）" \
                                                "建立院内网络医疗安全（不良）事件直报系统及数据库。" \
                                                "每百张床位年报告≥20 件。" \
                                                "对医疗安全（不良）事件有分析，采取防范措施。" \
                                                "持续改进安全（不良）事件报告系统的敏感性，有效降低漏报率。"
                    ws['a8'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a8'].font = Font(bold=False, size=12)
                    ws['a9'] = "四、术前访视得分"
                    ws['a9'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a9'].font = Font(bold=True, size=14)
                    ws['a10'] = "    贵科本考核周期术前访视得分为 " + str(round(depts_score_list[dept]['术前访视得分'],2)) \
                               + " 分。（满分10分）" + \
                               " 说明：三甲条款4.7.2 要求实行患者麻醉前病情评估制度，制订治疗计划/方案，风险评估结果记录在病历中。"
                    ws['a10'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a10'].font = Font(bold=False, size=12)
                    ws['a11'] = "五、麻醉复苏得分"
                    ws['a11'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a11'].font = Font(bold=True, size=14)
                    ws['a12'] = "    贵科本考核周期麻醉复苏得分为 " + str(round(depts_score_list[dept]['麻醉复苏得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款4.7.5（核心条款） 要求有麻醉后复苏室，管理措施到位，实施规范的全程监测，记录麻醉后患者的恢复状态，防范麻醉并发症的措施到位。"
                    ws['a12'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a12'].font = Font(bold=False, size=12)
                    ws['a13'] = "六、术后访视得分"
                    ws['a13'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a13'].font = Font(bold=True, size=14)
                    ws['a14'] = "    贵科本考核周期术后访视得分为 " + str(round(depts_score_list[dept]['术后访视得分'], 2)) \
                                + " 分。（满分10分）" + \
                                " 说明：三甲条款4.7.6 要求建立术后、慢性疼痛、癌痛患者的镇痛治疗管理的规范与流程，能有效地执行。"
                    ws['a14'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a14'].font = Font(bold=False, size=12)
                    ws['a15'] = "总分"
                    ws['a15'].alignment = Alignment(horizontal='left', wrap_text=False)
                    ws['a15'].font = Font(bold=True, size=16)
                    fen = str(round(depts_score_list[dept]['总分'], 2))
                    # standard_fen = str(round(float(fen) * 10/ 6,2))
                    ws['a16'] = "    贵科本考核周期总得分为 " + fen + " 分。（6项指标原始得分）"
                    ws['a16'].alignment = Alignment(horizontal='left', wrap_text=True)
                    ws['a16'].font = Font(bold=False, size=12)

            flash('质量考核反馈表已成功保存！')
            wb.save(os.path.join('./temp/', '质量考核反馈表.xlsx'))
            wb.close()


        for dept in depts:
            if dept[0] in ('中医科','体检科','体检科盐田','口腔科','急诊科','急诊科盐田','皮肤科','精神科','针灸推拿科'):
                queryOP(dept[0])
            elif dept[0] in ('中西医结合心血管内科','中西医结合老年病科','传染科', '儿科','内分泌科','呼吸内科',
                             '康复医学科','消化内科','神经内科','肾内科','肿瘤科','血液内科'):
                queryIP(dept[0])
            elif dept[0] in ('中西医结合肛肠科','妇产科盐田','妇科','普外科','普外科盐田','泌尿外科','甲乳外科'
                             ,'眼科','神经外科','耳鼻喉科','胸外科','骨科','骨科盐田'):
                querySurgery(dept[0])
            elif dept[0] in ('检验科','病理科','放射影像科','超声影像科'):
                queryExam(dept[0])
            elif dept[0] == '新生儿科':
                queryneonatus('新生儿科')
            elif dept[0] == '重症医学科':
                queryicu('重症医学科')
            elif dept[0] == '产科':
                queryObste('产科')
            elif dept[0] == '麻醉科':
                queryAnes('麻醉科')
            else:
                pass


        print(depts_score_list)
        quality_score_report(depts_score_list)

    return render_template('query.html', deptclass=deptclass,dataqs=dataqs,dataos=dataos, dataps=dataps,dataics=dataics,
                           datais=datais,datass=datass,dataes = dataes,dataas=dataas,scores=scores,queryform=queryform)






if __name__ == '__main__':
    app.run()