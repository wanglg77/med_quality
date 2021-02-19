from define import *


@app.route('/dept', methods=['GET','POST'])
def dept():
    deptform = Deptform()
    dept_name = deptform.deptname.data
    dept_class = deptform.deptclass.data
    dept_bed = deptform.deptbed.data
    deptifexist = Department.query.filter_by(name=dept_name).first()
    if deptform.submit.data and deptform.is_submitted():
        if deptifexist:
            flash('科室已存在')
        else:
            try:
                new_dept = Department(name=dept_name, classname = dept_class, bednumber = dept_bed)
                db.session.add(new_dept)
                db.session.commit()
                flash('科室添加成功')
            except BaseException as e:
                print(e)
                flash('添加科室失败,回滚数据')
                db.session.rollback()
    elif deptform.submit2.data and deptform.is_submitted():
        if deptifexist:
            try:
                Department.query.filter(Department.name == dept_name).delete()
                # del_dept = Department(name=dept_name, classname = dept_class, bednumber = dept_bed)
                # db.session.delete(del_dept)
                db.session.commit()
                flash('删除科室成功')
            except Exception as e:
                print(e)
                flash('删除科室失败,回滚数据')
                db.session.rollback()
    elif deptform.submit3.data and deptform.is_submitted():
        if deptifexist:
            try:
                # mod_data=Department.query.filter_by(Department.name == dept_name).first()
                # deptifexist.name=dept_name
                deptifexist.classname = dept_class
                deptifexist.bednumber = dept_bed
                db.session.commit()
                flash('修改科室成功')
            except Exception as e:
                print(e)
                flash('修改科室失败,回滚数据')
                db.session.rollback()
    depts = Department.query.all()
    return render_template('dept.html', depts=depts, deptform=deptform)

if __name__ == '__main__':
    app.run()