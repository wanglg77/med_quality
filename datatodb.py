from datatodb_def import *


print('='*40)
print('医疗质量数据自动化处理系统:')
print('='*40)
print('0. 病历质控统计(必须先0后1)')
print('1. 病历质控结果数据导入')
print('2. 危急值数据导入')
print('3. 不良事件数据导入')
print('4. HIS数据导入')
print('5. QCC数据导入')
print('6. 院感药剂数据导入')
print('7. 病案系统数据导入')
print('8. DRG数据导入')
print('9. 处理医技数据')
print('10. 所有空数据处理')
print('q. 退出')
print('=' *40)
key = input('请按相应数字键选择功能: ')
if key == '0':
    recstastic()
elif key == '1':
    choice=input('0. 病历质控统计已经运行过了吗???(Y/N)')
    if choice == 'y':
        rectodb()
    elif choice=='n':
        key = input('请重新按相应数字键选择功能//')
elif key == '2':
    wjztodb()
elif key == '3':
    blsjtodb()
elif key == '4':
    histodb()
elif key == '5':
    qcctodb()
elif key == '6':
    pharminfetodb()
elif key == '7':
    batodb()
elif key == '8':
    drgtodb()
elif key == '9':
    yijitodb()
elif key == '10':
    nullto0()
elif key == 'q':
    quit()