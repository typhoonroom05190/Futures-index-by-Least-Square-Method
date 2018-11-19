name = input('請輸入你的檔案名稱:')
try:
    wb = openpyxl.load_workbook(name + '.xlsx')
except Exception as X:
    print('找不到該檔案')
    response = input('請問你要新建一個檔案嗎?\n回答:')
    if response == 'yes':
        wb = openpyxl.Workbook()
        wb.save(name + '.xlsx')
        openpyxl.load_workbook(name + '.xlsx')
    else:
        sys.exit()
