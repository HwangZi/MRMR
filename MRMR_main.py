import openpyxl
from flask import Flask, render_template

# 엑셀 파일 열기
workbook = openpyxl.load_workbook('C:/Users/황지혁/Desktop/MRMR/flask_excel.xlsx')
sheet = workbook.active

# 열의 개수와 행의 개수
num_rows = 1000
num_cols = 7

# 딕셔너리
data_dict = {}
accuracyList = []


# 각 col에 대해 반복(정확도 col 제외한 나머지 col에 대해 가장 최신 정보로 key-value(sec-value) 딕셔너리 완성)
for col in range(2, num_cols):
    # 열 이름 가져오기
    col_name = sheet.cell(row=1, column=col).value

    # 딕셔너리 초기화
    column_dict = {}

    # 각 행에 대해 반복
    for row in range(2, num_rows + 1):
        # 시간 정보 가져오기
        time = sheet.cell(row=row, column=1).value

        # 값 가져오기
        value = sheet.cell(row=row, column=col).value

        # 값이 존재하면 딕셔너리에 저장, timeSet에 발생된 시간 추가
        if value is not None:
            column_dict = {}
            column_dict[time] = value

        # 해당 열이 마지막 행에 도달하면 반복 종료
        if row == num_rows:
            break

    # 열에 대한 딕셔너리 저장
    data_dict[col_name] = column_dict

col = num_cols + 1
for row in range(2, num_rows + 1):
    time = sheet.cell(row=row, column=1).value
    value = sheet.cell(row=row, column=col).value

    if value is not None:
        accuracyList.append(value)

    if row == num_rows:
            break

# Flask 앱 생성
app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html', data=data_dict, accuracy=accuracyList)

@app.route('/help')
def help():
    return render_template('help.html')

@app.route('/update')
def update():
    workbook = openpyxl.load_workbook('C:/Users/황지혁/Desktop/MRMR/flask_excel.xlsx')
    sheet = workbook.active

    # 열의 개수와 행의 개수
    num_rows = 1000
    num_cols = 7

    # 딕셔너리 초기화
    data_dict = {}
    accuracyList = []

    # 각 col에 대해 반복(정확도 col 제외한 나머지 col에 대해 가장 최신 정보로 key-value(sec-value) 딕셔너리 완성)
    for col in range(2, num_cols):
        # 열 이름 가져오기
        col_name = sheet.cell(row=1, column=col).value

        # 딕셔너리 초기화
        column_dict = {}

        # 각 행에 대해 반복
        for row in range(2, num_rows + 1):
            # 시간 정보 가져오기
            time = sheet.cell(row=row, column=1).value

            # 값 가져오기
            value = sheet.cell(row=row, column=col).value

            # 값이 존재하면 딕셔너리에 저장, timeSet에 발생된 시간 추가
            if value is not None:
                column_dict = {}
                column_dict[time] = value

            # 해당 열이 마지막 행에 도달하면 반복 종료
            if row == num_rows:
                break

        # 열에 대한 딕셔너리 저장
        data_dict[col_name] = column_dict

    col = num_cols + 1
    for row in range(2, num_rows + 1):
        time = sheet.cell(row=row, column=1).value
        value = sheet.cell(row=row, column=col).value

        if value is not None:
            accuracyList.append(value)

        if row == num_rows:
            break

    return render_template('index.html', data=data_dict, accuracy=accuracyList)

if __name__ == '__main__':
    app.run()