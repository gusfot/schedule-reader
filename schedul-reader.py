from openpyxl import load_workbook
import pymysql


def _read(filepath):
    load_wb = load_workbook(filepath, data_only=True)
    return load_wb


def regist_schedule_tmp(filepath):
    _excel2Dic(filepath)

# 엑셀파일에서 읽어서, schedule_tmp 테이블로 입력
def _excel2Dic(filepath):

    # load_wb = _read("C:/Users/gusfo/OneDrive/2019.11.24-2020.01.03.xlsx")
    load_wb = _read(filepath)
    dic_list = []

    load_ws = load_wb['Sheet1']
    rows = load_ws['A1':'B300']

    for row in rows:
        if row[0].value != None:
            d = row[0].value
            n = row[1].value
            dic = {'date': d, 'name': n}
            dic_list.append(dic)
        else:
            break

    # print(dic_list)

    for item in dic_list:

        if item['name'] != None:
            connect = pymysql.connect(host='52ch.org', user='root', password='jwchjwch',
                                      db='message', charset='utf8')
            try:
                with connect.cursor() as cursor:
                    cur = connect.cursor()
                    sql = "insert into schedule_tmp(sch_date, name) value (\'" + item['date'] + "\',\'" + item[
                        'name'] + "\')"
                    print(sql)
                    cur.execute(sql)
                    connect.commit()
            finally:
                connect.close()


# schedule_tmp 데이터로 예약문자에 등록하기
def booking():
    connect = pymysql.connect(host='52ch.org', user='root', password='jwchjwch',
                              db='message', charset='utf8')

    try:
        with connect.cursor() as cursor:
            cur = connect.cursor()
            sql = ""
            sql += "insert into message_schedule(calendar_seq, schedule_date, schedule_content, notifications) "
            sql += "select "
            sql += "    1 calendar_seq, concat(s.sch_date, \' 04:30:00\') as schedule_date "
            sql += "    , \'{{title}}{{name}} {{grade}}님.내일{{schedule_date}}담당자입니다.새벽을깨우는안수집사회!\' as schedule_content "
            sql += "    , json_object(\'msg\', json_array(json_object(\'name\',m.name, \'grade\', m.grade, \'mobile\', m.mobile, \'user_id\', m.user_id))) as notifications "
            sql += "from schedule_tmp s "
            sql += "inner join member m ON m.user_id = s.name "
            sql += "where s.sch_date >= CURRENT_DATE()"

            print(sql)
            cur.execute(sql)
            connect.commit()
    finally:
        connect.close()



if __name__ == '__main__':
    # filepath = "C:/Users/gusfo/OneDrive/2019.11.24-2020.01.03.xlsx"
    filepath = "C:\\Users\\user\\OneDrive\\오병이어교회\\2020.04.12-2020.05.16.xlsx"
    regist_schedule_tmp(filepath)
    booking()
