
import os
import sys
import pandas as pd


# 경로 파일 체크하기
start_path = 'C:/Users/COM/Desktop/bplace/statics_prj/subway/busan'
# 파일 리스트
file_list = os.listdir(start_path)
print("file_list = ", file_list)

# 파일 최종경로
file_path = start_path +'/'+ file_list[0]
print("file_path = ", file_path)


# 시트 이름, 시트 개수 등 파일 정보 print
sheet_name = pd.ExcelFile(file_path)
sheet_name_lists = sheet_name.sheet_names
print("시트 이름 리스트 = ", sheet_name_lists)
print("시트 개수 = ", len(sheet_name_lists))


# 엑셀 불러오기
for file_name in file_list:
    # 파일저장 경로 설정
    file_path = start_path + '/' + file_name
    # sheet name 을 기반으로 dataframe 출력

    subway = pd.read_excel(file_path)
    # 데이터 프레임 만들기
    subway_frame = pd.DataFrame({
        'region': file_name.replace('.xlsx', ''),
        'station_nm': subway['역명'],
        'surv_dt': pd.to_datetime(subway['날짜'].apply(str)),
        'week_dt': subway['요일'],
        'inout_type': subway['승하차'],
        'hour_1_psn_cnt': subway['00-01'],
        'hour_2_psn_cnt': subway['01-02'],
        'hour_3_psn_cnt': subway['02-03'],
        'hour_4_psn_cnt': subway['03-04'],
        'hour_5_psn_cnt': subway['04-05'],
        'hour_6_psn_cnt': subway['05-06'],
        'hour_7_psn_cnt': subway['06-07'],
        'hour_8_psn_cnt': subway['07-08'],
        'hour_9_psn_cnt': subway['08-09'],
        'hour_10_psn_cnt': subway['09-10'],
        'hour_11_psn_cnt': subway['10-11'],
        'hour_12_psn_cnt': subway['11-12'],
        'hour_13_psn_cnt': subway['12-13'],
        'hour_14_psn_cnt': subway['13-14'],
        'hour_15_psn_cnt': subway['14-15'],
        'hour_16_psn_cnt': subway['15-16'],
        'hour_17_psn_cnt': subway['16-17'],
        'hour_18_psn_cnt': subway['17-18'],
        'hour_19_psn_cnt': subway['18-19'],
        'hour_20_psn_cnt': subway['19-20'],
        'hour_21_psn_cnt': subway['20-21'],
        'hour_22_psn_cnt': subway['21-22'],
        'hour_23_psn_cnt': subway['22-23'],
        'hour_24_psn_cnt': subway['23-24']
    })
    print(subway_frame)

    # csv 파일 형태로 저장하기
    subway_frame.to_csv(f'C:/Users/COM/Desktop/bplace/statics_prj/csv_file/{file_name.replace(".xlsx", "")}.csv', sep=',',encoding='utf8', index=False)




