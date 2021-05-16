
import csv
import os
import sys
import pandas as pd


# 경로 파일 체크하기
start_path = 'C:/Users/COM/Desktop/bplace/statics_prj/subway/daejeon'
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
for sheet_name_list in sheet_name_lists:
    subway = pd.read_excel(file_path, header=2, sheet_name=sheet_name_list)
    print('현황 = ', sheet_name_list)
    print(subway.head())

    # 데이터 프레임 만들기
    subway_frame = pd.DataFrame({
        'region' : file_list[0].replace('.xlsx',''),
        'station_nm' : subway['Unnamed: 0'],
        'surv_dt' : subway['Unnamed: 1'],
        'week_dt' : subway['Unnamed: 2'],
        'inout_type' : sheet_name_list.replace('현황',''),
        'hour_6_psn_cnt' : subway['6시'],
        'hour_7_psn_cnt' : subway['7시'],
        'hour_8_psn_cnt' : subway['8시'],
        'hour_9_psn_cnt' : subway['9시'],
        'hour_10_psn_cnt' : subway['10시'],
        'hour_11_psn_cnt' : subway['11시'],
        'hour_12_psn_cnt' : subway['12시'],
        'hour_13_psn_cnt' : subway['13시'],
        'hour_14_psn_cnt' : subway['14시'],
        'hour_15_psn_cnt' : subway['15시'],
        'hour_16_psn_cnt' : subway['16시'],
        'hour_17_psn_cnt' : subway['17시'],
        'hour_18_psn_cnt' : subway['18시'],
        'hour_19_psn_cnt' : subway['19시'],
        'hour_20_psn_cnt' : subway['20시'],
        'hour_21_psn_cnt' : subway['21시'],
        'hour_22_psn_cnt' : subway['22시'],
        'hour_23_psn_cnt' : subway['23시'],
        'hour_24_psn_cnt' : subway['24시']
        })


    subway_csv = subway_frame[0:len(subway.index)-1]

    subway_join = subway_csv.append(subway_csv)

# csv 파일 형태로 저장하기
subway_join.to_csv(f'C:/Users/COM/Desktop/bplace/statics_prj/csv_file/{file_list[0].replace(".xlsx","")}.csv', sep=',', encoding='euckr')


