
import os
import sys
import pandas as pd


# 경로 파일 체크하기
start_path = 'C:/Users/COM/Desktop/bplace/statics_prj/subway/gwangju'
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


# vaccent dataframe 만들기
df = pd.DataFrame(columns=['region','station_nm','surv_dt','week_dt','inout_type','hour_5_psn_cnt','hour_6_psn_cnt','hour_7_psn_cnt','hour_8_psn_cnt','hour_9_psn_cnt','hour_10_psn_cnt','hour_11_psn_cnt',
                           'hour_12_psn_cnt','hour_13_psn_cnt','hour_14_psn_cnt','hour_15_psn_cnt','hour_16_psn_cnt','hour_17_psn_cnt','hour_18_psn_cnt','hour_19_psn_cnt','hour_20_psn_cnt',
                           'hour_21_psn_cnt','hour_22_psn_cnt','hour_23_psn_cnt','hour_24_psn_cnt'])

# 엑셀 불러오기 len(subway.index)-1

for file_name in file_list:
    # 파일저장 경로 설정
    file_path = start_path + '/' + file_name
    # sheet name 을 기반으로 dataframe 출력
    for sheet_name_list in sheet_name_lists:
        subway = pd.read_excel(file_path, header=2, sheet_name=sheet_name_list)
        print('현황 = ', sheet_name_list)
        print(subway.info())
        print(subway.head())

        # 데이터 프레임 만들기
        subway_frame = pd.DataFrame({
            'region' : file_name.replace('.xlsx',''),
            'station_nm' : subway['Unnamed: 0'],
            'surv_dt' : pd.to_datetime(subway['Unnamed: 1'].apply(str)),
            'week_dt' : subway['Unnamed: 2'],
            'inout_type' : sheet_name_list.replace('현황',''),
            'hour_5_psn_cnt': subway['5시'],
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
        print(subway_frame)

        if subway_frame.loc[len(subway.index)-1, 'station_nm'] == '합계':
            subway_frame = subway_frame[0:len(subway.index) - 1]
            df = df.append(subway_frame)
        else:
            df = df.append(subway_frame)

        # csv 파일 형태로 저장하기
        df.to_csv(f'C:/Users/COM/Desktop/bplace/statics_prj/csv_file/{file_name.replace(".xlsx","")}.csv', sep=',' ,encoding='utf8', index=False)
