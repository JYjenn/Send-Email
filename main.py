import sys
import os
import random
import pandas as pd
from time import sleep
# add env path
original_path = os.getcwd()
main_file_ex_path = os.path.join(sys.argv[0], os.path.pardir)
sys.path.append(main_file_ex_path)
os.chdir(main_file_ex_path)

from utills import load_dataframe_without_blank, extract_unique_names, send_mail, mail_component


if __name__ == '__main__':
    """
    cmd 창에서 입력받을 인자: [1]엑셀파일 [2]담당자메일주소 파일
    [가상환경 파이썬 경로] [main.pyc파일 경로] [입력인자 1] [입력인자 2]
    "C:\\Users\\USER\\Anaconda3\\envs\\venv38\\python.exe" "D:\\JY\\Send Email auto\\main.pyc" (아랫줄 이어서)
    "D:\\JY\\Send Email auto\\test_data\\2022-07-05_send_result.xlsx" "D:\\JY\\Send Email auto\\test_data\\mail_address.xlsx"
    """

    if sys.argv[1] == '--mode=client':
        xlsx_file = 'D:\\JY\\Send Email auto\\test_data\\2022-07-05_send_result.xlsx'  # 입력인자 [1]
        mail_address_file = 'D:\\JY\\Send Email auto\\test_data\\mail_address.xlsx'  # 입력인자 [2]
        os.chdir(original_path)
    else:
        xlsx_file = sys.argv[1]
        mail_address_file = sys.argv[2]


    data_frame = load_dataframe_without_blank(xlsx_file)  # 엑셀파일 데이터프레임 형식으로 읽기
    ## pre-processing (1). Nan값 없애기 & 일자 형식 맞추기
    data_frame.fillna('', inplace=True)
    datetime_col = data_frame['일자'].str.split(' ').apply(lambda x: pd.Series(x))
    date = datetime_col.iloc[:, 0]
    data_frame['일자'] = date
    ## pre-processing (2). 이메일 타입: 결과발송 or 출고
    email_type = data_frame.loc[0, '센터명'].split('(')[1][:-1]
    ## pre-processing (3). 타입에 따라 '일자'명 변경하기
    if email_type == '결과발송':
        data_frame.rename(columns={'일자': '분석일자'}, inplace=True)
    elif email_type == '출고':
        data_frame.rename(columns={'일자': '검진일자'}, inplace=True)
    else:
        print("Please Check Email Type!")

    address_df = load_dataframe_without_blank(mail_address_file, need_preprocessing=False)  # 담당자 메일주소
    centers = extract_unique_names(data_frame, column_name='센터명')  # 데이터프레임 내 센터명 확인 (중복x)

    # 참조자 이메일 명단 (영업팀)
    cc_address = 'jun.k@taewoongmedical.com; duet@taewoongmedical.com; bumseok@taewoongmedical.com'

    for tmp_center in centers:
        try:
            condition = data_frame['센터명'] == tmp_center
            tmp_center_df = data_frame.merge(data_frame[condition], how='inner')
            tmp_center_name = tmp_center.split('(')[0]
            to_address = address_df[address_df['센터명'] == tmp_center_name].values[0][1]  # 받는사람 메일주소 추출
            # 메일 구성 내용
            mail_to, mail_cc, mail_title, mail_content = mail_component(tmp_center_df, tmp_center_name,
                                                                        To=to_address, Cc=cc_address,
                                                                        mail_type=email_type)
            # 메일 발송
            send_mail(mail_to, mail_cc, mail_title, mail_content)
        except IndexError:
            pass
        # ip 부하가 걸릴 수 있으므로 sleep 1~3초 중 랜덤하게 걸어주기
        sleep(random.uniform(1,3))
