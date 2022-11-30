import openpyxl
import numpy as np
import pandas as pd
from datetime import date
import win32com.client as win32


def load_dataframe_without_blank(xlsx_path, need_preprocessing=True):
    if need_preprocessing:
        load_df = pd.read_excel(xlsx_path, header=None, names=['센터명', '일자', '성명', '사업장'], dtype=str)
        df_del_1st_row = load_df.drop([0], axis=0)
        df_del_1st_row.reset_index(drop=True, inplace=True)
        df_no_blank = df_del_1st_row.apply(lambda x: x.str.strip(), axis=1)
    else:
        load_df = pd.read_excel(xlsx_path)
        df_no_blank = load_df.apply(lambda x: x.str.strip(), axis=1)

    return df_no_blank


def extract_unique_names(df, column_name='센터명'):
    names = df[column_name]
    # names = names.apply(lambda x: x.split('(')[0])
    unique_names = names.unique()
    return unique_names


def mail_component(Center_df, Center_name,
                   To='wldus5010@naver.com', Cc='wldus081546@gmail.com',
                   mail_type='결과발송'):
    mail_to = To
    mail_cc = Cc

    Center_count = len(Center_df)
    today_date = date.today().strftime("%Y-%m-%d")
    mail_title = '[태웅메디칼-{}] {}일 유슬립 {}건'.format(Center_name, today_date, mail_type)

    body_content1 = """
                    <body>
                      <p>안녕하세요, 태웅메디칼 입니다.<br>
                         <br>
                      </p>
                    </body>
                    """
    body_content2 = "다음과 같이 총 {}건이 {} 되었습니다.".format(Center_count, mail_type) + """
    <body>
      <p> <br>
          <br>
      </p>
    </body>
    """
    body_content3_raw = Center_df.iloc[:, 1:]
    body_content3 = body_content3_raw.to_html(index=False, justify='center', col_space=100)
    body_content4 = """
                    <body>
                      <p> <br>
                          <br>
                         감사합니다.<br>
                      </p>
                    </body>  
                    """

    mail_content = body_content1 + body_content2 + body_content3 + body_content4
    return mail_to, mail_cc, mail_title, mail_content


def send_mail(to, cc, title, content, contain_cc=True):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    if contain_cc:
        mail.To = to
        mail.CC = cc
        mail.Subject = title
        mail.HTMLBody = content
    else:
        mail.To = to
        mail.Subject = title
        mail.HTMLBody = content

    mail.Send()
