import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import schedule
import time
import datetime
import numpy as np
import os

def get_form_data():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    sheet = client.open("NAME_SHEET").sheet1

    # Extract and print all of the values
    list_of_hashes = sheet.get_all_records()
    df = pd.DataFrame(list_of_hashes)

    df['Exchange Score'] = (df['Exchange Score'].replace('\.', '', regex=True).replace(',', '.', regex=True).astype(float))

    #df.to_csv(path_or_buf='/Users/eliogruttadauria/Desktop/googlesheet_exchange.csv')

    df.drop_duplicates('Student ID', keep='last', inplace=True)

    return df

def get_uni_data():
    university_data = pd.read_csv('uni_data.csv', sep=';')
    university_data.dropna(how='all', axis=1, inplace=True)
    university_data.dropna(how='all', axis=0, inplace=True)

    return university_data


def make_ranking():

    form_data = get_form_data()
    university_data = get_uni_data()

    ranking = pd.DataFrame(columns=['Student ID', 'Exchange Score', 'University'])

    gpa_list = form_data[['Exchange Score']].values
    gpa_list = gpa_list.reshape(-1)
    gpa_sorted = np.sort(gpa_list)[::-1]

    uni = university_data[['University']].values

    for gpa in gpa_sorted:

        i = np.where(form_data == gpa)[0][0]

        ranking.loc[i, 'Student ID'] = form_data.loc[i, 'Student ID']
        ranking.loc[i, 'Exchange Score'] = form_data.loc[i, 'Exchange Score']

        first_uni = form_data.loc[i, 'Your first choice:']
        second_uni = form_data.loc[i, 'Your second choice:']
        third_uni = form_data.loc[i, 'Your third choice:']

        try:
            j = np.where(uni == first_uni)[0][0]
        except:
            pass
        try:
            k = np.where(uni == second_uni)[0][0]
        except:
            pass
        try:
            h = np.where(uni == third_uni)[0][0]
        except:
            pass

        flag = False

        try:
            if university_data.loc[j, 'Current Slots'] < university_data.loc[j, 'Slots']:
                ranking.loc[i, 'University'] = first_uni
                university_data.loc[j, 'Current Slots'] += 1

            elif university_data.loc[k, 'Current Slots'] < university_data.loc[k, 'Slots']:
                ranking.loc[i, 'University'] = second_uni
                university_data.loc[k, 'Current Slots'] += 1

            elif university_data.loc[h, 'Current Slots'] < university_data.loc[h, 'Slots']:
                ranking.loc[i, 'University'] = third_uni
                university_data.loc[h, 'Current Slots'] += 1

            else:
                ranking.loc[i, 'University'] = 'OUT'
        except:
            flag = True
            pass

        if flag:
            ranking.loc[i, 'University'] = 'OUT'

    try:
        os.remove('/Users/eliogruttadauria/Desktop/ranking_exchange.xlsx')
    except:
        pass

    ranking.to_excel('/Users/eliogruttadauria/Desktop/ranking_exchange.xlsx')


def send_message():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    fromaddr = "e.gruttadauria98@gmail.com"
    toaddr = "e.gruttadauria98@gmail.com"

    # instance of MIMEMultipart
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    current_date = datetime.datetime.today().strftime('%Y-%m-%d')
    msg['Subject'] = "Updated_Rankings " + str(current_date)

    # string to store the body of the mail
    body = "This mail contains the new updated ranking for the exchange"

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    filename = "ranking_exchange.xlsx"
    attachment = open("/Users/eliogruttadauria/Desktop/ranking_exchange.xlsx", "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload((attachment).read())

    # encode into base64
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(fromaddr, "PASSWORD_MAIL")

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromaddr, toaddr, text)

    # terminating the session
    s.quit()


def job():
    make_ranking()
    send_message()
    return

schedule.every().day.at("08:00").do(job)

while True:
    schedule.run_pending()
    time.sleep(30)
