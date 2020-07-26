import pandas as pd
import datetime
import smtplib,os
os.chdir(r'C:\Users\Daya\PycharmProjects\BirthdayWisher')
# os.mkdir('Testing')


_gmailID='YOUR EMAIL ID'
_gmailPwd='YOUR EMAIL PASSWORD'

def sendEmail(to,subject,message,image):
    print(f'Email to {to} sent with subject {subject} and message {message} ')
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(_gmailID,_gmailPwd)
    s.sendmail(_gmailID,to,f'Subject: {subject}\n\n {message}\n\n {image}\n\n Best Regards,\n Dayasagar')
    s.quit()

if __name__ == '__main__':

    # sendEmail(_gmailID,'subject','message test')

    df = pd.read_excel('data.xlsx')

    yearNow=datetime.datetime.now().strftime('%Y')
    today = datetime.datetime.now().strftime('%d-%m')

    writeIndex = []

    for index,item in df.iterrows():
        bday=item['Birthday'].strftime('%d-%m')
        if (today==bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'],'HAPPY BIRTHDAY',item['Dialogue'],item['Image'])
            writeIndex.append(index)
    # print(writeIndex)
    if writeIndex is not None:

        for i in writeIndex:
            yr=df.loc[i,'Year']
            df.loc[i,'Year']=str(yr)+','+str(yearNow)
            # print(df.loc[i,'Year'])
        # print(df)
        df.to_excel('data.xlsx',index=False)


