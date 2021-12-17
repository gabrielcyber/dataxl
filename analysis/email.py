import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from datetime import datetime

class Email:
    def __init__(self):
        self.now_time = lambda: datetime.strftime(datetime.now(), '%H:%M:%S')
        self.__config = EmailConfig()

    def attachments(self, *files):
        for file in files:
            self.__config.email.attach(MIMEImage(open(file, 'rb').read()))

    def edit_email(self, addressees, subject, html_source):
        addressee = ''
        
        if addressee.__class__ == tuple:
            for send in addressees:
                self.addressee += '%s;' % send
        else:
            self.addressee = addressees

        self.__config.email['Subject'] = subject
        self.__config.HTMLBody.attach(MIMEText(html_source, 'html'))

    def login(self, user, password):
        self.__config.login(user, password)
        self.__config.smtp.starttls()
        self.__config.smtp.login(user, password)

    def send(self, automatic_time=None):
            try:
                if automatic_time is not None:
                    while True:
                        now = self.now_time()
                        if now == automatic_time:
                            break
            except BaseException as error:
                raise error
            else:
                self.__config.smtp.sendmail(from_addr=self.__config.user, to_addrs=self.addressee, msg=self.__config.email.as_string())                
                print(f"Email send at {self.now_time()} o'clock, to: {self.addressee}")

class EmailConfig:
    def __init__(self):
        self.email = MIMEMultipart('related')
        self.HTMLBody = MIMEMultipart('alternative')
        self.email.attach(self.HTMLBody)
        self.smtp = smtplib.SMTP(host="smtp.gmail.com", port="587")

    def login(self, user, password):
        self.user = user
        self.password = password
