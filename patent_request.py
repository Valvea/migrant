from PySide2 import QtCore
from PySide2.QtGui import QImage
import time
import requests as req
from bs4 import BeautifulSoup
import json
from threading import current_thread


class Message(QtCore.QObject):
    captcha = QtCore.Signal(QImage,int)
    response = QtCore.Signal(dict)


class Check_patent:
    def __init__(self):
        self.main_url = 'https://servicesmmc.mos.ru'
        self.url = 'https://servicesmmc.mos.ru/mmc-status/ndfl-payment-status.html'
        self.headers = {
            "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                          ' (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
            'Content-Type': 'application/json'}

        self.url_post = 'https://servicesmmc.mos.ru/mmc-status/rest/ndfl-payment-status'

        self.signal = Message()
        self.status = False
        self.captcha=None
        self.row=None


    def lock(self):
        while not self.status:
            time.sleep(1.5)
            print('waiting string','rows:',self.row,'current_thread',current_thread(),'obj',self)

    def unlock(self, captcha_string,row):
        print(captcha_string,row)
        if  row == self.row:
            self.captcha = captcha_string
            self.status = True


    def get_patent_status(self, data):

        self.row=data.name
        with req.Session() as session:

            resp = session.get(self.url, headers=self.headers)
            bs = BeautifulSoup(resp.text, 'html.parser')
            captcha_url = self.main_url + bs.find_all(id='captchaImage')[0].get('src')

            with open(f"{data['Номер патента']}.png", 'wb') as c:
                c.write(session.get(captcha_url).content)
            # self.data_ = data
            try:
                captcha_Image = QImage(f"{data['Номер патента']}.png")

                self.signal.captcha.emit(captcha_Image,data.name)

            except Exception as e:
                print("Не удалось прочитать каптчу")

            self.lock()

            load = {'findByPassport': False,
                    'captcha': self.captcha,
                    "patentSerial": int(data['Серия патента']),
                    "patentNumber": int(data['Номер патента']),
                    'falseissueDate': str(data['Дата выдачи'])}
            print(load)
            resp = session.post(self.url_post, headers=self.headers, data=json.dumps(load))
            answer = dict(resp.json())
            answer['captcha_row']=self.row
            print(answer)
            self.signal.response.emit(answer)

