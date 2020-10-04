from parser_email.settings import GMAIL_IMAP_HOST, GMAIL_IMAP_PORT, GMAIL_IMAP_LOGIN, GMAIL_IMAP_PASS, BASE_DIR
import imaplib
import email as email_lib
from email.header import decode_header, make_header
# from dateutil.parser import parse as date_parse
# import html2text
import os, csv
import xml.etree.ElementTree as ET
from .models import Model, Asset, Meter
import datetime


class ProcessMail:
    def __init__(self):
        self.imap_server = imaplib.IMAP4_SSL(host=GMAIL_IMAP_HOST, port=GMAIL_IMAP_PORT)
        self.newemails = None

    def connect(self):
        status, details = self.imap_server.login(GMAIL_IMAP_LOGIN, GMAIL_IMAP_PASS)
        assert status == 'OK', 'IMAP login problem'
        status, details = self.imap_server.select()
        assert status == 'OK', 'IMAP select problem'
        _, message_numbers_raw = self.imap_server.search(None, 'ALL')

        all_email_list = message_numbers_raw[0].split()
        # processed_emails = list(Emails.objects.values_list('email_id', flat=True))
        # self.newemails = [email for email in all_email_list if int(email) not in processed_emails]

    def fetch_mail(self, message_number):
        msg_status, msg = self.imap_server.fetch(message_number, '(RFC822)')
        assert msg_status == 'OK', 'fetching email problem' + message_number

        message = email_lib.message_from_bytes(msg[0][1])
        email_id = int(message_number)
        from_email = str(make_header(decode_header(message["from"])))
        email_subj = str(make_header(decode_header(message['subject']))) if message['subject'] else ''
        create_date = date_parse(message['Date'][:31])
        body = list()
        attachments = dict()

        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition'):
                filename = part.get_filename()
                attachments.update({'fileName': filename})
            else:
                bytes_part = part.get_payload(decode=True)
                charset = part.get_content_charset('iso-8859-1')
                part_body = html2text.html2text(bytes_part.decode(charset, 'replace'))
                if part_body:
                    body.append(part_body)
                print(type(body), body)

        return {'email_id': email_id,
                'from_email': from_email,
                'email_subj': email_subj,
                'create_date': create_date,
                'body': ''.join(body) if body else '',
                'attachments': attachments}

    def fetch_new_email(self):
        self.imap_server.select("inbox")
        result, data = self.imap_server.search(None, "from", '"(service@etc.com.ua)"')
        msgs = self.get_email(data)
        for ms in msgs:
            raw_email = ms[0][1]
            email_message = email_lib.message_from_bytes(raw_email)
            from_email = str(make_header(decode_header(email_message["from"])))
            email_subj = str(make_header(decode_header(email_message['subject']))) if email_message['subject'] else ''
            CWW, EDA = False, False
            if "CWW" in email_subj:
                CWW = True
            elif "EDA" in email_subj:
                EDA = True
            arrayFileCSV, arrayFileXML = [], []
            for part in email_message.walk():
                if part.get_content_type() == "text/csv":
                    file = part.get_filename()
                    filePath = os.path.join(BASE_DIR, file)
                    if not os.path.isfile(filePath):
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                    arrayFileCSV.append(filePath)
                elif part.get_content_type() == "text/xml":
                    file = part.get_filename()
                    filePath = os.path.join(BASE_DIR, file)
                    if not os.path.isfile(filePath):
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                    arrayFileXML.append(filePath)
            for file in arrayFileCSV:
                if CWW:
                    with open(file, mode='r') as f:
                        reader = csv.reader(f, delimiter=',')
                        array = list()
                        head_row = None

                        for row in reader:
                            if 'Printer Model' in row:
                                head_row = row
                            if not head_row is None and not head_row == row:
                                data = dict()
                                for i in range(0, len(row)):
                                    data[head_row[i]] = row[i]
                                array.append(data)
                        for i in array:
                            if not i['Printer Model'] == "" and 'Black Impressions' in i:
                                p_model = i['Printer Model']
                                try:
                                    p_model = Model.objects.get(model_name=p_model)
                                except Model.DoesNotExist:
                                    p_model = Model.objects.create(model_name=p_model)
                                try:
                                    asset = Asset.objects.get(serialnumber=i['Serial Number'])
                                except Asset.DoesNotExist:
                                    asset = Asset.objects.create(serialnumber=i['Serial Number'], model=p_model,
                                                                 location=i['Printer Location'])
                                try:
                                    meter = Meter.objects.get(asset=asset)
                                except Meter.DoesNotExist:
                                    try:
                                        bwa4 = int(i['Black Impressions'])
                                    except:
                                        bwa4 = 0
                                    try:
                                        cola4 = int(i['Color Impressions'])
                                    except:
                                        cola4 = 0
                                    try:
                                        bwa3 = int(i['Black Large Impressions'])
                                    except:
                                        bwa3 = 0
                                    try:
                                        cola3 = int(i['Color Large Impressions'])
                                    except:
                                        cola3 = 0
                                    meter = Meter.objects.create(asset=asset,
                                                                 date_operation=datetime.datetime.today(),
                                                                 bwa4=bwa4,
                                                                 cola4=cola4,
                                                                 bwa3=bwa3,
                                                                 cola3=cola3,
                                                                 bw_total_a4=2*bwa3+bwa4,
                                                                 col_total_a4=2*cola3+cola4)
                elif EDA:
                    with open(file, mode='r') as f:
                        reader = csv.reader(f, delimiter=',')
                        array = list()
                        head_row = None
                        for row in reader:
                            if 'Model' in row:
                                head_row = row
                            if not head_row is None and not head_row == row:
                                data = dict()
                                for i in range(0, len(row)):
                                    data[head_row[i]] = row[i]
                                array.append(data)
                        for i in array:
                            if not i['Model'] == "":
                                p_model = i['Model']
                                try:
                                    p_model = Model.objects.get(model_name=p_model)
                                except Model.DoesNotExist:
                                    p_model = Model.objects.create(model_name=p_model)
                                try:
                                    asset = Asset.objects.get(serialnumber=i['Serial Number'])
                                except Asset.DoesNotExist:
                                    asset = Asset.objects.create(serialnumber=i['Serial Number'], model=p_model,
                                                                 location=i['Location'])
                                try:
                                    meter = Meter.objects.get(asset=asset)
                                except Meter.DoesNotExist:
                                    try:
                                        bwa4 = int(i['A4 Total mono print pages (All-time)'])
                                    except:
                                        bwa4 = 0
                                    try:
                                        cola4 = int(i['A4 Total color print pages (All-time)'])
                                    except:
                                        cola4 = 0
                                    try:
                                        bwa3 = int(i['A3 Total mono print pages (All-time)'])
                                    except:
                                        bwa3 = 0
                                    try:
                                        cola3 = int(i['A3 Total color print pages (All-time)'])
                                    except:
                                        cola3 = 0
                                    meter = Meter.objects.create(asset=asset,
                                                                 date_operation=datetime.datetime.today(),
                                                                 bwa4=bwa4,
                                                                 cola4=cola4,
                                                                 bwa3=bwa3,
                                                                 cola3=cola3,
                                                                 bw_total_a4=2*bwa3+bwa4,
                                                                 col_total_a4=2*cola3+cola4)
                try:
                    os.remove(file)
                except:
                    ...
            for file in arrayFileXML:
                tree = ET.parse(file)
                all_rec = tree.findall('record')
                for x in all_rec:
                    res = dict()
                    for i in x:
                        res[i.tag] = i.text
                    p_model = res['p_model'].replace('"', '')
                    try:
                        p_model = Model.objects.get(model_name=p_model)
                    except Model.DoesNotExist:
                        p_model = Model.objects.create(model_name=p_model)
                    p_serialNumber = res['p_serialNumber'].replace('"', '')
                    try:
                        asset = Asset.objects.get(serialnumber=p_serialNumber)
                    except Asset.DoesNotExist:
                        asset = Asset.objects.create(serialnumber=p_serialNumber, model=p_model,
                                                     location=res['p_location'].replace('"', ''))
                try:
                    os.remove(file)
                except:
                    ...

    def get_email(self, result_bytes):
        msgs = []
        for num in result_bytes[0].split():
            typ, data = self.imap_server.fetch(num, 'RFC822')
            msgs.append(data)
        return msgs

    def close_connection(self):
        self.imap_server.close()
        self.imap_server.logout()

    # TODO add send email class.
