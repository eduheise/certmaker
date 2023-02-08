import os
import json
import pandas
from tqdm import tqdm
from PIL import Image, ImageDraw, ImageFont
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders


class CertMaker:
    def __init__(self, template) -> None:
        self.template = template
        self.output_folder = template.replace('templates', 'certificates')
        os.makedirs(self.output_folder, exist_ok=True)
        self._load_info()

    def _load_info(self) -> None:
        self.data = pandas.read_csv(f'{self.template}/data.csv')
        self.cert_image = Image.open(f"{self.template}/template.jpg")
        with open(f'{self.template}/meta.json', 'r') as file:
            self.meta = json.load(file)

    def generate(self) -> None:
        for _, row in tqdm(self.data.iterrows()):
            cert = self.cert_image.copy()
            try:
                for field in self.meta['fields']:
                    cert = self._add_field(cert, row, field)
            except Exception as e:
                print(e)
                print(f"Failed: {row}")
                raise e
            filename = row[self.meta['save_column']]
            filename = filename.replace(' ', '_').upper()
            cert_path = f'{self.output_folder}/{filename}.pdf'
            cert.save(cert_path, dpi=(300, 300))
            if self.meta["send_mail"]:
                self.prepare_delivery(row, cert_path)

    def _add_field(self, cert, row, field: dict):
        draw = ImageDraw.Draw(cert)
        font = ImageFont.truetype(
            f'resources/{field["font-family"]}', field['font-size'])
        if type(field["column"]) is list:
            column_value = [row[f] for f in field['column']]
            message = field['formatter'].format(*column_value)
        else:
            column_value = row[field['column']]
            message = field['formatter'].format(column_value)

        if "max_elements" in field:
            message_elements = message.split(" ")
            m = field['max_elements']

            broken_message = []
            for i in range(0, len(message_elements), m):
                broken_message.append(" ".join(message_elements[i:i + m]))
            message = broken_message
        else:
            message = [message]

        while len(message) < 3:
            message.insert(0, "")
            message.append("")

        for i, m in enumerate(message):
            _, _, w, h = draw.textbbox((0, 0), m, font=font)
            W, H = field['coords'][0], field['coords'][1]
            H += i * field["pad"]
            draw.text((W-w/2, H-h/2), m, font=font,
                      fill=field['font-color'])
        return cert

    def prepare_delivery(self, row, filename):
        send_from = self.meta["mail"]["send_from"]
        subject = self.meta["mail"]["subject"]
        send_to = self.meta["mail"]["send_to"].strip()
        content = self.meta["mail"]["content"]
        parameters = self.meta["mail"]["parameters"]

        server = self.meta["mail"]["server"]
        port = self.meta["mail"]["port"]
        password = self.meta["mail"]["password"]

        formatted_content = content.format(*[row[p] for p in parameters])

        CertMaker.send_mail(
            send_from,
            row[send_to],
            subject,
            formatted_content,
            files=[filename],
            server=server,
            port=port,
            password=password)

    @staticmethod
    def send_mail(send_from, send_to, subject, message, files=[],
                  server="localhost", port=587, password=''):
        """Compose and send email with provided info and attachments.

        Args:
            send_from (str): from name
            send_to (list[str]): to name(s)
            subject (str): message title
            message (str): message body
            files (list[str]): list of file paths to be attached to email
            server (str): mail server host name
            port (int): port number
            username (str): server auth username
            password (str): server auth password
            use_tls (bool): use TLS mode
        """
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject

        msg.attach(MIMEText(message))

        for path in files:
            part = MIMEBase('application', "octet-stream")
            with open(path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            'attachment', filename=Path(path).name)
            msg.attach(part)

        print(send_to)
        smtp = smtplib.SMTP(server, port)
        smtp.starttls()
        smtp.verify(send_to)
        smtp.login(send_from, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
