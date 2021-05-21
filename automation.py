from PyPDF2 import PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas
import reportlab
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.charts.textlabels import _text2Path
from reportlab.lib.pagesizes import A4
import io, mimetypes, os, csv, smtplib
from email.message import EmailMessage
from PIL import ImageFont

BASE_DIR = "C:/Users/flat planet/Desktop/YunusEmre/aja/genPDF"

# Import user defined font to system.
TTFSearchPath = (BASE_DIR)
reportlab.rl_config.TTFSearchPath.append(BASE_DIR)
pdfmetrics.registerFont(TTFont('Itim-Regular', 'Itim-Regular.ttf'))

PAGE_WIDTH = PdfFileReader(open("BASE.pdf", "rb")).getPage(0).mediaBox[2]


# Get list utility function.
# Reads the given csv file in the same directory and converts to python list.
def getList(file):
    queue = []
    with open(file) as csvfile:
        reader = csv.reader(csvfile)
        for name, email in reader:
            queue.append([name, email])
    return queue

# Send mail utility function.
def sendMail(name, mail):
    global mailServer

    message = EmailMessage()
    
    message["From"] = os.environ.get("EMAIL")
    message["To"] = mail
    message["Subject"] = "KMO Ankara Öğrenci"
    body = f"""Sayın {name},\n08.05.2021 tarihli Workshop etkinliğimize katılarak katılım sertifikamızı almaya hak kazandınız. İlgili sertifika ektedir.
    Etkinliğimize katıldığınız için teşekkür ederiz. Diğer etkinliklerimizden haberdar olmak için bizi sosyal medyada takip edin:\n
    Instagram: https://www.instagram.com/kmoankaraogrenci/
    LinkedIn: https://tr.linkedin.com/in/kmo-ankara-%C5%9Fubesi-%C3%B6%C4%9Frenci-komisyonu-970930201\n\n
    KMO Ankara Öğrenci koordinasyon ekibi"""
    message.set_content(body)

    mime_type, _ = mimetypes.guess_type(f"{name}.pdf")
    mime_type, mime_subtype = mime_type.split("/", 1)

    with open(f"{name}.pdf", "rb") as ap:
        message.add_attachment(ap.read(), maintype=mime_type, subtype=mime_subtype, filename=f"{name}.pdf")
    
    mailServer.send_message(message)



# Takes list and loops through the list to generates pdf as the given list's
# 0. indexed value and sends mail to list item's 1. indexed value.
def main(queue):
    returnList = []
    for name, email in queue:
        
        packet = io.BytesIO()
        output = PdfFileWriter()
        existing_pdf = PdfFileReader(open("BASE.pdf", "rb"))
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFillColorRGB(99/256, 87/256, 82/256)
        can.setFont("Itim-Regular", 32)
        font = ImageFont.truetype('Itim-Regular.ttf', 30)
        size = font.getsize(name)
        can.drawString((int(PAGE_WIDTH) - size[0])/2.0 , 340, name.encode("utf-8"))
        can.save()
        packet.seek(0)
        new_pdf = PdfFileReader(packet)

        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)

        outputStream = open(f"{name}.pdf", "wb")
        output.write(outputStream)
        outputStream.close()
        os.rename(f"{name}.pdf", f"files/{name}.pdf")
        sendMail(name, email)
    mailServer.quit()
	

mailServer = smtplib.SMTP_SSL("smtp.gmail.com")
mailServer.set_debuglevel(1)
mailServer.login(os.environ.get("EMAIL"), os.environ.get("EMAIL_PASSWD"))

main(getList("list.csv"))
