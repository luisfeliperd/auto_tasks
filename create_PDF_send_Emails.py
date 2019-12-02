# coding: utf-8

#Source Send Emails - https://www.geeksforgeeks.org/send-mail-attachment-gmail-account-using-python/
#Source Write PDF - https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python Author - David Dehghan

from PyPDF2 import PdfFileWriter, PdfFileReader
from pw_fradd import pw_fradd
import io
import unidecode
import time
import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

# ## Write PDF

def write_pdf(s_path, d_path, name):
    
    #Avoid accents errors
    file_name = unidecode.unidecode(name)
    name = name.split("_")[1]
    
    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Helvetica-Bold", 18)
        
    if len(name) > 30:
        can.setFont("Helvetica-Bold", 17)
        can.drawString(200, 393, name.upper())
        
    elif len(name) <20:
        can.setFont("Helvetica-Bold", 18)
        can.drawString(230, 393, name.upper())
        
    else:
        can.setFont("Helvetica-Bold", 18)
        can.drawString(230, 393, name.upper())
        
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(s_path, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open(d_path+file_name.split(' ')[0]+'.pdf', "wb")
    output.write(outputStream)
    outputStream.close()
    
    return d_path+file_name.split(' ')[0]+'.pdf'

# ## Send Emails

info_dict = {
    "fromaddr":"",
    "toaddr":"",
    "subject":"Email de Teste - Emissão de Certificados",
    "body":"",
    "filename":"certificado.pdf",
    "attachment":"certificado.pdf",
}


def send_email(info_dict):
    # instance of MIMEMultipart 
    msg = MIMEMultipart() 

    # storing the senders email address   
    msg['From'] = info_dict["fromaddr"]

    # storing the receivers email address  
    msg['To'] = info_dict["toaddr"]

    # storing the subject  
    msg['Subject'] = info_dict["subject"]

    # attach the body with the msg instance 
    msg.attach(MIMEText(info_dict["body"], 'plain')) 

    # open the file to be sent  
    attachment = open(info_dict["attachment"], "rb")

    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form 
    p.set_payload((attachment).read())

    # encode into base64 
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % info_dict["filename"])

    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 

    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 

    # start TLS for security 
    s.starttls() 

    # Authentication 
    s.login(info_dict["fromaddr"], pw_fradd)

    # Converts the Multipart msg into a string 
    text = msg.as_string() 

    # sending the mail 
    s.sendmail(info_dict["fromaddr"], info_dict["toaddr"], text) 

    # terminating the session 
    s.quit() 

# ## Iterate Names

s_path = "certificado.pdf"
s_path2 = "certificado_2.pdf"
s_path3 = "certificado_3.pdf"
d_path = ""

counter = 0

#In case of same first names
count_name = 0

with open('nomes_emails.csv', encoding='utf-8') as f:
    csv_reader = csv.reader(f)
    for row in csv_reader:
        count_name+=1
        
        if len(row[0]) > 30:
            file_p = write_pdf(s_path2, d_path, str(count_name)+"_"+row[0])
        
        elif len(row[0]) < 20:
            file_p = write_pdf(s_path3, d_path, str(count_name)+"_"+row[0])
            
        else:
            file_p = write_pdf(s_path, d_path, str(count_name)+"_"+row[0])
        
        info_dict["toaddr"] = row[1]
        info_dict["body"] = ""
        info_dict["filename"] = file_p.split("\\")[-1]
        info_dict["attachment"] = file_p
        
        try:
            send_email(info_dict)
            print("Email para {} ({}) funcionou.".format(row[0], row[1]))
        except:
            print("Email para {} ({}) não funcionou.".format(row[0], row[1]))
            continue
        
        counter+=1
            
        if counter == 10:
            counter = 0
            print("10 Emails Enviados - Esperando 5 Minutos")
            time.sleep(300)

