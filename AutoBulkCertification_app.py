'''
Important configurations:

1. Check - srno
2. Check if proper "Certificate Template" is chosen.
3. Make sure if the text size and pixel positions are ok.
4. Check the email ID.
5. Check the email SUBJECT and CONTENT. Also make sure you attached the
   correct file to the email.
6. Check the XL file That is been added.
7. MOST IMP. after running the program, make sure to delete the entries from the excel file.
   to avoid duplicating the email to the receiver later.
'''

import pandas as pd
import smtplib
from email.message import EmailMessage
import imghdr
from PIL import Image, ImageDraw, ImageFont
import xlrd

#  Variable Initialization
excelFilLoc = "Excel_Data/DemoExcelFile.xlsx"
gmail_id = "YOUR_EMAIL@gmail.com"  # Only Gmail Id is allowed on sender side. No restriction on receiver email ID.
gmail_pd = "YOUR_PASSWORD"
gmail_subject = 'DEMO Certificate from KSHITIJ SANGAR'
gmail_content ="""
Hello <name>, 

This is a Demo Certificate.

Regards,
Kshitij Sangar
"""

template_loc = "images/CertificateTemplate.jpg"
name_fontcolor = (255, 191, 0)
srno_fontcolor = (135, 3, 120)
enduser_fname = 'CERTIFICATE.pdf'

file = pd.ExcelFile(excelFilLoc)  # FinalExcel_1.xlsx

# Email Setup
s = smtplib.SMTP("smtp.gmail.com", 587)
s.starttls()  # Traffic encryption
s.login(gmail_id, gmail_pd)

for sheet in file.sheet_names:
    print("\n\n New Sheet...\n")
    df1 = file.parse(sheet)
    for i in range(len(df1['EMAIL'])):
        # Certificate's file name - to save
        CertificateFileName = str(df1['SRNO'][i]) + "_" + str(df1['NAME'][i].replace(" ","")) + "_" + "2020.pdf"

        # Setting up the Certificate
        cft = Image.open(template_loc)  # Cft = Certificate
        fnt_type_name = ImageFont.truetype('arial.ttf', 50)  #----------------------------
        fnt_type_srno = ImageFont.truetype('arial.ttf', 25)

        draw = ImageDraw.Draw(cft)
        draw.text(xy=(350, 805), text=df1['NAME'][i], fill=name_fontcolor, font=fnt_type_name, stroke_width=1)
        draw.text(xy=(895, 1523), text='Sr.no:' + str(df1['SRNO'][i]), fill=srno_fontcolor, font=fnt_type_srno, stroke_width=1)

        #cft.show()
        cft.save('certificates/' + CertificateFileName)



        msg = EmailMessage()
        msg['Subject'] = gmail_subject
        msg['From'] = gmail_id
        msg['To'] = df1['EMAIL'][i]
        gmail_content = gmail_content.replace("<name>", df1['NAME'][i])
        msg.set_content(gmail_content)

        #Attaching the Poster
        f = open('certificates/' + CertificateFileName, 'rb')
        fdata = f.read()
        #fname = 'images/' + CertificateFileName
        fname = enduser_fname

        file_type = imghdr.what(f.name)
        msg.add_attachment(fdata, maintype='application', subtype='octet-stream', filename=fname)

        s.send_message(msg)
        print("--> ", df1['SRNO'][i], ": ", df1['EMAIL'][i], " : Sent")
s.quit()
print("\n Certificates sent...")
