# -*- coding: utf-8 -*-
"""
Created on Thu Jul 12 13:46:48 2018

@author: Kunal Vaid
"""

import os
import sys
import xlrd
import smtplib
from datetime import date as DATE
from datetime import datetime
from email.mime.text import MIMEText
from PIL import Image, ImageFont, ImageDraw
from workshop_content_file_2 import certificate_content
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#Enter Email id and Password around line 130 to activate mail system before execution


generate_date = ""
ref = ""
vertical_orient = 850
horizontal_orient = 500

# Add name, institute and project to certificate
def certificate_generator(ID, name, gender, college, start, end, title, collabrator):
    img = Image.open("templates/Forsk-Letter-2018.jpg")
    image_size = img.size[0]-300

    draw = ImageDraw.Draw(img)

    # Load font
    #font = ImageFont.truetype("fonts/GothamRounded.ttf", 32)
    font = ImageFont.truetype("fonts/Vollkorn-SemiBoldItalic.ttf", 70)
    date_ref_font = ImageFont.truetype("fonts/GothamRoundedMedium.ttf", 36)
    bold_font = ImageFont.truetype("fonts/Vollkorn-BlackItalic.ttf", 70)
    line_spacing = draw.textsize('A', font=font)[1] + 45

    final_content = certificate_content(name, gender, college, font, image_size, start, end, title, collabrator)

    stud_ref = ref + str(ID).zfill(3)

    # Check sizes and if it is possible to abbreviate
    # if not the IDs are added to an error list
    if name == -1:
        return -1
    else:
        # size for REF: (1060, 340)
       # draw.text((1060, 340), stud_ref, font=date_ref_font, fill="rgb(0,0,0)")
        draw.text((2200,300), stud_ref, font=date_ref_font, fill="rgb(5,0,0)")
        # size for Date: (325, 420)
        draw.text((300, 300), generate_date, font=date_ref_font, fill="rgb(0,0,0)")
        


        # Insert text into image template
        #vertical_orient = 750
        global vertical_orient
        global horizontal_orient
        """ Here we have to input custom content manually with the font it uses
        Different parts of statements entered with different fonts
        """
        
        space = ' '
        len1 = int((len(name)+3)/2)
        date=len(start)+len(end)
        date=int(date/2)
        til=len(title)
        
        custom_content("This certificate is awarded to", font, image_size, draw)
        if gender == "M":
            custom_content("Mr.", bold_font, image_size, draw)
        if gender == "F":
            custom_content("Ms.", bold_font, image_size, draw)
        custom_content(name+53*space, bold_font, image_size, draw)
        #draw.line((100,200, 150,300), fill=500)
        #custom_content(" from", font, image_size, draw)
        #custom_content(college, bold_font, image_size, draw)
        custom_content("On Successful Completing of Faculty Training Program on "+30*space, font, image_size, draw)
        #custom_content("Two-Weeks", bold_font, image_size, draw)
        custom_content("\""+title+"\""+30*space, bold_font, image_size, draw)
        custom_content("Conducted By Forsk Coding School, Held Between ", font, image_size, draw)
#put (50-len(date)/2) in next line
        custom_content(start,  font, image_size, draw)
        custom_content("-", font, image_size, draw)
        custom_content(end+".",  font, image_size, draw)

        horizontal_orient = 500
        vertical_orient += 227 + line_spacing
        for para in final_content:
            for line in para:               
                draw.text((horizontal_orient, vertical_orient), line, font=font, fill="rgb(3,24,105)")
                vertical_orient += line_spacing
            vertical_orient += 227

        if not os.path.exists( 'workshop_certificates' ) :
            os.makedirs( 'workshop_certificates' )

        # Save as a PDF inside designated folder
        img.save( './/workshop_certificates//Forsk_SEP_Certificates_'+name+'.pdf', "PDF", resolution=100.0)
        #img.save( 'Forsk_FTP_certificates_'+name+'.pdf', "PDF", resolution=100.0)

        print ('Certificate Generated for '+ name)
        #set back default values of horizontal and vertical orientation
        vertical_orient = 900
        horizontal_orient = 500
       # return 'workshop_certificates\\'+name+'.pdf'
        return './/workshop_certificates//Forsk_SEP_Certificates_'+name+'.pdf'
        
#printing Lines with custom content in the certificate manually
def custom_content(content, font, image_size, draw):
        bold_font = ImageFont.truetype("fonts/Vollkorn-BlackItalic.ttf", 40)
        line_spacing = draw.textsize('A', font=font)[1] + 45
        # split the content by spaces to get words
        words = content.split('\n')
        #print (words)
        i = 0
        global horizontal_orient
        global vertical_orient
        printable = words[i]+" "
        printable_size = font.getsize(printable)[0]
        # Print a word while the width of the word does not exceed the image width
        while i < len(words) and printable_size <= image_size:
            draw.text((horizontal_orient, vertical_orient), printable + " ", font=font, fill="rgb(3,24,105)")
            horizontal_orient = horizontal_orient + printable_size
            
            #print(font.getsize(printable))   
            #print(horizontal_orient , vertical_orient)
            i += 1
            if i < len(words):
                printable = words[i]
                if font == bold_font:
                    printable_size = font.getsize(printable)[0] + 8 #10 is the size of space in bold
                else:
                    printable_size = font.getsize(printable)[0] + 7 #9 is the size of space in semibold
            if printable_size > image_size - horizontal_orient + 190:
               horizontal_orient = 500
               vertical_orient += line_spacing
        

# Email the certificate as an attachment
def email_sender( filename, receiver, stud_name ):
    #Enter email and password used to sending the email
    username = ""
    password = ""
    """ Google blocks SMTP access for apps so you need to enable 
        "allow less secure apps to access" inside your gmail account
        Also, Make sure to enter password for the email you are using
    """
    # sender = username + '@gmail.com'
    stud_name = stud_name.split()[0]

    msg = MIMEMultipart()
    #msg['Subject'] = 'Workshop Certificate'
    msg['Subject'] = 'Forsk Technologies :Certificate'
    msg['From'] = "Forsk Labs"
    msg['Reply-to'] = username 
    msg['To'] = receiver

    # That is what u see if dont have an email reader:
    msg.preamble = 'Multipart massage.\n'
    mail_body = "Dear " + stud_name + ''',\n\nGreetings from Forsk!!!!\n\nThank you for participating in Python Skill Enhancement Program.\n\nPlease find your attached Certificate of participation. Forsk team wishes you all the good luck!.\n\nThe Forsk Coding School by Forsk Technologies is on a mission to Strengthen Data Science, Data Engineering, Python, Machine Learning and Artificial Intelligence ecosystem in India. We work towards the vision of Indian students, working professionals, researchers becoming critical contributors, owners, and shapers of the coming advances in Artificial Intelligence and Machine Learning.\n\nThe Forsk Coding School Jaipur will organize the 4th “Forsk Summer Students Developer Program 2020” in emerging technology trends — Python, Data Engineering, Machine Learning, Deep Learning, Bigdata, DevOps, AI, and Web” — in Jaipur, India between May — July 2020.\n\nWe are happy to announce, that you are eligible for discount of INR 2500/- while enrolling for Forsk Summer Students Developer Program-2020.\n\nPlease find details here about Summer Program:\n1. Registration Link-http://bit.ly/fssdp20\n2. Details-https://medium.com/forsk-labs/forsk-summer-students-developer-program-2020-aaa108bcd89c\n3. Please review us on google-https://g.page/forsklabs/review?rc\n\nWith Regards\nSurendra\nForsk Technologies\nwww.forsk.in\n+91-78519 29944'''.format(stud_name)
    # Body
    part = MIMEText( mail_body )
    msg.attach( part )

    # Attachment
    part = MIMEApplication(open(filename,"rb").read())
    part.add_header('Content-Disposition', 'attachment',
                    filename = os.path.basename(filename))
    msg.attach( part )

    # Login
    server = smtplib.SMTP( 'smtp.gmail.com:587' )
    server.starttls()
    server.login( username, password )

    #print "[INFO]: ",server.verify(msg['To'])
    # Send the email
#    is_valid = validate_email(msg["To"],verify=True)
#    is_valid = True
#    if is_valid:
    server.sendmail( msg['From'], msg['To'], msg.as_string() )
#        server.quit()
    return True
#    else:
#        return False

if __name__ == "__main__":

    error_list, not_sent_list = [], []
    error_count, not_sent_count = 0, 0

    os.chdir(os.path.dirname(os.path.abspath((sys.argv[0]))))

    # Read data from an excel sheet from row 2
    Book = xlrd.open_workbook('workshop_Student_List.xlsx')
    WorkSheet = Book.sheet_by_name('Sheet1')

    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1

        ID = WorkSheet.cell_value( row, 0 )
        name = WorkSheet.cell_value( row, 1 )
        gender = WorkSheet.cell_value( row, 2 )
        college = WorkSheet.cell_value( row, 4 )
        receiver = WorkSheet.cell_value( row, 6 )
        collabrator = WorkSheet.cell_value( row, 8 )
        title = WorkSheet.cell_value( row, 9 )
        duration = WorkSheet.cell_value( row, 10 )
        start_date = WorkSheet.cell_value( row, 11 )
        end_date = WorkSheet.cell_value( row, 12 )
        mail = WorkSheet.cell_value( row, 13 )
        cert_generate_status = WorkSheet.cell_value( row, 14 )
        generate_date = WorkSheet.cell_value( row, 15 )
     
        #generate_date = generation_date_xls
        xls_date_time_obj = ''
        if len(generate_date) > 1:
            xls_date_time_obj = datetime.strptime(generate_date,"%d %B %Y")
            #ref = "FT/HR/EXP/" + xls_date_time_obj.strftime("%B/%Y/").upper()
            #ref = "FT/WS/" + xls_date_time_obj.strftime("%b/%Y/").upper()
            ref = "SEP/" + xls_date_time_obj.strftime("%b/%Y/").upper()
        # duration = WorkSheet.cell_value( row, 15 )
        #duration = "45 hours"





        if cert_generate_status == "YES":
            # Make certificate and check if it was successful
            filename = certificate_generator( int(ID), name, gender, college, start_date, end_date, title, collabrator)
            # Successfully made certificate

            try:
                if filename != -1 and mail=="YES" :

                    sent = email_sender( filename, receiver, name )
                    if sent:
                        print ("Sent to " + name)
                    else:
                        not_sent_list.append( str(ID) )
                        not_sent_count += 1
                        print ("[INFO]: Invalid Email Address")

            except Exception as e:
                # Add to error list
                print ("[ INFO ]: "+str(e))
                error_list.append( str(ID ))
                error_count += 1

    # Print all failed IDs
    print (str(int(not_sent_count)) + " Not Sent- List:" + ','.join(not_sent_list))
    print (str(int(error_count)) + " Errors- List:" + ','.join(error_list))
