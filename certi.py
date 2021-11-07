import xlrd
import smtplib
import os
import sys
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import datetime



# Add participant name to certificate
def make_certi(name):

    img = Image.open("aa.png")#use only png format templates
    draw = ImageDraw.Draw(img)

    # Load font
    font = ImageFont.truetype("Anton.TTF", 85)

    if name == -1:
        return -1
    else:
        # Insert text into image template
        
        draw.text( (700,700), name, (0,0,0), font=font ) #draw.text((x, y), message, fill=color, font=font)

        if not os.path.exists( 'certificates' ) :
            os.makedirs( 'certificates' )

        # Save as a PDF
        rgb = Image.new('RGB', img.size, (255, 255, 255))  # white background
        rgb.paste(img, mask=img.split()[3])               # paste using alpha channel as mask


        rgb.save( 'certificates/'+str(name)+'.pdf', "PDF", resolution=100.0)
        #rgb.save("WEBD_Certificate_" + str(name) + ".pdf")#local storage
        return 'certificates/'+str(name)+'.pdf'

# Email the certificate as an attachment
def email_certi( filename, receiver, name ):
    username = "iicsinhgad"
    password = "Sinhgad@iic"
    sender = username + '@gmail.com'

    msg = MIMEMultipart()
    msg['Subject'] = 'IPR & Patent Filing Session Certificate'
    msg['From'] = 'IIC,SIT'#username+'@gmail.com'
    msg['Reply-to'] = 'NO-REPLY'
    msg['To'] = receiver

    # That is what u see if dont have an email reader:
    msg.preamble = 'Multipart massage.\n'
    
    # Body
    body = "Hello, {}\nThank you For Attending Interactive Talk Session of Adv.Swapnil Gawande Sir for his webinar on IPR and Patent Filing!\nCheck attachment for your certificate,Hope You Attend Such More Events Organised By IIC,SIT\nRegards Team IIC Mycouncil".format(name)
    part = MIMEText(body)
    msg.attach( part )

    # Attachment
    part = MIMEApplication(open(filename,"rb").read())
    part.add_header('Content-Disposition', 'attachment', filename = os.path.basename(filename))
    msg.attach( part )

    # Login
    server = smtplib.SMTP( 'smtp.gmail.com:587' )
    server.starttls()
    server.login( username, password )

    # Send the email
    server.sendmail( msg['From'], msg['To'], msg.as_string() )

if __name__ == "__main__":
    error_list = []
    error_count = 0

    os.chdir(os.path.dirname(os.path.abspath((sys.argv[0]))))

    # Read data from an excel sheet from row 2
    Book = xlrd.open_workbook('data12.xlsx')
    WorkSheet = Book.sheet_by_name('Form responses 1')
    
    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1
        
        name = WorkSheet.cell_value( row, 2 )
        email = WorkSheet.cell_value( row, 1 )
        
        # Make certificate and check if it was successful
        filename = make_certi(name)
        
        # Successfully made certificate
        if filename != -1:
            email_certi( filename, email, name)

            datetime_object = datetime.datetime.now()

            print("Certificate Sucessfully Sent to ", name," with email ",email, "at time ",datetime_object)

        # Add to error list
        else:
            error_list.append( ID )
            error_count += 1

    # Print all failed IDs
    print(str(error_count)," Errors- List:", ','.join(error_list))


#https://myaccount.google.com/u/2/lesssecureapps?pli=1&rapt=AEjHL4ML87nmzD9B4mzLflwpvoMbY5yEQP9DpZRq9dNNNc0LiQ5SBlU3UcIyPLxcdl-VRhCI0lyGZhBKGSP8EyaxM83XH1fJ5w