#Define variable for current date and import lib, needed before custom params

import time
import datetime
import pymysql
import xlsxwriter
import smtplib
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import Encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.Utils import formatdate
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt
from PIL import Image

day = time.strftime("%m/%d/%Y")
current = datetime.datetime.now()
mnth = current.strftime('%B')

#########################
### CUSTOM PARAMETERS ###
from_email = ''
cc = []
reg_dict = {}
text = ""

to_send_list_full = pd.read_csv("C:\\Users\\brian.preisler\\Dropbox\\Growth\\Data Analysis\\Python Projects\\DM Chart Emails\\to_send_list.csv")

to_send_list_test = pd.read_csv("C:\\Users\\brian.preisler\\Dropbox\\Growth\\Data Analysis\\Python Projects\\DM Chart Emails\\to_send_list_test.csv")

to_send_list_full_test = pd.read_csv("C:\\Users\\brian.preisler\\Dropbox\\Growth\\Data Analysis\\Python Projects\\DM Chart Emails\\to_send_list_full_test.csv")

campus_goals = pd.read_csv("C:\\Users\\brian.preisler\\Dropbox\\Growth\\Data Analysis\\Python Projects\\DM Chart Emails\\campus_goals.csv")

campus_goals = campus_goals.set_index('campus', inplace = False)

to_send_list = to_send_list_full

print to_send_list

for index,row in to_send_list.iterrows():
    email = row['email']
    area = row['area']
    area_goal = row['goal']

    subject = 'FYI: Update on %s DM Goal' %area

    """
    Create Acts database connection ()
    """
    db=pymysql.connect( host="",
                        user="",
                        passwd="",
                        db="",
                        charset='utf8',
                        use_unicode=True)
    

    """
        Create cursor which will allow use to execute a query
    """

    
    
    """
        Actual SQL queries to run
    """
    c=db.cursor()
    
    c.execute("""     SELECT  C.name as Campus,
        count(distinct case when DS.discipleship_type_id in (4,5) THEN DS.user_id END) as 'Ds',
        sub.Sept
        FROM discipleship_statuses as DS
        LEFT JOIN users as U on U.id = DS.user_id
        LEFT JOIN campuses as C on C.id = U.campus_id
        left join regions as R on R.id = C.region_id
        join 
            (SELECT  C.name as Campus, C.id as id,
            count(distinct case when DS.discipleship_type_id in (4,5) THEN DS.user_id END) as Sept
            FROM discipleship_statuses as DS
            LEFT JOIN users as U on U.id = DS.user_id
            LEFT JOIN campuses as C on C.id = U.campus_id
            left join regions as R on R.id = C.region_id
            WHERE (DS.end_date IS NULL OR DS.end_date >= '2017-09-01')
            AND DS.start_date <= '2017-09-31'
            AND C.region_id NOT IN (10,26,44)
            and R.name like %s
            AND (U.user_role_type_id = 3  OR (C.region_id = 45 and U.user_role_type_id IN (4,6) and U.is_affiliate =1))
            group by C.name) as sub on sub.id = C.id      
        
        WHERE (DS.end_date IS NULL OR DS.end_date >= DATE_SUB(CURRENT_DATE, INTERVAL DAYOFMONTH(CURRENT_DATE)-1 DAY))
        AND DS.start_date <= curdate()
        AND C.region_id NOT IN (10,26,44)
        and R.name LIKE %s
        AND (U.user_role_type_id = 3  OR (C.region_id = 45 and U.user_role_type_id IN (4,6) and U.is_affiliate =1))
        group by C.name
                    """, (area,area))


    """
        Fetch entire row from cursor and store into result object
    """
    result = c.fetchall()
    c.close()


    
    both = {}
    
  
    for a in result:
        name = a[0]
        current_DMs = int(a[1])
        sept = int(a[2])
        
        goal = campus_goals.loc[name]['goal']
        

        """
            Open passwords file which holds the QGI sendgrid password
        """
        file = open("", "r")

        """
            Extract contents of file
        """
        Mypassword = file.read()

        """
            Close the password file to prevent corruption and locking
        """
        file.close()


        #graph plot
        sns.set_style("whitegrid")
        
        x = ['Sept DMs', 'Current DMs', 'April Goal DMs']
        y = [sept, current_DMs, goal]
        
        
        current_palette = sns.color_palette("Blues")
        sns.set_palette(current_palette)
        graph = sns.barplot(x, y)
        graph.grid(False)
        graph.set_title(name + ' DM Growth', size = 16)
        
        for p in graph.patches:
            height = p.get_height()
            graph.text(p.get_x()+p.get_width()/2.,
            height+ 0.05,
            '{:1.2f}'.format(height),
            ha="center") 
        
        
        plt.savefig(a[0] + '.png')
        #plt.show()

    
    """
        Function to send email
    """
        # Send an HTML email with an embedded image and a plain text message for
    # email clients that don't want to display the HTML.

    from email.MIMEMultipart import MIMEMultipart
    from email.MIMEText import MIMEText
    from email.MIMEImage import MIMEImage

    # Define these once; use them twice!
    strFrom = ''
    strTo = ''

    # Create the root message and fill in the from, to, and subject headers
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    msgRoot['To'] = strTo
    msgRoot.preamble = 'An Update on DM Growth'

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    msgText = MIMEText('This is the alternative plain text message.')
    msgAlternative.attach(msgText)
    
    emailtext = ''
    counter = 0
    imgcounter = 1
    
    for b in result:
        
        pic = str(result[counter][0])+'.png'
        #pic2 = str(result[1][0])+'.png'
        
        fp = open(pic, 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()
        image = '<image' + str(imgcounter) + '>'
        # Define the image's ID as referenced above
        msgImage.add_header('Content-ID', image)
        msgRoot.attach(msgImage)
        counter += 1
        emailtext += '<br><img src="cid:image' + str(imgcounter) + '"><br>'
        imgcounter += 1
        
    
    # We reference the image in the IMG SRC attribute by the ID we give it below
    msgText = MIMEText(' %s   \
                       <p><em><sup class="versenum">&nbsp; \
                       </sup>I planted, Apollos watered, but God gave the growth. 1 Cor 3:6</em></p>' %emailtext, 'html')
    msgAlternative.attach(msgText)

    # Send the email (this example assumes SMTP authentication is required)
    import smtplib
    smtp = smtplib.SMTP()
    smtp.connect('smtp.office365.com', 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    smtp.login('brian.preisler@focus.org', Mypassword)
    smtp.ehlo()
    smtp.sendmail(strFrom, strTo, msgRoot.as_string())
    smtp.quit()
    print "Email sent to " + area + " at " + email

db.close()

print "Script Complete"
