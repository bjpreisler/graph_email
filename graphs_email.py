#Define variable for current date and import lib, needed before custom params

import time
import datetime
day = time.strftime("%m/%d/%Y")
current = datetime.datetime.now()
mnth = current.strftime('%B')

#########################
### CUSTOM PARAMETERS ###
from_email = 'qgi@focus.org'
cc = ['brian.preisler@focus.org']
reg_dict = {}

reg_dict['Brian Preisler'] = ['brian.preisler@focus.org', 'Great North%',575]
#all other recipient emails

to_emails = [to_email] + cc


#########################

"""
    Import all required libraries
"""
import MySQLdb
import xlsxwriter
import smtplib
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


for i in reg_dict:
    strg = reg_dict[i][1]
    subject = 'FYI: Update on %s DM Goal' %strg
    
    """
        Create Acts database connection ()
    """
    db=MySQLdb.connect( host="acts247.focus.org",
                        user="python",
                        passwd="",
                        db="acts247",
                        charset='utf8',
                        use_unicode=True)
    

    """
        Create cursor which will allow use to execute a query
    """
    c=db.cursor()
    c2=db.cursor()

    """
        Actual SQL queries to run
    """
    c.execute("""     SELECT  C.name as Campus,
        count(distinct case when DS.discipleship_type_id in (4,5) THEN DS.user_id END) as 'Male Ds'	
        FROM discipleship_statuses as DS
        LEFT JOIN users as U on U.id = DS.user_id
        LEFT JOIN campuses as C on C.id = U.campus_id
        left join regions as R on R.id = C.region_id
        WHERE (DS.end_date IS NULL OR DS.end_date >= DATE_SUB(CURRENT_DATE, INTERVAL DAYOFMONTH(CURRENT_DATE)-1 DAY))
        AND DS.start_date <= curdate()
        AND C.region_id NOT IN (10,26,44)
        and R.name LIKE %s
        AND (U.user_role_type_id = 3  OR (C.region_id = 45 and U.user_role_type_id IN (4,6) and U.is_affiliate =1))
                    """, (reg_dict[i][1]))

    c2.execute("""  SELECT  C.name as Campus,
        count(distinct case when DS.discipleship_type_id in (4,5) THEN DS.user_id END) as 'Male Ds'
        FROM discipleship_statuses as DS
        LEFT JOIN users as U on U.id = DS.user_id
        LEFT JOIN campuses as C on C.id = U.campus_id
        left join regions as R on R.id = C.region_id
        WHERE (DS.end_date IS NULL OR DS.end_date >= '2016-09-01')			
        AND DS.start_date <= '2016-09-31'
        AND C.region_id NOT IN (10,26,44)
        and R.name like %s
        AND (U.user_role_type_id = 3  OR (C.region_id = 45 and U.user_role_type_id IN (4,6) and U.is_affiliate =1))
                    """, (reg_dict[i][1]))

    """
        Fetch entire row from cursor and store into result object
    """
    result = c.fetchall()
    result2 = c2.fetchall()
    current_DMs = result[0][1]
    sept = result2[0][1]
    goal_17 = reg_dict[i][2]
    pergoal = 100*(goal_17 - current_DMs)/goal_17

    """
        Close database connection
    """
    db.close()

    """
        Open passwords file which holds the QGI sendgrid password
    """
    file = open("C:\\Users\\brian.preisler\\Desktop\\pythontxtfile.txt", "r")

    """
        Extract contents of file
    """
    Mypassword = file.read()

    """
        Close the password file to prevent corruption and locking
    """
    file.close()


    #graph plot

    import plotly.plotly as py
    import plotly.graph_objs as go
    import plotly
    plotly.tools.set_credentials_file(username='bjpreisler', api_key='w7hhb1ym9y')

    trace1 = go.Bar(
        x=['Sept DMs','Current Month DMs', 'April Goal DMs'],
        y=[sept,current_DMs, goal_17],
        marker=dict(
            color=['rgb(55, 83, 109)', 'rgb(55, 83, 109)',
               'rgb(185,211,238)']),
    )
    
    x = ['Sept DMs','Current DMs', 'April Goal DMs']
    y = [sept,current_DMs,goal_17]
    data = [trace1]
    layout = go.Layout(
            annotations=[
            dict(x=xi,y=yi,
                 text=str(yi),
                 xanchor='center',
                 yanchor='bottom',
                 showarrow=False,
            ) for xi, yi in zip(x, y)],   

        title='%s DM Progress' %strg,
        xaxis=dict(
            tickfont=dict(
                size=14,
                color='rgb(107, 107, 107)'
            )
        ),
        yaxis=dict(
            title='DM',
            titlefont=dict(
                size=16,
                color='rgb(107, 107, 107)'
            ),
            tickfont=dict(
                size=14,
                color='rgb(107, 107, 107)'
            )
        ),
        legend=dict(
            x=0,
            y=1.0,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        bargap=0.1,
        bargroupgap=0.3
    )


    fig = go.Figure(data=data, layout=layout)

    py.image.save_as(fig, filename='DM.png')

    from IPython.display import Image
    Image('DM.png')



    """
        Function to send email
    """
    def send_mail(send_from, send_to, subject, text, files=None,
                              data_attachments=None, images=None):

        """
            Email sending parameters
        """
        server = 'smtp.office365.com'
        port = 587
        tls = True
        username = 'brian.preisler@focus.org'
        password = Mypassword

        COMMASPACE = ', '

        if files is None:
            files = []

        if images is None:
            images = []

        if data_attachments is None:
            data_attachments = []

        msg = MIMEMultipart('related')
        msg2 = MIMEMultipart('related')
        msg['From'] = send_from
        msg['To'] = send_to if isinstance(send_to, basestring) else COMMASPACE.join(send_to)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject



        # Encapsulate the plain and HTML versions of the message body in an
        # 'alternative' part, so message agents can decide which they want to display.
        msgAlternative = MIMEMultipart('alternative')
        msg2Alternative = MIMEMultipart('alternative')
        msg.attach(msgAlternative)
        msg2.attach(msg2Alternative)

        msgText = MIMEText('This is the alternative plain text message.')
        msgAlternative.attach(msgText)

        # We reference the image in the IMG SRC attribute by the ID we give it below
        msgText = MIMEText('<img src="cid:image1">', 'html')
        msgAlternative.attach(msgText)

        # This example assumes the image is in the current directory
        fp = open('DM.png', 'rb')
        fp2 = open('BSP.png', 'rb')
        msgImage = MIMEImage(fp.read())
        msgImage2 = MIMEImage(fp2.read())
        fp.close()
        fp2.close()

        # Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<image1>')
        msgImage2.add_header('Content-ID', '<image1>')
        msg.attach(msgImage)
        msg2.attach(msgImage2)

        ##################################

        for f in files:
            part = MIMEBase('application', "octet-stream")
            part.set_payload( open(f,"rb").read() )
            Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
            msg.attach(part)

        for f in data_attachments:
            part = MIMEBase('application', "octet-stream")
            part.set_payload( f['data'] )
            Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % f['filename'])
            msg.attach(part)

        for (n, i) in enumerate(images):
            fp = open(i, 'rb')
            msgImage = MIMEImage(fp.read())
            fp.close()
            msgImage.add_header('Content-ID', '<image{0}>'.format(str(n+1)))
            msg.attach(msgImage)
            msg2.attach(msgImage2)

        smtp = smtplib.SMTP(server, int(port))
        if tls:
            smtp.starttls()

        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.close()



    send_mail(from_email,reg_dict[i][0], subject, text)