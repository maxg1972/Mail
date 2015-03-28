"""
    dbDrive - database query executer
    =================================
    Utilities to facilitate mail sending. This is accomplished with an object that give you methods to send mail

    Requires
    -----
    smtplib and  email modules

    Usage
    -----
    Import the Mail.py file into your project and use the Mail object.

    Authors & Contributors
    ----------------------
        * Massimo Guidi <maxg1972@gmail.com>,

    License
    -------
    This module is free software, released under the terms of the Python
    Software Foundation License version 2, which can be found here:

        http://www.python.org/psf/license/

"""

__author__ = "Massimo Guidi"
__author_email__ = "maxg1972@gmail.com"
__version__ = '1.0'

import smtplib
import getpass
import sys
import os
from  email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

#------------------------#
# Handling error classes #
#------------------------#
class ConnectionError(Exception): pass
class AuthError(Exception): pass
class SendError(Exception): pass

# -------------------- #
# e-mail sending class #
# -------------------- #
class Mail(object):
    # Error codes
    __ERROR_SMTP_CONNECTION=1000
    __ERROR_SMTP_AUTHENTICATION=1001
    
    
    def __init__(self,server='localhost',port=25,user=None,password=None,encryption=None,delivery_notification=False,return_receipt=False):
        """
        Constructor

        @param server: smtp server address (default: localhost)
        @param port: connection port (default: 25)
        @param user: smtp username (default: None)
        @param password: smtp password (default: None)
        @param encryption: encryption method (values: None, 'SSL', 'TLS'; default: None)
        @param delivery_notification: send delivery notification (values: True, False; default: False)
        @param return_receipt: send return receipt (values: True, False; default: False)
        """
        self.__smtp = {'server':server,
                      'port':port,
                      'user':user,
                      'password':password,
                      'encryption':encryption}
        self.__notify = {'delivery_notification':delivery_notification,
                         'return_receipt':return_receipt}
        self.__attacments = None
        
    def set_smtp(self,server,port=25,user=None,password=None,encryption=None):
        """
        Set/change stmp server values.

        @param server: smtp server address
        @param port: connection port (default: 25)
        @param user: smtp username (default: None)
        @param password: smtp password (default: None)
        @param encryption: encryption method (values: None, 'starttls'; default: None)
        """
        self.__smtp['server'] = server
        self.__smtp['port'] = port
        self.__smtp['user'] = user
        self.__smtp['password'] = password
        self.__smtp['encryption'] = encryption
        
    def set_notify(self,delivery_notification=False,return_receipt=False):
        """
        Set/change notification options: 'delivery notification' and 'return receipt'

        @param delivery_notification: send delivery notification (values: True, False; default: False)
        @param return_receipt: send return receipt (values: True, False; default: False)
        """
        self.__notify['delivery_notification'] = delivery_notification
        self.__notify['return_receipt'] = return_receipt
                      
    def set_attachments(self,files):
        """
        Files to be attached

        @param files: list of filenames (with path) to be attached
        """
        self.__attacments = files
        
    def get_parameters(self):
        """
        Return current configuration parameters

        @return: Dictionary with all configuration parameters
        """
        return {'smtp':self.__smtp,
                'notify':self.__notify,
                'attacments':self.__attacments}
    
    def __str__(self):
        """
        Return current configuration parameters as string.

        @return: String with all configuration parameters
        """
        return repr(self.get_parameters())

    def __send_mail(self,send_from, send_to, send_cc, send_bcc, subject, message, message_type):
        """
        Internal method that send mail with current configuration parameters (smtp, notification and attachments)

        @param send_from: sender (e-mail address)
        @param send_to: List of recipients (e-mail addresses)
        @param send_cc: List of recipients (email addresses) as carbon copy
        @param send_bcc: List of recipients (email addresses) as blind carbon copy
        @param subject: Message subject
        @param message: Message body
        @param message_type: Message type (values: 'text' o 'html')
        @return: Sending result (True or False)
        """
        # Message data
        msg = None
        if self.__attacments != None:
            # --- Message with attachments ---
            msg = MIMEMultipart()
            
            # sender and recipients
            msg['From'] = send_from
            msg['To'] = COMMASPACE.join(send_to)

            # CC recipients
            if send_cc:
                msg['Cc'] = COMMASPACE.join(send_cc)

            # sending date (current date)
            msg['Date'] = formatdate(localtime=True)
            
            # message body
            msg['Subject'] = subject
            
            # delivery notification address (sender)
            if self.__notify['delivery_notification']:
                msg['Disposition-Notification-To'] = send_from
                
            # return receipt address (sender)
            if self.__notify['return_receipt']:
                msg['Return-Receipt-To'] = send_from
                
            # Message type
            if message_type == 'html':
                msg.attach(MIMEText(message,'html'))
            else:
                msg.attach(MIMEText(message,'text'))
            
            # Attachemnt files
            for f in self.__attacments:
                part = MIMEBase('application', "octet-stream")
                try:
                    part.set_payload(open(f,"rb").read())
                    Encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
                    msg.attach(part)
                except:
                    pass
        else:
            # --- Message without attachments ---
            
            # Message type
            if message_type == 'html':
                msg = MIMEText(message,'html')
            else:
                msg = MIMEText(message,'text')
            
            # sender and recipients
            msg['From'] = send_from
            msg['To'] = COMMASPACE.join(send_to)

            # CC recipients
            if send_cc:
                msg['Cc'] = COMMASPACE.join(send_cc)

            # sending date (current date)
            msg['Date'] = formatdate(localtime=True)
            
            # message body
            msg['Subject'] = subject
            
            # delivery notification address (sender))
            if self.__notify['delivery_notification']:
                msg['Disposition-Notification-To'] = send_from
                
            # return receipt address (sender)
            if self.__notify['return_receipt']:
                msg['Return-Receipt-To'] = send_from
                
        # open STMP server connection
        try:
            if (self.__smtp['encryption']) and (self.__smtp['encryption'] == "SSL"):
                # active encryption
                smtp = smtplib.SMTP_SSL(self.__smtp['server'],self.__smtp['port'])
            else:
                # noe encryption
                smtp = smtplib.SMTP(self.__smtp['server'],self.__smtp['port'])
        except smtplib.socket.gaierror:
            raise ConnectionError("Server connection error (%s)" % (self.__smtp['server']))

        # active encryption TLS
        if (self.__smtp['encryption']) and (self.__smtp['encryption'] == "TLS"):
            smtp.ehlo_or_helo_if_needed()
            smtp.starttls()

        # execute STMP server login
        if self.__smtp['user']:
            smtp.ehlo_or_helo_if_needed()
            try:
                smtp.login(self.__smtp['user'], self.__smtp['password'])
            except smtplib.SMTPAuthenticationError:
                smtp.close()
                raise AuthError("Invalid username or password (%s)" % (self.__smtp['user']))

        # send e-mail
        try:
            if send_cc:
                send_to += send_cc
            if send_bcc:
                send_to += send_bcc

            smtp.sendmail(send_from, send_to, msg.as_string())
            return True
        except smtplib.something.senderror, errormsg:
            raise SendError("Unable to send e-mail: %s" % (errormsg))
        except smtp.socket.timeout:
            raise ConnectionError("Unable to send e-mail: timeout")
        finally:
            # close SMTP server connection
            smtp.close()
        
    def send_text(self,send_from, send_to, subject, message, send_cc=None, send_bcc=None):
        """
        Send text e-mail.

        @param send_from: sender (e-mail address)
        @param send_to: List of recipients (comma separated e-mail addresses)
        @param subject: Message subject
        @param message: Message body
        @param send_cc: List of CC recipients (comma separated e-mail addresses)
        @param send_bcc: List of BCC recipients (comma separated e-mail addresses)
        @return: Sending result (True or False)
        """
        # Se presenti i destinatari, ne crea una lista
        if send_to:
            send_to = send_to.split(',')
        if send_cc:
            send_cc = send_cc.split(',')
        if send_bcc:
            send_bcc = send_bcc.split(',')
            
        # Send e-mail
        return self.__send_mail(send_from, send_to, send_cc, send_bcc, subject, message, 'text')

    def send_html(self,send_from, send_to, subject, message, send_cc=None, send_bcc=None):
        """
        Send HTML e-mail.

        @param send_from: sender (e-mail address)
        @param send_to: List of recipients (comma separated e-mail addresses)
        @param subject: Message subject
        @param message: Message body
        @param send_cc: List of CC recipients (comma separated e-mail addresses)
        @param send_bcc: List of BCC recipients (comma separated e-mail addresses)
        @return: Sending result (True or False)
        """
        # Se presenti i destinatari, ne crea una lista
        if send_to:
            send_to = send_to.split(',')
        if send_cc:
            send_cc = send_cc.split(',')
        if send_bcc:
            send_bcc = send_bcc.split(',')
            
        # Invio dell' email
        return self.__send_mail(send_from, send_to, send_cc, send_bcc, subject, message, 'html')
