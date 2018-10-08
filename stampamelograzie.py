#!/usr/bin/python

'''
STAMPAMELOGRAZIE
ver. 1.1 - ZIRCONET 14-03-2018
Programma che controlla le email ricevute nella casella definita 
e stampa automaticamente gli allegati delle email ricevute,
non appena la stampante di rete risulta in linea.
Invia email risposta.
'''

import sys
import email, imaplib, os, smtplib
import urllib
import requests
import subprocess
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText


ORG_EMAIL   = "@gmail.com"                      ## dominio email
MY_EMAIL  = "-------------" + ORG_EMAIL         ## account email
MY_PWD    = "-------------"                     ## pwd email
IMAP_SERVER = "imap.gmail.com"          
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 993
IMAP_PORT   = 587
detach_dir = "/.stampe"                         ## directory tmp stampe
IP_PRINTER = '-------------'                    ## IP stampante


def read_email_from_gmail():
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(MY_EMAIL,MY_PWD)
        mail.select('inbox')
        
        type, data = mail.search(None, 'ALL')
        mail_ids = data[0]

        id_list = mail_ids.split()
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        status, response = mail.search(None, '(UNSEEN)')
        
        id_list = response[0].split()
        
        for e_id in id_list:
            e_id = e_id.decode('utf-8')
            typ, response = mail.fetch(e_id, '(RFC822)')
            email_message = email.message_from_string(response[0][1])
            email_subject = email_message['Subject']
	    email_date = email_message['Date']
            email_from = email.utils.parseaddr(email_message['From'])
            email_response = str(email_from[1])
            print 'From : ' + str(email_from[0])
            print 'email : ' + str(email_response)
	    print 'Data: ' + str(email_date)
            print 'Subject : ' + str(email_subject) + '\n'

            if email_message.get_content_maintype() != 'multipart':
                continue
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

		data = email_date.replace(' ', '_')
		data = data.replace (',','')
		data = data.replace (':','_')
		data = data.replace ('+','')
		print 'data: ' + str(data)

		filename = part.get_filename()

		stato = '--'
		stato = ping()
	        rispondi_ricevimento(email_response, filename, stato)

                filename = data + '_' + part.get_filename()
                att_path = os.path.join(detach_dir, filename)

                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

                print 'Salvato : ' + filename + '\n'

		messaggio = 'Ricevuto: ' + filename

        for e_id in id_list:
            mail.store(e_id, '+FLAGS', '\Seen')

        mail.close()
        mail.logout()


def rispondi_ricevimento(email_response, filename, stato):
    	gmail_user = MY_EMAIL
   	gmail_pwd = MY_PWD 
    	TO = email_response
    	SUBJECT = "Ricevuta stampa (stampante " + stato +")"
    	TEXT = "Conferma allegato in lavorazione --->  " + filename
    	server = smtplib.SMTP('smtp.gmail.com', 587)
    	server.ehlo()
    	server.starttls()
    	server.login(gmail_user, gmail_pwd)
    	BODY = '\r\n'.join(['To: %s' % TO,
            	'From: %s' % gmail_user,
            	'Subject: %s' % SUBJECT,
            	'', TEXT])

    	server.sendmail(gmail_user, [TO], BODY)
    	print 'email spedita a: ' + email_response

def ping():
        var = os.popen("ping "+ IP_PRINTER + " -c 1").read()
        l = var.index('received')
        l = l-2
	if var[l] == '1':
            stato = 'ON'
	else:
            stato = 'OFF'	
	return stato

def stampante():
        var = os.popen("ping "+ IP_PRINTER + " -c 1").read()
        l = var.index('received')
        l = l-2
	stampa = '--'

        if var[l] == '1':
            print 'stampante - raggiungibile\n'

	    path = detach_dir
	    dirs = os.listdir(path)

	    for dr in dirs:
		nomefile = path + '/' + dr
		print nomefile

		try:
                    subprocess.check_call(["lp", nomefile])
		except OSError:
		    pass

		try:
		    os.remove(nomefile)
		except OSError:
		    pass

		messaggio = 'Stampato: ' + nomefile
                print messaggio






try:
        'Controllo la connessione ad Internet'
        url = "http://www.google.it"
        r = requests.get(url)
        messaggio = 'Connessione... OK\n'
	print messaggio
except:
	e = str(sys.exc_info()[0])
	messaggio = 'Errore di connessione a Internet: ' + e
	print messaggio	

try:
	read_email_from_gmail()
except:
	e = str(sys.exc_info()[0])
	messaggio = 'Errore lettura account!!!  ' + e
	print messaggio

try:
	stampante()
except:
	e = str(sys.exc_info()[0])
	messaggio = 'Errore stampa!!!  ' + e
	print messaggio

